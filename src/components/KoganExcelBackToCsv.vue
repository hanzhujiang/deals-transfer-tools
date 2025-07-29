<template>
  <div>
    <h2>合并 manifest.csv 与 Excel 生成新 CSV</h2>

    <div>
      <label>上传 manifest.csv：</label>
      <input type="file" accept=".csv" @change="handleSampleUpload" />
    </div>

    <div>
      <label>Excel 中 sheet 名称：</label>
      <input v-model="sheetName" placeholder="例如：发货记录" />
    </div>

    <div>
      <label>上传 Excel 文件（包含“订单号”、“发货单号”、“CARRIER”）：</label>
      <input type="file" accept=".xlsx, .xls" @change="handleExcelUpload" />
    </div>

    <button @click="generateAndExport" :disabled="!sampleData.length || !excelMap.size">
      生成并导出 CSV
    </button>
  </div>
</template>

<script setup>
import { ref } from 'vue'
import * as XLSX from 'xlsx'
import { kogan_couriers_map } from "../const/koganConstants";

const sheetName = ref('')
const sampleData = ref([])
const excelMap = ref(new Map())

function handleSampleUpload(e) {
  const file = e.target.files[0]
  if (!file) return

  const reader = new FileReader()
  reader.onload = () => {
    const workbook = XLSX.read(reader.result, { type: 'binary' })
    const firstSheet = workbook.Sheets[workbook.SheetNames[0]]
    sampleData.value = XLSX.utils.sheet_to_json(firstSheet, { defval: '' })
  }
  reader.readAsBinaryString(file)
}

function handleExcelUpload(e) {
  const file = e.target.files[0]
  if (!file) return

  const reader = new FileReader()
  reader.onload = () => {
    const workbook = XLSX.read(reader.result, { type: 'binary' })
    const targetSheet = workbook.Sheets[sheetName.value.trim()]
    if (!targetSheet) {
      alert(`未找到 Sheet: ${sheetName.value}`)
      return
    }

    const rows = XLSX.utils.sheet_to_json(targetSheet, { header: 1, defval: '' })
    if (rows.length < 3) return

    const headers = rows[2]
    const dataRows = rows.slice(3)

    const formatted = dataRows.map(row => {
      const obj = {}
      headers.forEach((key, index) => {
        obj[key] = row[index] ?? ''
      })
      return obj
    })

    excelMap.value.clear()
    formatted.forEach(row => {
      const orderId = row['订单号']?.toString().trim()

      const trackingId = row['发货单号']?.toString().trim()
      const trackingCarrier = row['CARRIER']?.toString().trim()
      if (orderId && trackingId) {
        excelMap.value.set(orderId, {
          trackingCode: row['发货单号']?.toString() || '',
          carrier: trackingCarrier || ''
        })
      }
    })
  }
  reader.readAsBinaryString(file)
}

function generateAndExport() {
  const today = new Date()
  const dispatchDate = `${String(today.getDate()).padStart(2, '0')}/${String(today.getMonth() + 1).padStart(2, '0')}/${today.getFullYear()}`

  const output = sampleData.value.map(row => {
    const orderId = row['OrderID']?.toString().trim()
    const match = excelMap.value.get(orderId)

    return {
      CONNOTE: match?.trackingCode || '',
      ITEM: row['ProductCode'] || '',
      SERIAL_NUMBER: '',
      DISPATCH_DATE: dispatchDate,
      ORDER_ID: row['OrderID'] || '',
      QUANTITY: row['Quantity'] || '',
      CARRIER:  kogan_couriers_map[match?.carrier] || match?.carrier,
      WAREHOUSE: row['OriginWarehouse'] || ''
    }
  })

  const ws = XLSX.utils.json_to_sheet(output)
  const wb = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(wb, ws, 'Manifest CSV')
  XLSX.writeFile(wb, 'manifest_output.csv')
}
</script>

<style scoped>
input {
  margin: 8px 0;
}

button {
  margin-top: 10px;
}
</style>
