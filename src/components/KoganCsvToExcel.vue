<template>
    <div>
        <h2>生成发货 CSV</h2>

        <div>
            <label>上传 sample CSV：</label>
            <input type="file" accept=".csv" @change="handleSampleCsv" />
        </div>

        <div>
            <label>输入 Excel 的 Sheet 名：</label>
            <input v-model="sheetName" placeholder="例如：发货记录表" />
        </div>

        <div>
            <label>上传 Excel 文件（含订单号、发货单号、CARRIER）：</label>
            <input type="file" accept=".xlsx" @change="handleExcelUpload" />
        </div>

        <button @click="generateCsv" :disabled="!sampleData.length || !excelMap.size">
            生成并导出 CSV
        </button>
    </div>
</template>

<script setup>
import { ref } from 'vue'
import * as XLSX from 'xlsx'
import Papa from 'papaparse'
import { saveAs } from 'file-saver'

const sheetName = ref('')
const sampleData = ref([])
const excelMap = ref(new Map())

function handleSampleCsv(e) {
    const file = e.target.files[0]
    if (!file) return

    Papa.parse(file, {
        header: true,
        skipEmptyLines: true,
        complete(results) {
            sampleData.value = results.data
        }
    })
}

function handleExcelUpload(e) {
    const file = e.target.files[0]
    if (!file) return

    const reader = new FileReader()
    reader.onload = () => {
        const wb = XLSX.read(reader.result, { type: 'binary' })
        const sheet = wb.Sheets[sheetName.value.trim()]
        if (!sheet) {
            alert(`未找到 Sheet: ${sheetName.value}`)
            return
        }

        const rawRows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' })
        const headers = rawRows[2]
        const rows = rawRows.slice(3)

        const structured = rows.map(row => {
            const obj = {}
            headers.forEach((key, i) => {
                obj[key] = row[i] ?? ''
            })
            return obj
        })

        excelMap.value.clear()
        structured.forEach(row => {
            const orderId = row['订单号']?.toString().trim()
            if (orderId) {
                excelMap.value.set(orderId, {
                    trackingCode: row['发货单号'] || '',
                    carrier: row['CARRIER'] || ''
                })
            }
        })
    }
    reader.readAsBinaryString(file)
}

function generateCsv() {
    const today = new Date()
    const dispatchDate = `${String(today.getDate()).padStart(2, '0')}/${String(today.getMonth() + 1).padStart(2, '0')}/${today.getFullYear()}`

    const output = sampleData.value.map(row => {
        const orderId = row['OrderID']?.toString().trim()
        const productCode = row['ProductCode'] || ''
        const match = excelMap.value.get(orderId)

        return {
            CONNOTE: match?.trackingCode || '',
            ITEM: productCode.replace(/^AXS-/, ''),
            SERIAL_NUMBER: '',
            DISPATCH_DATE: dispatchDate,
            ORDER_ID: orderId,
            QUANTITY: row['Quantity'] || '',
            CARRIER: match?.carrier || '',
            WAREHOUSE: row['OriginWarehouse'] || ''
        }
    })

    const csv = Papa.unparse(output)
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' })
    saveAs(blob, 'manifest_output.csv')
}
</script>

<style scoped>
input {
    margin: 6px 0;
}
button {
    margin-top: 12px;
}
</style>
