<template>
  <div>
    <h2>MyDeal CSV to Excel 日控表</h2>

    <!-- 上传 CSV 文件 -->
    <label>
      MyDeal导出的csv
      <input type="file" accept=".csv" @change="handleFileUpload" />
      <br /><br />
    </label>


    <label>
      日控表的 Sheet 名称：
      <input v-model="markingSheetName" placeholder="" />
    </label>
    <br /><br />

    <!-- 上传 标记 Excel 文件 -->
    <label>
      上传日控表
      <input type="file" accept=".xlsx, .xls" @change="handleMarkingUpload" />
      <br /><br />
    </label>


    <!-- 导出按钮 -->
    <button @click="exportToExcel" :disabled="!convertedData.length">导出 Excel</button>

    <!-- 显示表格 -->
    <table v-if="convertedData.length" border="1" style="margin-top: 20px; border-collapse: collapse;">
      <thead>
        <tr>
          <th v-for="header in tableHeaders" :key="header">{{ header }}</th>
        </tr>
      </thead>
      <tbody>
        <tr v-for="(row, rowIndex) in convertedData" :key="rowIndex"
          :style="row.markedForExclusion ? 'text-decoration: line-through; color: grey;' : ''">
          <td v-for="header in tableHeaders" :key="header">{{ row[header] }}</td>
        </tr>
      </tbody>
    </table>
  </div>
</template>

<script setup>
import { ref } from 'vue'
import * as XLSX from 'xlsx'

const convertedData = ref([])
const tableHeaders = ref([])
const markedOrderNumbers = ref(new Set())
const markingSheetName = ref('物控日报表 (2025年7月)') // 用户输入的 sheet 名

// 上传 CSV 文件处理
function handleFileUpload(e) {
  const file = e.target.files[0]
  if (!file) return

  const reader = new FileReader()
  reader.onload = () => {
    const workbook = XLSX.read(reader.result, { type: 'binary' })
    const sheetName = workbook.SheetNames[0]
    const rawData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName])

    convertedData.value = rawData.map(row => {
      const address = [
        row['Address'] || '',
        row['Address 2'] || '',
        row['Suburb'] || '',
        row['State'] || '',
        row['Postcode'] || ''
      ].filter(Boolean).join(', ')

      const orderNo = row['Order No'] || ''
      const isMarked = markedOrderNumbers.value.has(String(orderNo).trim())

      return {
        '销售平台': 'MyDeal',
        '日期': row['Purchased Date'] || '',
        '订单号': orderNo,
        'SKU': row['SKU'] || '',
        '名称': '',
        '数量': row['Quantity'] || '',
        '价格': row['Total Sale Price'] || '',
        '运费': row['Total Shipping'] || '',
        '地址': address,
        '联系人': row['Name'] || '',
        '电话': row['Phone'] || '',
        markedForExclusion: isMarked
      }
    })

    tableHeaders.value = Object.keys(convertedData.value[0] || {}).filter(h => h !== 'markedForExclusion')
  }

  reader.readAsBinaryString(file)
}

// 上传用于标记的 Excel 文件
function handleMarkingUpload(e) {
  const file = e.target.files[0]
  if (!file) return

  const reader = new FileReader()
  reader.onload = () => {
    const workbook = XLSX.read(reader.result, { type: 'binary' })

    const sheetName = markingSheetName.value.trim()
    if (!sheetName) {
      alert('请先填写 Sheet 名称')
      return
    }

    const sheet = workbook.Sheets[sheetName]
    if (!sheet) {
      alert(`未找到名称为 "${sheetName}" 的 Sheet, 请检查拼写`)
      return
    }


    const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 })

    const keyRowIndex = 2
    const keys = jsonData[keyRowIndex] || []

    const orderNoIndex = keys.indexOf('订单号')
    if (orderNoIndex === -1) {
      alert('标记用 Excel 文件的第三行没有找到“订单号”字段')
      return
    }

    // 提取订单号值
    for (let i = keyRowIndex + 1; i < jsonData.length; i++) {
      const row = jsonData[i]
      if (row[orderNoIndex]) {
        markedOrderNumbers.value.add(String(row[orderNoIndex]).trim())
      }
    }

    // 重新标记 convertedData
    convertedData.value = convertedData.value.map(item => {
      const isMarked = markedOrderNumbers.value.has(String(item['订单号']).trim())
      return { ...item, markedForExclusion: isMarked }
    })
  }

  reader.readAsBinaryString(file)
}

// 导出 Excel（过滤删除线记录）
function exportToExcel() {
  const filteredData = convertedData.value.filter(row => !row.markedForExclusion)
  const worksheet = XLSX.utils.json_to_sheet(filteredData)
  const workbook = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1')
  XLSX.writeFile(workbook, 'mydeal_orders.xlsx')
}
</script>
