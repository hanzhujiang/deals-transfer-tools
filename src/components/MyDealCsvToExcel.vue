<template>
  <div>
    <h2>MyDeal CSV 转 Excel</h2>

    <!-- 上传 CSV 文件 -->
    <label>
      MyDeal导出的csv
      <input type="file" accept=".csv" @change="handleFileUpload" />
      <br /><br />
    </label>

    <label>
      MyDeal导出的csv,for Australia Post
      <input type="file" accept=".csv" @change="handleFileUploadAuPost" />
      <br /><br />
    </label>


    <!-- 上传 标记 Excel 文件 -->
    <label>
      上传日控表
      <input type="file" accept=".xlsx, .xls" @change="handleMarkingUpload" />
      <br /><br />
    </label>


    <!-- 导出按钮 -->
    <button @click="exportToExcel('SMCC')" :disabled="!convertedData.length">导出 Excel SMCC</button>
    <button @click="exportToExcel('ASTS')" :disabled="!convertedData.length">导出 Excel ASTS</button>
    <button @click="exportToExcel('AXS')" :disabled="!convertedData.length">导出 Excel AXS</button>
    <button @click="exportToExcel()" :disabled="!convertedData.length">导出 Excel ALL</button>

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
import { itemNoPriceConstants, stateMap } from '../const/astsToysConstants.js'
const sheet_tab_name = import.meta.env.VITE_SHEET_TAB;

const convertedData = ref([])
const tableHeaders = ref([])
const markedOrderNumbers = ref(new Set())


function excelSerialToDate(serial) {
  const msPerDay = 24 * 60 * 60 * 1000
  const excelEpoch = new Date(Date.UTC(1899, 11, 30)) // 注意：Excel起点是1900-01-00，但JS从1899-12-30起
  const date = new Date(excelEpoch.getTime() + serial * msPerDay)
  return date.toISOString().slice(0, 10) // 格式为 YYYY-MM-DD
}
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

      const stringDate = row['Purchased Date'] ? row['Purchased Date'].toString().trim() : ''
      // try {
      //    stringDate = excelSerialToDate(row['Purchased Date'])
      // } catch (error) {
      //    stringDate = row['Purchased Date']
      // }


      return {
        '销售平台': 'MyDeal',
        '日期': stringDate || '',
        '订单号': orderNo,
        'SKU': row['SKU'] || '',
        '名称': '',
        '数量': row['Quantity'] || '',
        '价格': Number(row['Price(Per Unit)']) * Number(row['Quantity']).toString() || '',
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
    const sheetName = sheet_tab_name
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

// 导出 Excel（过滤删除线记录和按SKU前缀筛选）
function exportToExcel(skuPrefix = '') {
  let filteredData = convertedData.value.filter(row => !row.markedForExclusion)

  // 如果提供了SKU前缀，则进一步筛选
  if (skuPrefix) {
    filteredData = filteredData.filter(row => {
      const sku = row['SKU'] || ''
      return sku.toString().startsWith(skuPrefix)
    })
  }

  // ✅ 只有当 skuPrefix === 'ASTS' 时，才自动补充 itemNum 和 importPrice
  if (skuPrefix === 'ASTS') {
    filteredData = filteredData.map(row => {
      const sku = row['SKU'] || ''

      const matched = itemNoPriceConstants.find(
        item => item.sku === sku
      )

      return {
        ...row,
        itemNum: matched ? matched.itemNum : '',
        importPrice: matched ? matched.importPrice : ''
      }
    })
  }

  const worksheet = XLSX.utils.json_to_sheet(filteredData)
  const workbook = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1')
  const randomSuffix = Math.floor(1000 + Math.random() * 9000) // 1000–9999
  XLSX.writeFile(workbook, `mydeal_orders_${randomSuffix}.xlsx`)
}



// 3 导入 Australia Post 专用 CSV 文件处理

function exportToCsvByXlsx(data, filename) {
  const worksheet = XLSX.utils.json_to_sheet(data)
  const workbook = XLSX.utils.book_new()

  XLSX.utils.book_append_sheet(workbook, worksheet, 'Aupost')

  XLSX.writeFile(workbook, filename, {
    bookType: 'csv',
    type: 'array'
  })
}
function formatState(state) {
  return stateMap[state?.trim()] || state
}

const handleFileUploadAuPost = (e) => {
  const file = e.target.files[0]
  if (!file) return

  const reader = new FileReader()

  reader.onload = (evt) => {
    const data = new Uint8Array(evt.target.result)

    // ✅ 读取 CSV / Excel
    const workbook = XLSX.read(data, { type: 'array' })
    const sheetName = workbook.SheetNames[0]
    const worksheet = workbook.Sheets[sheetName]

    // ✅ 转成 JSON（mydeal1 源数据）并且只保存 ASTS 开头的 SKU
    const mydealData = XLSX.utils.sheet_to_json(worksheet).filter(row => row.SKU && row.SKU.toUpperCase().startsWith('ASTS'))

    const aupostData = mydealData.map(row => {
      const sku = row.SKU?.trim()

      // ✅ 匹配 SKU 尺寸 & 重量
      const matchedItem = itemNoPriceConstants.find(
        item => item.sku === sku
      )

      let length = ''
      let width = ''
      let height = ''
      let weight = ''

      if (matchedItem?.size) {
        const [l, w, h] = matchedItem.size.split(',')
        length = l
        width = w
        height = h
      }

      if (matchedItem?.weight) {
        weight = matchedItem.weight
      }

      return {
        // ✅ 固定发件人
        'Send From Name': 'Alan',
        'Send From Address Line 1': '1/191 McCredie Rd',
        'Send From Suburb': 'Smithfield',
        'Send From State': 'NSW',
        'Send From Postcode': '2164',
        'Send From Phone Number': '0427116928',

        // ✅ 收件人（来自 mydeal1）
        'Deliver To Name': row.Name,
        'Deliver To Address Line 1': row.Address,
        'Deliver To Suburb': row.Suburb,
        'Deliver To State': formatState(row.State),
        'Deliver To Postcode': row.Postcode,
        'Deliver To Phone Number': row.Phone,

        // ✅ 物品信息
        'Item Description': sku,
        'Item Packaging Type': 'OWN_PACKAGING',
        'Item Delivery Service': 'PP',

        // ✅ 尺寸 & 重量
        'Item Length': length,
        'Item Width': width,
        'Item Height': height,
        'Item Weight': weight
      }
    })

    // ✅ 导出为 CSV（Australia Post 模板）
    exportToCsvByXlsx(aupostData, 'aupost_output.csv')
  }

  reader.readAsArrayBuffer(file)
}



</script>
