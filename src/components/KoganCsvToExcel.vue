<template>
    <div>
        <h2>Kogan CSV 转 Excel</h2>
        <label>上传 manifest.csv
            <input type="file" accept=".csv" @change="handleFileUpload" />
        </label>
        <button @click="exportToExcel('SMCC')" :disabled="!processedData.length">导出 Excel SMCC</button>
        <button @click="exportToExcel('ASTS')" :disabled="!processedData.length">导出 Excel ASTS</button>
        <button @click="exportToExcel('AXH')" :disabled="!processedData.length">导出 Excel AXH</button>
        <button @click="exportToExcel()" :disabled="!processedData.length">导出 Excel ALL</button>



            <label>
        Kogan CSV,for Australia Post
      <input type="file" accept=".csv" @change="handleFileUploadAuPost" />
      <br /><br />
    </label>

    </div>
</template>

<script setup>
import * as XLSX from "xlsx"
import { ref } from "vue"
import { shippingRules, kogan_weight_map } from "../const/koganConstants";
import { itemNoPriceConstants, stateMap } from '../const/astsToysConstants.js'


const processedData = ref([])

function handleFileUpload(event) {
    const file = event.target.files[0]
    if (!file) return

    const reader = new FileReader()
    reader.onload = (e) => {
        const csv = e.target.result
        const workbook = XLSX.read(csv, { type: "binary" })
        const sheet = workbook.Sheets[workbook.SheetNames[0]]
        const jsonData = XLSX.utils.sheet_to_json(sheet)

        processedData.value = jsonData.map(row => {
            const sku = (row.ProductCode || "").replace("AXS-", "")
            const quantity = Number(row.Quantity || 0)
            const itemPrice = Number(row.ItemPrice || 0)

            return {
                销售平台: "Kogan",
                日期: row.OrderDate || "",
                订单号: row.OrderID || "",
                SKU: sku,
                名称: "",
                数量: quantity,
                MyDeal价格: "",
                MyDeal运费: "",
                总价格: itemPrice * quantity,
                总运费: calculateShipping(row.ProductCode, row.DeliveryPostCode),
                地址: [row.DeliveryAddress1, row.DeliveryAddress2, row.DeliverySuburb, row.DeliveryState, row.DeliveryPostCode]
                    .filter(Boolean)
                    .join(" "),
                联系人: row.DeliveryName || "",
                电话: row.DeliveryPhone || "",
            }
        })
    }
    reader.readAsBinaryString(file)
}

function calculateShipping(productCode, postcode) {
    const postcodeRanges = []
    for (const rule of shippingRules) {
        const { base, rate, ranges } = rule
        ranges.split(",").forEach(range => {
            const trimmed = range.trim()
            if (trimmed.includes("-")) {
                const [start, end] = trimmed.split("-").map(Number)
                postcodeRanges.push({ start, end, base, rate })
            } else {
                const val = Number(trimmed)
                postcodeRanges.push({ start: val, end: val, base, rate })
            }
        })
    }
    const weight = kogan_weight_map[productCode] || 0
    const post = Number(postcode)

    for (const entry of postcodeRanges) {
        if (post >= entry.start && post <= entry.end) {
            return +(entry.base + weight * entry.rate).toFixed(2)
        }
    }
    return 0
}

function exportToExcel(skuPrefix = '') {

    // to filter processedData based on skuPrefix
    if (skuPrefix) {
        processedData.value = processedData.value.filter(row => {
            const sku = row['SKU'] || ''
            return sku.toString().startsWith(skuPrefix)
        })
    }

    // ✅ 只有当 skuPrefix === 'ASTS' 时，才自动补充 itemNum 和 importPrice
    if (skuPrefix === 'ASTS') {
        processedData.value = processedData.value.map(row => {
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
    const worksheet = XLSX.utils.json_to_sheet(processedData.value)
    const workbook = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(workbook, worksheet, "KoganOrders")
    const randomSuffix = Math.floor(1000 + Math.random() * 9000) // 1000–9999
    XLSX.writeFile(workbook, `kogan_orders_${randomSuffix}.xlsx`)
}

const handleFileUploadAuPost = (e) => {

    // TODO: Implement Australia Post CSV handling similar to MyDealCsvToExcel.vue
//   const file = e.target.files[0]
//   if (!file) return

//   const reader = new FileReader()

//   reader.onload = (evt) => {
//     const data = new Uint8Array(evt.target.result)

//     // ✅ 读取 CSV / Excel
//     const workbook = XLSX.read(data, { type: 'array' })
//     const sheetName = workbook.SheetNames[0]
//     const worksheet = workbook.Sheets[sheetName]

//     // ✅ 转成 JSON（mydeal1 源数据）并且只保存 ASTS 开头的 SKU
//     const mydealData = XLSX.utils.sheet_to_json(worksheet).filter(row => row.SKU && row.SKU.toUpperCase().startsWith('ASTS'))

//     const aupostData = mydealData.map(row => {
//       const sku = row.SKU?.trim()

//       // ✅ 匹配 SKU 尺寸 & 重量
//       const matchedItem = itemNoPriceConstants.find(
//         item => item.sku === sku
//       )

//       let length = ''
//       let width = ''
//       let height = ''
//       let weight = ''

//       if (matchedItem?.size) {
//         const [l, w, h] = matchedItem.size.split(',')
//         length = l
//         width = w
//         height = h
//       }

//       if (matchedItem?.weight) {
//         weight = matchedItem.weight
//       }

//       return {
//         // ✅ 固定发件人
//         'Send From Name': 'Alan',
//         'Send From Address Line 1': '1 AMOR ST',
//         'Send From Suburb': 'ASQUITH',
//         'Send From State': 'NSW',
//         'Send From Postcode': '2077',
//         'Send From Phone Number': '0427116928',

//         // ✅ 收件人（来自 mydeal1）
//         'Deliver To Name': row.Name,
//         'Deliver To Address Line 1': row.Address,
//         'Deliver To Suburb': row.Suburb,
//         'Deliver To State': formatState(row.State),
//         'Deliver To Postcode': row.Postcode,
//         'Deliver To Phone Number': row.Phone,

//         // ✅ 物品信息
//         'Item Description': sku,
//         'Item Packaging Type': 'OWN_PACKAGING',
//         'Item Delivery Service': 'PP',

//         // ✅ 尺寸 & 重量
//         'Item Length': length,
//         'Item Width': width,
//         'Item Height': height,
//         'Item Weight': weight
//       }
//     })

//     // ✅ 导出为 CSV（Australia Post 模板）
//     exportToCsvByXlsx(aupostData, 'aupost_output.csv')
//   }

//   reader.readAsArrayBuffer(file)
}
</script>
