<template>
    <div>
        <h2>Kogan CSV 转 Excel</h2>
        <label>上传 manifest.csv
            <input type="file" accept=".csv" @change="handleFileUpload" />
        </label>
        <button @click="exportToExcel" :disabled="!processedData.length">导出 Excel</button>
    </div>
</template>

<script setup>
import * as XLSX from "xlsx"
import { ref } from "vue"
import { shippingRules, kogan_weight_map } from "../const/koganConstants";


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

function exportToExcel() {
    const worksheet = XLSX.utils.json_to_sheet(processedData.value)
    const workbook = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(workbook, worksheet, "KoganOrders")
    const randomSuffix = Math.floor(1000 + Math.random() * 9000) // 1000–9999
    XLSX.writeFile(workbook, `kogan_orders_${randomSuffix}.xlsx`)
}
</script>
