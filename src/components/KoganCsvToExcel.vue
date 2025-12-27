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


        <div>
            <br /><br />
            <label>
                Kogan CSV,for Australia Post
                <input type="file" accept=".csv" @change="handleFileUploadAuPost" />
                <br /><br />
            </label>
        </div>


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
    const file = e.target.files?.[0]
    if (!file) return

    const reader = new FileReader()

    // ✅ 从各种字段里取值（兼容不同 CSV/Excel 表头）
    const pick = (row, keys, def = '') => {
        for (const k of keys) {
            const v = row?.[k]
            if (v !== undefined && v !== null && String(v).trim() !== '') return v
        }
        return def
    }

    // ✅ 把 ProductCode 变成你常用的 sku：如 "AXS-ASTS-1168" -> "ASTS-1168"
    const normaliseSku = (raw) => {
        const s = String(raw || '').trim().toUpperCase()
        if (!s) return ''
        const idx = s.indexOf('ASTS')
        if (idx >= 0) return s.slice(idx) // 从 ASTS 开始截取
        return s
    }

    // ✅ 澳洲手机号格式兜底：把纯数字转字符串 + 补 0
    const formatPhoneAU = (raw) => {
        if (raw === undefined || raw === null) return ''
        let s = String(raw).trim()
        // 如果是科学计数/小数这种，先转整数
        if (/^\d+(\.\d+)?$/.test(s)) s = String(Math.trunc(Number(s)))
        // 常见：9位且不以0开头 -> 补0
        if (/^\d{9}$/.test(s) && !s.startsWith('0')) s = '0' + s
        return s
    }

    reader.onload = (evt) => {
        const data = new Uint8Array(evt.target.result)

        // ✅ 读取 CSV / Excel
        const workbook = XLSX.read(data, { type: 'array' })
        const sheetName = workbook.SheetNames[0]
        const worksheet = workbook.Sheets[sheetName]

        // ✅ 转 JSON（支持空值）
        const sourceData = XLSX.utils.sheet_to_json(worksheet, { defval: '' })

        // ✅ 只保留 ASTS 的行：SKU 或 ProductCode 任意命中
        const filtered = sourceData.filter((row) => {
            const rawSku = pick(row, ['SKU', 'Sku', 'sku', 'ProductCode', 'Product Code', 'productCode'])
            const sku = normaliseSku(rawSku)
            return sku && sku.startsWith('ASTS')
        })

        // ✅ 生成澳邮模板数据（并按 Quantity 展开）
        debugger;
        const aupostData = filtered.flatMap((row) => {
            const rawSku = pick(row, ['SKU', 'Sku', 'sku', 'ProductCode', 'Product Code', 'productCode'])
            const sku = normaliseSku(rawSku)
            debugger;
            const qty = Number(pick(row, ['Quantity', 'Qty', 'quantity'], 1)) || 1
            debugger;
            // ✅ 匹配 SKU 尺寸 & 重量（你的常量表）
            const matchedItem = itemNoPriceConstants.find(
                (item) => String(item.sku || '').trim().toUpperCase() === sku
            )

            let length = ''
            let width = ''
            let height = ''
            let weight = ''

            if (matchedItem?.size) {
                const [l, w, h] = String(matchedItem.size).split(',').map(s => String(s).trim())
                length = l || ''
                width = w || ''
                height = h || ''
            }

            if (matchedItem?.weight !== undefined && matchedItem?.weight !== null) {
                weight = String(matchedItem.weight).trim()
            }

            const deliverName = pick(row, ['DeliveryName', 'Name', 'Customer Name', 'Receiver Name'])
            const addr1 = pick(row, ['DeliveryAddress1', 'Address', 'Delivery Address 1', 'Address Line 1'])
            const addr2 = pick(row, ['DeliveryAddress2', 'Address 2', 'Delivery Address 2', 'Address Line 2'], '')
            const suburb = pick(row, ['DeliverySuburb', 'Suburb', 'City'])
            const state = pick(row, ['DeliveryState', 'State'])
            const postcode = pick(row, ['DeliveryPostCode', 'DeliveryPostcode', 'Postcode', 'Postal Code'])
            const phone = formatPhoneAU(pick(row, ['DeliveryPhone', 'Phone', 'Phone Number', 'Mobile']))

            // ✅ 1个数量 -> 1行；数量>1 -> 多行
            return Array.from({ length: qty }, () => ({
                // ✅ 固定发件人
                'Send From Name': 'Alan',
                'Send From Address Line 1': '1/191 McCredie Rd',
                'Send From Suburb': 'Smithfield',
                'Send From State': 'NSW',
                'Send From Postcode': '2164',
                'Send From Phone Number': '0427116928',

                // ✅ 收件人（新 CSV 格式）
                'Deliver To Name': deliverName,
                'Deliver To Address Line 1': addr1,
                'Deliver To Address Line 2': addr2,
                'Deliver To Suburb': suburb,
                'Deliver To State': formatState(state),
                'Deliver To Postcode': String(postcode).trim(),
                'Deliver To Phone Number': phone,

                // ✅ 物品信息
                'Item Description': sku,
                'Item Packaging Type': 'OWN_PACKAGING',
                'Item Delivery Service': 'PP',

                // ✅ 尺寸 & 重量
                'Item Length': length,
                'Item Width': width,
                'Item Height': height,
                'Item Weight': weight
            }))
        })

        console.log('aupostData', aupostData);

        exportToCsvByXlsx(aupostData, 'aupost_output.csv')
    }

    reader.readAsArrayBuffer(file)
}

</script>
