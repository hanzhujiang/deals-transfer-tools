<template>
    <div>
        <h2>日控表Excel 转 CSV</h2>
        <div>
            <label>MyDeal导出的csv：</label>
            <input type="file" accept=".csv" @change="handleCsvUpload" />
                  <br /><br />
        </div>

        <div>
            <label>上传日控表：</label>
            <input type="file" accept=".xlsx" @change="handleExcelUpload" />
                  <br /><br />
        </div>

        <button @click="mergeAndExport" :disabled="!csvData.length || !excelMap.size">
            合并并导出CSV
        </button>
    </div>
</template>

<script setup>
import { ref } from 'vue'
import * as XLSX from 'xlsx'

const csvData = ref([])
const excelMap = ref(new Map())
const headers = ref([])
import { mydeal_couriers_map } from "../const/myDealConstants";
const sheet_tab_name = import.meta.env.VITE_SHEET_TAB;

function handleCsvUpload(e) {
    const file = e.target.files[0]
    if (!file) return

    const reader = new FileReader()
    reader.onload = () => {
        const workbook = XLSX.read(reader.result, { type: 'binary' })
        const sheet = workbook.Sheets[workbook.SheetNames[0]]
        csvData.value = XLSX.utils.sheet_to_json(sheet, { defval: '' })
        headers.value = Object.keys(csvData.value[0] || {})
    }
    reader.readAsBinaryString(file)
}

function handleExcelUpload(e) {
    const file = e.target.files[0]
    if (!file) return

    const reader = new FileReader()
    reader.onload = () => {
        const workbook = XLSX.read(reader.result, { type: 'binary' })
        const sheet = workbook.Sheets[sheet_tab_name]; // 指定 sheet 名字
        const excelRows = XLSX.utils.sheet_to_json(sheet, {
            header: 1,       // 先读取为二维数组（每行一个数组）
            defval: '',
        })

        const headers = excelRows[2]
        const dataRows = excelRows.slice(3) // 第4行开始是数据


        const formatted = dataRows.map(row => {
            const obj = {}
            headers.forEach((key, i) => {
                obj[key] = row[i] ?? ''
            })
            return obj
        })


        excelMap.value.clear()
        formatted.forEach(row => {
            const orderNo = row['订单号']
            const trackingCode = row['发货单号'].toString().trim() || ''
            const raw_courier = row['物流公司'] || ''
            const courier = mydeal_couriers_map[raw_courier]
            if (orderNo&&trackingCode) {
                excelMap.value.set(orderNo, { trackingCode, courier })
            }
        })
    }
    reader.readAsBinaryString(file)
}

function mergeAndExport() {
    const updated = csvData.value.map(row => {
        const orderNo = row['Order No']
        if (excelMap.value.has(orderNo)) {
            const { trackingCode, courier } = excelMap.value.get(orderNo)
            return {
                ...row,
                'Tracking Code': trackingCode,
                'Courier': courier
            }
        }
        return row
    })

    const ws = XLSX.utils.json_to_sheet(updated)
    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, ws, 'Updated CSV')
    XLSX.writeFile(wb, 'updated_mydeal_order_list.csv')
}
</script>
