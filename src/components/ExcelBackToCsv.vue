<template>
    <div>
        <h2>回填发货信息到原始CSV</h2>

        <div>
            <label>上传原始CSV：</label>
            <input type="file" accept=".csv" @change="handleCsvUpload" />
        </div>
        <label>
            日控表的 Sheet 名称：
            <input v-model="markingSheetName" placeholder="" />
        </label>
        <br /><br />
        <div>
            <label>上传包含发货信息的Excel：</label>
            <input type="file" accept=".xlsx" @change="handleExcelUpload" />
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
const markingSheetName = ref('物控日报表 (2025年7月)')
import { mydeal_couriers_map } from "../const/myDealConstants";

function handleCsvUpload(e) {
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
        const sheet = workbook.Sheets['物控日报表 (2025年7月)'];
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
            console.log('orderNo', orderNo);

            const trackingCode = row['发货单号'] || ''
            const raw_courier = row['物流公司'] || ''
            const courier = mydeal_couriers_map[raw_courier]
            if (orderNo && trackingCode) {
                console.log('orderNo', orderNo);
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
    XLSX.writeFile(wb, 'updated_sample.csv')
}
</script>
