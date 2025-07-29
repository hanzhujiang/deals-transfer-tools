<script setup>
import { ref } from 'vue'
import MyDealCsvToExcel from './components/MyDealCsvToExcel.vue'
import ExcelBackToCsv from './components/ExcelBackToCsv.vue'
import KoganCsvToExcel from './components/KoganCsvToExcel.vue'
import KoganExcelBackToCsv from './components/KoganExcelBackToCsv.vue'

const activeTab = ref('MyDeal')

const tabs = [
  { name: 'MyDeal', value: 'MyDeal' },
  { name: 'Kogan', value: 'Kogan' },
]
</script>

<template>
  <div>
    <div style="">
      <button v-for="tab in tabs" :key="tab.value" @click="activeTab = tab.value" :style="{
        padding: '10px 20px',
        border: '1px solid #ccc',
        borderBottom: activeTab === tab.value ? '3px solid #007BFF' : '1px solid #ccc',
        backgroundColor: activeTab === tab.value ? '#eef6ff' : '#fff',
        cursor: 'pointer'
      }">
        {{ tab.name }}
      </button>
    </div>

    <div v-if="activeTab === 'MyDeal'">
      <MyDealCsvToExcel />
      <ExcelBackToCsv />
    </div>


    <div v-if="activeTab === 'Kogan'">
      <KoganCsvToExcel />
      <KoganExcelBackToCsv/>
    </div>
  </div>
</template>
