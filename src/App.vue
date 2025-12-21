<script setup>
import { ref } from 'vue'
import MyDealCsvToExcel from './components/MyDealCsvToExcel.vue'
import MyDealExcelBackToCsv from './components/MyDealExcelBackToCsv.vue'
import KoganCsvToExcel from './components/KoganCsvToExcel.vue'
import KoganExcelBackToCsv from './components/KoganExcelBackToCsv.vue'
const sheet = import.meta.env.VITE_SHEET_TAB;
const activeTab = ref('MyDeal')

const tabs = [
  { name: 'MyDeal', value: 'MyDeal' },
  { name: 'Kogan', value: 'Kogan' },
]
</script>

<template>
  <div>
    <h2 style="margin: 40px;">Sheet name is {{ sheet }}. Today is {{ new Date().toISOString().split('T')[0]}}</h2>
    <div style="">
      <button v-for="tab in tabs" :key="tab.value" @click="activeTab = tab.value" :style="{
        padding: '10px 20px',
        border: '1px solid #ccc',
        borderBottom: activeTab === tab.value ? '3px solid #007BFF' : '1px solid #ccc',
        backgroundColor: activeTab === tab.value ? '#eef6ff' : '#fff',
        cursor: 'pointer',
        margin: '20px'
      }">
        {{ tab.name }}
      </button>
    </div>

    <div v-if="activeTab === 'MyDeal'">
      <MyDealCsvToExcel />
      <MyDealExcelBackToCsv />
    </div>


    <div v-if="activeTab === 'Kogan'">
      <KoganCsvToExcel />
      <KoganExcelBackToCsv/>
    </div>
  </div>
</template>
