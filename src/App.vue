<script setup>
import MyDealCsvToExcel from './components/MyDealCsvToExcel.vue'
import ExcelBackToCsv from "./components/ExcelBackToCsv.vue";
import KoganCsvToExcel from "./components/KoganCsvToExcel.vue";
</script>

<template>
  <h1>MyDeal</h1>
<MyDealCsvToExcel />
<ExcelBackToCsv/>

<h1>Kogan</h1>
<KoganCsvToExcel/>
</template>
