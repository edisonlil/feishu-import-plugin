<script setup>
import { ref } from 'vue'
import ExcelImport from './components/ExcelImport.vue'
import ExcelExport from './components/ExcelExport.vue'

const mode = ref('home') // home | import | export

const enterImport = () => (mode.value = 'import')
const enterExport = () => (mode.value = 'export')
const backHome = () => (mode.value = 'home')
</script>

<template>
  <div v-if="mode === 'home'" class="home">
    <div class="home-card">
      <h2>请选择功能</h2>
      <p>支持导入 Excel 到当前多维表格，或从当前表导出 Excel。</p>
      <div class="home-actions">
        <button class="btn primary" @click="enterImport">导入 Excel</button>
        <button class="btn" @click="enterExport">导出 Excel</button>
      </div>
    </div>
  </div>
  <div v-else-if="mode === 'import'" class="page">
    <div class="page-header">
      <button class="link" @click="backHome">
        <svg width="16" height="16" viewBox="0 0 16 16" fill="none" xmlns="http://www.w3.org/2000/svg">
          <path d="M10 12L6 8L10 4" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/>
        </svg>
        返回
      </button>
      <h3>导入 Excel</h3>
    </div>
    <ExcelImport />
  </div>
  <div v-else-if="mode === 'export'" class="page">
    <div class="page-header">
      <button class="link" @click="backHome">
        <svg width="16" height="16" viewBox="0 0 16 16" fill="none" xmlns="http://www.w3.org/2000/svg">
          <path d="M10 12L6 8L10 4" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/>
        </svg>
        返回
      </button>
      <h3>导出 Excel</h3>
    </div>
    <ExcelExport />
  </div>
</template>

<style>
* {
  margin: 0;
  padding: 0;
  box-sizing: border-box;
}

body {
  font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'PingFang SC', 'Hiragino Sans GB', 'Microsoft YaHei', 'Helvetica Neue', Helvetica, Arial, sans-serif;
  background-color: #f8fafc;
  color: #1f2937;
  line-height: 1.6;
}

#app {
  width: 100%;
  height: 100vh;
}

.home {
  display: flex;
  align-items: center;
  justify-content: center;
  height: 100%;
}

.home-card {
  background: #ffffff;
  border: 1px solid #e5e7eb;
  border-radius: 12px;
  padding: 24px;
  box-shadow: 0 6px 20px rgba(31, 41, 55, 0.06);
  width: 92%;
  max-width: 520px;
  text-align: center;
}

.home-card h2 {
  margin-bottom: 8px;
  color: #1f2937;
}

.home-card p {
  color: #6b7280;
  margin-bottom: 20px;
}

.home-actions { display: flex; gap: 12px; justify-content: center; }

.btn {
  padding: 12px 20px;
  border-radius: 8px;
  border: 1px solid #d1d5db;
  background: #ffffff;
  color: #374151;
  cursor: pointer;
  font-weight: 500;
  font-size: 14px;
  transition: all 0.2s ease;
  box-shadow: 0 1px 2px rgba(0, 0, 0, 0.05);
}
.btn.primary {
  background: linear-gradient(135deg, #3b82f6, #1d4ed8);
  color: #ffffff;
  border-color: #3b82f6;
  box-shadow: 0 2px 4px rgba(59, 130, 246, 0.2);
}
.btn:hover { 
  transform: translateY(-1px);
  box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);
}
.btn.primary:hover { 
  transform: translateY(-1px);
  box-shadow: 0 4px 12px rgba(59, 130, 246, 0.3);
}
.btn:active {
  transform: translateY(0);
}

.page { 
  padding: 16px; 
  max-width: 1200px;
  margin: 0 auto;
}

@media (max-width: 768px) {
  .page {
    padding: 12px;
  }
  
  .page-header {
    flex-direction: column;
    align-items: flex-start;
    gap: 8px;
  }
  
  .page-header h3 {
    font-size: 16px;
  }
}
.page-header { 
  display: flex; 
  align-items: center; 
  gap: 12px; 
  margin-bottom: 16px; 
  padding: 16px 0;
  border-bottom: 1px solid #e5e7eb;
}

.page-header .link { 
  display: inline-flex;
  align-items: center;
  gap: 6px;
  background: transparent; 
  border: none; 
  color: #3b82f6; 
  cursor: pointer;
  font-size: 14px;
  font-weight: 500;
  padding: 8px 12px;
  border-radius: 6px;
  transition: all 0.2s ease;
  text-decoration: none;
}

.page-header .link:hover {
  background-color: #f0f8ff;
  color: #1d4ed8;
  transform: translateX(-2px);
}

.page-header .link:active {
  transform: translateX(0);
  background-color: #e0f2fe;
}

.page-header h3 {
  color: #1f2937;
  font-size: 18px;
  font-weight: 600;
  margin: 0;
}
</style>
