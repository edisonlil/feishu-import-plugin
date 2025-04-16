<template>
  <div class="excel-import">
    <!-- 步骤指示器 -->
    <div class="steps">
      <div class="step" :class="{ active: currentStep >= 1, completed: currentStep > 1 }">
        <div class="step-number">1</div>
        <div class="step-title">上传文件</div>
      </div>
      <div class="step" :class="{ active: currentStep >= 2, completed: currentStep > 2 }">
        <div class="step-number">2</div>
        <div class="step-title">列映射</div>
      </div>
      <div class="step" :class="{ active: currentStep >= 3, completed: currentStep > 3 }">
        <div class="step-number">3</div>
        <div class="step-title">数据预览</div>
      </div>
      <div class="step" :class="{ active: currentStep >= 4, completed: currentStep > 4 }">
        <div class="step-number">4</div>
        <div class="step-title">导入完成</div>
      </div>
    </div>

    <!-- 步骤1: 文件上传 -->
    <div v-if="currentStep === 1" class="step-content">
      <h3>上传Excel文件</h3>
      <el-upload
        ref="uploadRef"
        class="upload-demo"
        drag
        :auto-upload="false"
        :on-change="handleFileChange"
        :before-upload="beforeUpload"
        accept=".xlsx,.xls"
        :limit="1"
      >
        <el-icon class="el-icon--upload"><upload-filled /></el-icon>
        <div class="el-upload__text">
          将文件拖到此处，或<em>点击上传</em>
        </div>
        <template #tip>
          <div class="el-upload__tip">
            只能上传 xlsx/xls 文件，且不超过 10MB
          </div>
        </template>
      </el-upload>
      
      <div v-if="uploadedFile" class="file-info">
        <el-alert
          :title="`已选择文件: ${uploadedFile.name}`"
          type="success"
          :closable="false"
          show-icon
        />
        <el-button type="primary" @click="parseExcel" :loading="parsing">
          解析文件
        </el-button>
      </div>
    </div>

    <!-- 步骤2: 列映射 -->
    <div v-if="currentStep === 2" class="step-content">
      <h3>建立列对应关系</h3>
      <p class="step-description">请将Excel列与多维表格列进行对应</p>
      
      <div class="mapping-container">
        <div class="mapping-item" v-for="(excelCol, index) in excelColumns" :key="index">
          <div class="excel-column">
            <span class="column-label">Excel列:</span>
            <el-tag type="info">{{ excelCol }}</el-tag>
          </div>
          <div class="mapping-arrow">→</div>
          <div class="table-column">
            <span class="column-label">多维表格列:</span>
            <el-select 
              v-model="columnMapping[index]" 
              placeholder="请选择对应列"
              style="width: 200px"
            >
              <el-option
                v-for="tableCol in tableColumns"
                :key="tableCol.id"
                :label="tableCol.name"
                :value="tableCol.id"
                :disabled="isColumnMapped(tableCol.id, index)"
              />
            </el-select>
          </div>
        </div>
      </div>
      
      <div class="step-actions">
        <el-button @click="prevStep">上一步</el-button>
        <el-button type="primary" @click="nextStep" :disabled="!isMappingComplete">
          下一步
        </el-button>
      </div>
    </div>

    <!-- 步骤3: 数据预览 -->
    <div v-if="currentStep === 3" class="step-content">
      <h3>数据预览</h3>
      <p class="step-description">请确认以下数据是否正确，确认无误后点击导入</p>
      
      <div class="preview-table">
        <el-table :data="previewData" style="width: 100%" max-height="400">
          <el-table-column
            v-for="(mapping, index) in columnMapping"
            :key="index"
            :prop="`col_${index}`"
            :label="getTableColumnName(mapping)"
            width="150"
          />
        </el-table>
      </div>
      
      <div class="step-actions">
        <el-button @click="prevStep">上一步</el-button>
        <el-button type="primary" @click="importData" :loading="importing">
          确认导入
        </el-button>
      </div>
    </div>

    <!-- 步骤4: 导入完成 -->
    <div v-if="currentStep === 4" class="step-content">
      <div class="success-result">
        <el-icon class="success-icon"><circle-check-filled /></el-icon>
        <h3>导入完成！</h3>
        <p>成功导入 {{ importResult.successCount }} 条记录</p>
        <el-button type="primary" @click="resetImport">重新导入</el-button>
      </div>
    </div>
  </div>
</template>

<script setup>
import { ref, computed, onMounted } from 'vue'
import { bitable } from '@lark-base-open/js-sdk'
import { UploadFilled, CircleCheckFilled } from '@element-plus/icons-vue'
import * as XLSX from 'xlsx'

// 响应式数据
const currentStep = ref(1)
const uploadedFile = ref(null)
const excelData = ref([])
const excelColumns = ref([])
const tableColumns = ref([])
const columnMapping = ref([])
const previewData = ref([])
const parsing = ref(false)
const importing = ref(false)
const importResult = ref({ successCount: 0, errorCount: 0 })

// 计算属性
const isMappingComplete = computed(() => {
  return columnMapping.value.every(mapping => mapping !== null && mapping !== undefined)
})

// 方法
const handleFileChange = (file) => {
  uploadedFile.value = file.raw
}

const beforeUpload = (file) => {
  // 支持所有Excel格式
  const isExcel = file.type === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' || 
                 file.type === 'application/vnd.ms-excel' ||
                 file.name.endsWith('.xlsx') ||
                 file.name.endsWith('.xls') ||
                 file.name.endsWith('.csv')
  const isLt20M = file.size / 1024 / 1024 < 20 // 增加文件大小限制到20MB

  if (!isExcel) {
    ElMessage.error('只能上传Excel文件(.xlsx, .xls, .csv)!')
    return false
  }
  if (!isLt20M) {
    ElMessage.error('文件大小不能超过 20MB!')
    return false
  }
  return false // 阻止自动上传
}

const parseExcel = async () => {
  if (!uploadedFile.value) {
    ElMessage.warning('请先选择文件')
    return
  }
  
  parsing.value = true
  console.log('开始解析Excel文件...')
  
  try {
    console.log('读取Excel文件...')
    const data = await readExcelFile(uploadedFile.value)
    console.log('Excel数据:', data)
    
    if (!data || data.length === 0) {
      throw new Error('Excel文件为空或格式不正确')
    }
    
    excelData.value = data
    excelColumns.value = Object.keys(data[0] || {})
    console.log('Excel列名:', excelColumns.value)
    
    // 获取当前表格的列信息
    console.log('获取表格列信息...')
    await loadTableColumns()
    console.log('表格列信息:', tableColumns.value)
    
    // 初始化列映射
    columnMapping.value = new Array(excelColumns.value.length).fill(null)
    
    currentStep.value = 2
    ElMessage.success(`文件解析成功！共 ${data.length} 行数据，${excelColumns.value.length} 列`)
  } catch (error) {
    console.error('解析失败:', error)
    ElMessage.error('文件解析失败: ' + error.message)
  } finally {
    parsing.value = false
  }
}

const readExcelFile = (file) => {
  return new Promise((resolve, reject) => {
    console.log('开始读取文件:', file.name, '大小:', file.size)
    
    const reader = new FileReader()
    reader.onload = (e) => {
      try {
        console.log('文件读取完成，开始解析...')
        const data = e.target.result
        
        // 尝试不同的解析方式
        let workbook
        try {
          workbook = XLSX.read(data, { type: 'binary' })
        } catch (binaryError) {
          console.log('二进制解析失败，尝试ArrayBuffer方式...')
          // 如果二进制解析失败，尝试ArrayBuffer方式
          const arrayBuffer = e.target.result
          workbook = XLSX.read(arrayBuffer, { type: 'array' })
        }
        
        console.log('工作簿解析成功，工作表:', workbook.SheetNames)
        
        if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
          throw new Error('Excel文件中没有找到工作表')
        }
        
        const sheetName = workbook.SheetNames[0]
        const worksheet = workbook.Sheets[sheetName]
        
        if (!worksheet) {
          throw new Error('无法读取第一个工作表')
        }
        
        console.log('开始转换为JSON...')
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { 
          header: 1, // 使用数字作为列名
          defval: '' // 空单元格的默认值
        })
        
        console.log('JSON转换完成，数据行数:', jsonData.length)
        
        if (jsonData.length === 0) {
          throw new Error('Excel文件中没有数据')
        }
        
        // 将第一行作为列名，其余作为数据
        const headers = jsonData[0]
        const dataRows = jsonData.slice(1)
        
        const result = dataRows.map(row => {
          const obj = {}
          headers.forEach((header, index) => {
            obj[header || `列${index + 1}`] = row[index] || ''
          })
          return obj
        })
        
        console.log('数据处理完成，最终数据:', result)
        resolve(result)
      } catch (error) {
        console.error('Excel解析错误:', error)
        reject(new Error('Excel文件解析失败: ' + error.message))
      }
    }
    
    reader.onerror = (error) => {
      console.error('文件读取错误:', error)
      reject(new Error('文件读取失败: ' + error.message))
    }
    
    reader.onabort = () => {
      reject(new Error('文件读取被中断'))
    }
    
    // 尝试二进制方式读取
    reader.readAsBinaryString(file)
  })
}

const loadTableColumns = async () => {
  try {
    // 检查是否在飞书环境中
    if (typeof bitable === 'undefined' || !bitable.base) {
      console.log('不在飞书环境中，使用模拟数据')
      // 模拟表格列信息，用于开发测试
      tableColumns.value = [
        { id: 'field1', name: '姓名', type: 'text' },
        { id: 'field2', name: '年龄', type: 'number' },
        { id: 'field3', name: '邮箱', type: 'text' },
        { id: 'field4', name: '部门', type: 'text' },
        { id: 'field5', name: '入职日期', type: 'date' }
      ]
      return
    }
    
    // 在边栏插件中，需要先获取当前选中的表格
    const selection = await bitable.base.getSelection()
    if (!selection.tableId) {
      throw new Error('请先选择一个表格')
    }
    
    const table = await bitable.base.getTableById(selection.tableId)
    const fieldList = await table.getFieldMetaList()
    tableColumns.value = fieldList
  } catch (error) {
    console.error('获取表格列信息失败:', error)
    // 如果获取失败，使用模拟数据
    console.log('使用模拟表格列数据')
    tableColumns.value = [
      { id: 'field1', name: '姓名', type: 'text' },
      { id: 'field2', name: '年龄', type: 'number' },
      { id: 'field3', name: '邮箱', type: 'text' },
      { id: 'field4', name: '部门', type: 'text' },
      { id: 'field5', name: '入职日期', type: 'date' }
    ]
    ElMessage.warning('无法获取表格列信息，使用模拟数据。请在飞书环境中使用完整功能。')
  }
}

const isColumnMapped = (columnId, currentIndex) => {
  return columnMapping.value.some((mapping, index) => 
    mapping === columnId && index !== currentIndex
  )
}

const getTableColumnName = (columnId) => {
  const column = tableColumns.value.find(col => col.id === columnId)
  return column ? column.name : ''
}

const nextStep = () => {
  if (currentStep.value === 2) {
    // 生成预览数据
    generatePreviewData()
  }
  currentStep.value++
}

const generatePreviewData = () => {
  previewData.value = excelData.value.slice(0, 10).map(row => {
    const previewRow = {}
    columnMapping.value.forEach((mapping, index) => {
      if (mapping) {
        // 使用字段ID作为key，但显示字段名称
        const fieldName = tableColumns.value.find(col => col.id === mapping)?.name || `字段${index + 1}`
        previewRow[`col_${index}`] = row[excelColumns.value[index]]
      }
    })
    return previewRow
  })
}

const importData = async () => {
  importing.value = true
  try {
    // 检查是否在飞书环境中
    if (typeof bitable === 'undefined' || !bitable.base) {
      // 模拟导入过程
      ElMessage.success('模拟导入完成！成功导入 ' + excelData.value.length + ' 条记录')
      importResult.value = { successCount: excelData.value.length, errorCount: 0 }
      currentStep.value = 4
      importing.value = false
      return
    }
    
    // 获取当前选中的表格
    const selection = await bitable.base.getSelection()
    if (!selection.tableId) {
      throw new Error('请先选择一个表格')
    }
    
    const table = await bitable.base.getTableById(selection.tableId)
    
    // 验证字段映射
    console.log('验证字段映射...')
    for (const mapping of columnMapping.value) {
      if (mapping) {
        const fieldExists = await table.isFieldExist(mapping)
        if (!fieldExists) {
          throw new Error(`字段ID ${mapping} 不存在，请重新选择列映射`)
        }
      }
    }
    let successCount = 0
    let errorCount = 0
    
    // 批量处理数据，每批200条
    const batchSize = 200
    const totalRows = excelData.value.length
    
    for (let i = 0; i < totalRows; i += batchSize) {
      const batch = excelData.value.slice(i, i + batchSize)
      const batchRecords = []
      
      // 准备批量记录
      for (const row of batch) {
        try {
          const fields = {}
          columnMapping.value.forEach((mapping, index) => {
            if (mapping) {
              // 使用字段ID而不是字段名称
              const fieldId = mapping
              const fieldValue = row[excelColumns.value[index]]
              if (fieldValue !== undefined && fieldValue !== null && fieldValue !== '') {
                fields[fieldId] = fieldValue
              }
            }
          })
          batchRecords.push({ fields })
        } catch (error) {
          console.error('准备记录失败:', error)
          errorCount++
        }
      }
      
      // 批量插入
      try {
        console.log('准备插入的记录:', batchRecords)
        console.log('字段映射:', columnMapping.value)
        console.log('表格列信息:', tableColumns.value)
        
        await table.addRecords(batchRecords)
        successCount += batchRecords.length
        ElMessage.success(`已导入 ${Math.min(i + batchSize, totalRows)}/${totalRows} 条记录`)
      } catch (error) {
        console.error('批量插入失败:', error)
        console.error('失败的记录:', batchRecords)
        errorCount += batchRecords.length
        ElMessage.error(`第 ${Math.floor(i/batchSize) + 1} 批数据导入失败: ${error.message}`)
      }
    }
    
    importResult.value = { successCount, errorCount }
    currentStep.value = 4
    
    if (errorCount > 0) {
      ElMessage.warning(`导入完成，成功 ${successCount} 条，失败 ${errorCount} 条`)
    } else {
      ElMessage.success(`成功导入 ${successCount} 条记录`)
    }
  } catch (error) {
    ElMessage.error('导入失败: ' + error.message)
  } finally {
    importing.value = false
  }
}

const prevStep = () => {
  if (currentStep.value > 1) {
    currentStep.value--
  }
}

const resetImport = () => {
  currentStep.value = 1
  uploadedFile.value = null
  excelData.value = []
  excelColumns.value = []
  columnMapping.value = []
  previewData.value = []
  importResult.value = { successCount: 0, errorCount: 0 }
}

onMounted(async () => {
  // 初始化时获取表格列信息
  await loadTableColumns()
})
</script>

<style scoped>
.excel-import {
  width: 100%;
  height: 100vh;
  padding: 16px;
  background-color: #f8f9fa;
  overflow-y: auto;
}

.steps {
  display: flex;
  justify-content: space-between;
  margin-bottom: 24px;
  padding: 16px 0;
  background: white;
  border-radius: 8px;
  box-shadow: 0 2px 4px rgba(0,0,0,0.1);
}

.step {
  display: flex;
  flex-direction: column;
  align-items: center;
  flex: 1;
  position: relative;
  padding: 0 8px;
}

.step:not(:last-child)::after {
  content: '';
  position: absolute;
  top: 15px;
  left: calc(50% + 15px);
  width: calc(100% - 30px);
  height: 2px;
  background-color: #e4e7ed;
  z-index: 1;
}

.step.completed:not(:last-child)::after {
  background-color: #00d4aa;
}

.step-number {
  width: 24px;
  height: 24px;
  border-radius: 50%;
  background-color: #e4e7ed;
  color: #909399;
  display: flex;
  align-items: center;
  justify-content: center;
  font-weight: bold;
  font-size: 12px;
  position: relative;
  z-index: 2;
}

.step.active .step-number {
  background-color: #00d4aa;
  color: white;
}

.step.completed .step-number {
  background-color: #00d4aa;
  color: white;
}

.step-title {
  margin-top: 6px;
  font-size: 12px;
  color: #606266;
  text-align: center;
  line-height: 1.2;
}

.step.active .step-title {
  color: #00d4aa;
  font-weight: 600;
}

.step-content {
  min-height: 300px;
  background: white;
  border-radius: 8px;
  padding: 20px;
  box-shadow: 0 2px 4px rgba(0,0,0,0.1);
}

.step-content h3 {
  margin-bottom: 12px;
  color: #1f2329;
  font-size: 16px;
  font-weight: 600;
}

.step-description {
  color: #646a73;
  margin-bottom: 16px;
  font-size: 14px;
}

.upload-demo {
  margin-bottom: 16px;
}

.file-info {
  margin-top: 16px;
}

.mapping-container {
  margin-bottom: 20px;
  max-height: 300px;
  overflow-y: auto;
}

.mapping-item {
  display: flex;
  align-items: center;
  margin-bottom: 12px;
  padding: 12px;
  border: 1px solid #e4e7ed;
  border-radius: 6px;
  background-color: #fafafa;
  gap: 12px;
}

.excel-column, .table-column {
  display: flex;
  align-items: center;
  gap: 8px;
  flex: 1;
}

.column-label {
  font-weight: 500;
  color: #646a73;
  font-size: 14px;
  white-space: nowrap;
}

.mapping-arrow {
  font-size: 16px;
  color: #00d4aa;
  font-weight: bold;
  margin: 0 8px;
}

.preview-table {
  margin-bottom: 16px;
  border: 1px solid #e4e7ed;
  border-radius: 6px;
  overflow: hidden;
  max-height: 300px;
  overflow-y: auto;
}

.step-actions {
  display: flex;
  justify-content: space-between;
  gap: 12px;
  margin-top: 16px;
}

.success-result {
  text-align: center;
  padding: 32px 16px;
}

.success-icon {
  font-size: 48px;
  color: #00d4aa;
  margin-bottom: 12px;
}

.success-result h3 {
  color: #00d4aa;
  margin-bottom: 8px;
  font-size: 18px;
}

.success-result p {
  color: #646a73;
  margin-bottom: 16px;
  font-size: 14px;
}

/* 飞书边栏插件样式优化 */
:deep(.el-upload-dragger) {
  border: 2px dashed #d1d5db;
  border-radius: 6px;
  padding: 20px;
  text-align: center;
  transition: all 0.3s;
}

:deep(.el-upload-dragger:hover) {
  border-color: #00d4aa;
}

:deep(.el-button) {
  border-radius: 6px;
  font-weight: 500;
}

:deep(.el-button--primary) {
  background-color: #00d4aa;
  border-color: #00d4aa;
}

:deep(.el-button--primary:hover) {
  background-color: #00b894;
  border-color: #00b894;
}

:deep(.el-select) {
  width: 100%;
}

:deep(.el-table) {
  font-size: 13px;
}

:deep(.el-table th) {
  background-color: #f8f9fa;
  color: #1f2329;
  font-weight: 600;
}

:deep(.el-tag) {
  border-radius: 4px;
}

:deep(.el-alert) {
  border-radius: 6px;
}

/* 响应式优化 */
@media (max-width: 400px) {
  .mapping-item {
    flex-direction: column;
    align-items: flex-start;
    gap: 8px;
  }
  
  .mapping-arrow {
    align-self: center;
    transform: rotate(90deg);
  }
  
  .excel-column, .table-column {
    width: 100%;
  }
}
</style>

