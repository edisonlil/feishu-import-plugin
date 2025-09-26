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
      <div class="search-tip">
        <el-alert
          title="💡 搜索提示"
          type="info"
          :closable="false"
          show-icon
        >
          <template #default>
            <p>您可以在下拉框中输入关键词进行搜索：</p>
            <ul>
              <li>输入完整列名：如"姓名"</li>
              <li>输入部分关键词：如"名"</li>
              <li>输入拼音首字母：如"xm"（姓名）</li>
            </ul>
          </template>
        </el-alert>
      </div>
      
      <div class="mapping-container">
        <p class="step-description" style="margin-top:-4px">未选择映射的 Excel 列将在导入时被忽略，不会影响后续步骤。</p>
        <div class="mapping-item" v-for="(excelCol, index) in excelColumns" :key="index" 
             :class="{ 'auto-matched': columnMapping[index] }">
          <div class="excel-column">
            <span class="column-label">Excel列:</span>
            <el-tag type="info">{{ excelCol }}</el-tag>
          </div>
          <div class="mapping-arrow">→</div>
          <div class="table-column">
            <span class="column-label">多维表格列:</span>
            <el-select 
              v-model="columnMapping[index]" 
              placeholder="请选择对应列（支持输入搜索）"
              style="width: 200px"
              filterable
              clearable
              remote
              :remote-method="(query) => filterTableColumns(query, index)"
              :loading="false"
              no-data-text="没有找到匹配的列"
              no-match-text="没有找到匹配的列"
            >
              <el-option
                v-for="tableCol in filteredTableColumns[index] || tableColumns"
                :key="tableCol.id"
                :label="tableCol.name"
                :value="tableCol.id"
                :disabled="isColumnMapped(tableCol.id, index)"
              />
            </el-select>
            <el-tag v-if="columnMapping[index]" type="success" size="small" style="margin-left: 8px">
              已匹配
            </el-tag>
          </div>
        </div>
      </div>

      <!-- 合并单元格处理策略 -->
      <div class="merge-policy">
        <h4 class="policy-title">合并单元格处理策略</h4>
        <el-form label-width="140px" label-position="left" class="policy-form">
          <el-form-item label="文本列：">
            <el-radio-group v-model="mergePolicy.text" @change="onMergePolicyChange">
              <el-radio label="fill">向下填充</el-radio>
              <el-radio label="none">不处理</el-radio>
            </el-radio-group>
          </el-form-item>
          <el-form-item label="数字列：">
            <el-select v-model="mergePolicy.number" style="width: 220px" @change="onMergePolicyChange">
              <el-option label="向下填充" value="fill" />
              <el-option label="平均值填充" value="average" />
              <el-option label="递增填充（+1）" value="increment" />
            </el-select>
          </el-form-item>
        </el-form>
      </div>

      <!-- 更新模式设置 -->
      <div class="upsert-panel">
        <el-switch v-model="upsertEnabled" active-text="开启更新模式（存在则更新，不存在则新增）" />
        <div v-if="upsertEnabled" class="upsert-config">
          <el-form label-position="left" label-width="120px">
            <el-form-item label="条件列（Excel 列）">
              <el-select
                v-model="upsertKeyIndex"
                placeholder="请选择用作匹配条件的 Excel 列"
                filterable
                style="width: 320px"
              >
                <el-option
                  v-for="(name, idx) in excelColumns"
                  :key="idx"
                  :label="name"
                  :value="idx"
                />
              </el-select>
              <div class="upsert-hint" v-if="upsertKeyIndex !== null">
                <el-tag v-if="columnMapping[upsertKeyIndex]" type="success" size="small">已映射到字段：{{ getTableColumnName(columnMapping[upsertKeyIndex]) }}</el-tag>
                <el-tag v-else type="warning" size="small">请为该 Excel 列选择对应的多维表格列</el-tag>
              </div>
            </el-form-item>
          </el-form>
        </div>
      </div>
      
      <div class="step-actions">
        <el-button @click="prevStep">上一步</el-button>
        <el-button type="primary" @click="nextStep" :disabled="!canProceedMappingStep">
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
const filteredTableColumns = ref({})
// Upsert 设置
const upsertEnabled = ref(false)
const upsertKeyIndex = ref(null) // 作为匹配条件的 Excel 列索引
// 合并单元格策略 + 原始矩阵数据
const mergePolicy = ref({ text: 'fill', number: 'fill' })
const rawMatrix = ref([]) // 包含表头在内的二维数组
const rawMerges = ref([])
const rawHeaders = ref([])

// 计算属性
const isMappingComplete = computed(() => {
  return columnMapping.value.every(mapping => mapping !== null && mapping !== undefined)
})

// 是否允许从映射步骤进入下一步：
// - 常规模式：任何映射数量都可（可为空，表示全部忽略）
// - 更新模式：要求已选择条件列，且该列已映射到有效字段
const canProceedMappingStep = computed(() => {
  if (!upsertEnabled.value) return true
  if (upsertKeyIndex.value === null || upsertKeyIndex.value === undefined) return false
  const mappedField = columnMapping.value[upsertKeyIndex.value]
  return !!mappedField
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
    const parsed = await readExcelFile(uploadedFile.value)
    console.log('Excel原始矩阵:', parsed)
    
    if (!parsed || !parsed.matrix || parsed.matrix.length === 0) {
      throw new Error('Excel文件为空或格式不正确')
    }
    rawMatrix.value = parsed.matrix
    rawMerges.value = parsed.merges || []
    rawHeaders.value = parsed.headers || (parsed.matrix[0] || [])
    // 根据策略重建行对象数据
    rebuildDataFromPolicy()
    console.log('Excel列名:', excelColumns.value)
    
    // 获取当前表格的列信息
    console.log('获取表格列信息...')
    await loadTableColumns()
    console.log('表格列信息:', tableColumns.value)
    
    // 初始化列映射，智能匹配相同名称的列
    columnMapping.value = new Array(excelColumns.value.length).fill(null)
    
    // 智能匹配：支持多种匹配模式
    excelColumns.value.forEach((excelCol, index) => {
      const matchedColumn = tableColumns.value.find(tableCol => {
        const excelName = excelCol.trim()
        const tableName = tableCol.name.trim()
        
        // 1. 完全匹配
        if (tableName === excelName) return true
        
        // 2. 忽略大小写匹配
        if (tableName.toLowerCase() === excelName.toLowerCase()) return true
        
        // 3. 忽略空格和特殊字符匹配
        const normalizeExcel = excelName.replace(/[\s\-_]/g, '').toLowerCase()
        const normalizeTable = tableName.replace(/[\s\-_]/g, '').toLowerCase()
        if (normalizeTable === normalizeExcel) return true
        
        // 4. 包含匹配（Excel列名包含在表格列名中，或反之）
        if (tableName.includes(excelName) || excelName.includes(tableName)) return true
        
        return false
      })
      
      if (matchedColumn) {
        columnMapping.value[index] = matchedColumn.id
        console.log(`智能匹配: "${excelCol}" -> "${matchedColumn.name}"`)
      }
    })
    
    // 显示匹配结果
    const matchedCount = columnMapping.value.filter(mapping => mapping !== null).length
    if (matchedCount > 0) {
      ElMessage.success(`智能匹配成功！已自动匹配 ${matchedCount} 个列`)
    }
    
    currentStep.value = 2
    ElMessage.success(`文件解析成功！共 ${excelData.value.length} 行数据，${excelColumns.value.length} 列`)
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
        let jsonData = XLSX.utils.sheet_to_json(worksheet, { 
          header: 1, // 使用数字作为列名
          defval: '' // 空单元格的默认值
        })

        // 处理合并单元格：将合并区域内的空单元格填充为首格的值
        try {
          const merges = worksheet['!merges'] || []
          if (merges.length > 0) {
            console.log('检测到合并区域数量:', merges.length)
            jsonData = fillMergedCells(jsonData, merges)
          }
        } catch (mergeErr) {
          console.warn('处理合并单元格时出现问题（已忽略）：', mergeErr)
        }
        
        console.log('JSON转换完成，数据行数:', jsonData.length)
        
        if (jsonData.length === 0) {
          throw new Error('Excel文件中没有数据')
        }
        
        // 返回更丰富的结构，便于策略重建
        const headers = jsonData[0] || []
        const dataRows = jsonData.slice(1)
        resolve({
          matrix: jsonData,
          headers,
          rows: dataRows,
          merges: worksheet['!merges'] || []
        })
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

// 将合并区域的首格值下填到区域内空单元格
const fillMergedCells = (matrix, merges) => {
  // 深拷贝二维数组，避免就地修改带来副作用
  const data = matrix.map(row => row.slice())
  merges.forEach(range => {
    const { s, e } = range // s: start {r,c}, e: end {r,c}
    const startRow = s.r, startCol = s.c
    const endRow = e.r, endCol = e.c
    const topLeft = (data[startRow] && data[startRow][startCol]) !== undefined ? data[startRow][startCol] : ''
    for (let r = startRow; r <= endRow; r++) {
      // 确保行存在
      if (!data[r]) data[r] = []
      for (let c = startCol; c <= endCol; c++) {
        const cur = data[r][c]
        if (cur === undefined || cur === null || cur === '') {
          data[r][c] = topLeft
        }
      }
    }
  })
  return data
}

// 根据策略从原始矩阵生成 excelData/excelColumns
const rebuildDataFromPolicy = () => {
  // 拿原始矩阵（包含表头）
  let matrix = rawMatrix.value
  const headers = rawHeaders.value
  const merges = rawMerges.value

  // 先从未处理矩阵开始，再按策略处理（对合并区域进行不同策略）
  let processed = matrix.map(row => row.slice())

  if (merges && merges.length) {
    // 文本策略
    if (mergePolicy.value.text === 'fill') {
      processed = fillMergedCells(processed, merges)
    }
    // 数字策略
    if (mergePolicy.value.number !== 'none') {
      processed = applyNumberMergeStrategy(processed, merges, mergePolicy.value.number)
    }
  }

  // 构建行为对象
  const hdr = headers
  const dataRows = processed.slice(1)
  excelColumns.value = hdr
  excelData.value = dataRows.map(row => {
    const obj = {}
    hdr.forEach((h, i) => {
      obj[h || `列${i + 1}`] = row[i] ?? ''
    })
    return obj
  })
}

// 对数字列的合并区域应用策略：fill/average/increment
const applyNumberMergeStrategy = (matrix, merges, mode) => {
  const data = matrix.map(r => r.slice())
  merges.forEach(range => {
    const { s, e } = range
    const startRow = s.r, startCol = s.c
    const endRow = e.r, endCol = e.c

    // 检测首格是否为数字
    const top = data[startRow]?.[startCol]
    const topNum = Number(top)
    const isNum = !isNaN(topNum)
    console.log('检测首格是否为数字:'+isNum + ' 首格值:'+top + 'mode:'+mode)
    if (!isNum) return

    mode = 'average'
    if (mode === 'fill') {
      for (let r = startRow; r <= endRow; r++) {
        for (let c = startCol; c <= endCol; c++) {
          const cur = data[r][c]
          if (cur === undefined || cur === null || cur === '') data[r][c] = topNum
        }
      }
    } else if (mode === 'average') {
      // 计算区域内已有数字的平均值，否则用首格
      let sum = 0, count = 0
      for (let r = startRow; r <= endRow; r++) {
        for (let c = startCol; c <= endCol; c++) {
          const v = Number(data[r][c])
          if (!isNaN(v)) { sum += v; count++ }
        }
      }
      console.log('sum:'+sum + ' count:'+count)
      const avg = count > 0 ? sum / count : topNum
      for (let r = startRow; r <= endRow; r++) {
        for (let c = startCol; c <= endCol; c++) {
          const cur = data[r][c]
          console.log('avg:'+avg)
          if (cur === undefined || cur === null || cur === '') data[r][c] = avg
        }
      }
    } else if (mode === 'increment') {
      // 按行优先递增：首格为 topNum，后续单元依次 +1
      let val = topNum
      for (let r = startRow; r <= endRow; r++) {
        for (let c = startCol; c <= endCol; c++) {
          if (r === startRow && c === startCol) continue
          val += 1
          const cur = data[r][c]
          if (cur === undefined || cur === null || cur === '') data[r][c] = val
        }
      }
    }
  })
  return data
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

const filterTableColumns = (query, index) => {
  if (!query || query.trim() === '') {
    filteredTableColumns.value[index] = tableColumns.value
    return
  }
  
  const searchQuery = query.toLowerCase().trim()
  const filtered = tableColumns.value.filter(tableCol => {
    const name = tableCol.name.toLowerCase()
    
    // 1. 完全匹配
    if (name === searchQuery) return true
    
    // 2. 包含匹配
    if (name.includes(searchQuery)) return true
    
    // 3. 开头匹配
    if (name.startsWith(searchQuery)) return true
    
    // 4. 拼音首字母匹配
    const pinyinInitials = getPinyinInitials(tableCol.name).toLowerCase()
    if (pinyinInitials.includes(searchQuery)) return true
    
    return false
  })
  
  // 按匹配度排序
  filtered.sort((a, b) => {
    const aName = a.name.toLowerCase()
    const bName = b.name.toLowerCase()
    
    // 完全匹配优先
    if (aName === searchQuery && bName !== searchQuery) return -1
    if (bName === searchQuery && aName !== searchQuery) return 1
    
    // 开头匹配次优先
    if (aName.startsWith(searchQuery) && !bName.startsWith(searchQuery)) return -1
    if (bName.startsWith(searchQuery) && !aName.startsWith(searchQuery)) return 1
    
    // 其他按字母顺序
    return aName.localeCompare(bName)
  })
  
  filteredTableColumns.value[index] = filtered
}

// 简单的拼音首字母提取（可以后续优化为更完整的拼音库）
const getPinyinInitials = (text) => {
  // 这里是一个简单的实现，实际项目中可以使用完整的拼音库
  const pinyinMap = {
    // 基础信息
    '姓名': 'xm', '名字': 'mz', '用户': 'yh', '客户': 'kh', '人员': 'ry',
    '联系': 'lx', '信息': 'xx', '资料': 'zl', '档案': 'da',
    
    // 联系方式
    '电话': 'dh', '手机': 'sj', '邮箱': 'yx', '地址': 'dz', '邮编': 'yb',
    '传真': 'cz', 'QQ': 'qq', '微信': 'wx', '微博': 'wb',
    
    // 组织信息
    '部门': 'bm', '职位': 'zw', '公司': 'gs', '单位': 'dw', '机构': 'jg',
    '团队': 'td', '小组': 'xz', '科室': 'ks', '车间': 'cj',
    
    // 财务信息
    '金额': 'je', '价格': 'jg', '费用': 'fy', '成本': 'cb', '收入': 'sr',
    '支出': 'zc', '利润': 'lr', '预算': 'ys', '报销': 'bx',
    
    // 时间信息
    '日期': 'rq', '时间': 'sj', '开始': 'ks', '结束': 'js', '创建': 'cj',
    '更新': 'gx', '修改': 'xg', '删除': 'sc', '完成': 'wc',
    
    // 状态信息
    '状态': 'zt', '备注': 'bz', '说明': 'sm', '描述': 'ms', '详情': 'xq',
    '类型': 'lx', '分类': 'fl', '标签': 'bq', '标记': 'bj',
    
    // 其他常用
    '编号': 'bh', '代码': 'dm', 'ID': 'id', '序号': 'xh', '排序': 'px',
    '等级': 'dj', '级别': 'jb', '权限': 'qx', '角色': 'js'
  }
  
  // 直接匹配
  if (pinyinMap[text]) return pinyinMap[text]
  
  // 部分匹配（查找包含的词汇）
  for (const [key, value] of Object.entries(pinyinMap)) {
    if (text.includes(key)) {
      return value
    }
  }
  
  // 默认返回原文本
  return text
}

const getTableColumnName = (columnId) => {
  const column = tableColumns.value.find(col => col.id === columnId)
  return column ? column.name : ''
}

const nextStep = () => {
  if (currentStep.value === 2) {
    // 校验 Upsert 条件
    if (upsertEnabled.value) {
      if (upsertKeyIndex.value === null || upsertKeyIndex.value === undefined) {
        ElMessage.warning('请先在更新模式中选择用于匹配的 Excel 条件列')
        return
      }
      const mappedField = columnMapping.value[upsertKeyIndex.value]
      if (!mappedField) {
        ElMessage.warning('条件列未映射到多维表格字段，请先完成映射')
        return
      }
    }
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

    // 若启用更新模式：构建现有记录索引（keyValue -> recordId）
    let upsertFieldId = null
    let existingIndex = null
    const makeKeyVariants = (val) => {
      if (val === undefined || val === null) return []
      const s = String(val).trim()
      const lower = s.toLowerCase()
      const sanitized = lower.replace(/[^a-z0-9\u4e00-\u9fa5]/g, '') // 去除非中英文与数字
      return Array.from(new Set([s, lower, sanitized]))
    }
    const isAutoNumberField = (fieldId) => {
      const col = tableColumns.value.find(c => c.id === fieldId)
      const n = (col?.name || '').toLowerCase()
      return n.includes('自动编号') || n.includes('auto')
    }
    if (upsertEnabled.value) {
      upsertFieldId = columnMapping.value[upsertKeyIndex.value]
      try {
        const recordIds = await table.getRecordIdList()
        const upsertField = await table.getFieldById(upsertFieldId)
        existingIndex = new Map()
        // 注意：大量数据时可能较慢，可按需优化分页/视图过滤
        for (const rid of recordIds) {
          try {
            const cellStr = await upsertField.getCellString(rid)
            const variants = makeKeyVariants(cellStr)
            for (const k of variants) {
              if (k) {
                if (!existingIndex.has(k)) existingIndex.set(k, rid)
              }
            }
          } catch (e) {
            // 忽略单元读取错误，继续
          }
        }
        console.log('已建立索引，条目数:', existingIndex.size)
      } catch (e) {
        console.error('构建更新索引失败：', e)
        ElMessage.error('构建更新索引失败：' + e.message)
        // 回退为纯新增
        upsertEnabled.value = false
      }
    }
    
    // 批量处理数据，每批200条
    const batchSize = 200
    const totalRows = excelData.value.length
    
    for (let i = 0; i < totalRows; i += batchSize) {
      const batch = excelData.value.slice(i, i + batchSize)
      const batchRecordsToCreate = []
      const batchRecordsToUpdate = [] // { recordId, fields }

      // 准备当前批次记录
      for (const row of batch) {
        try {
          const fields = {}
          columnMapping.value.forEach((mapping, index) => {
            if (mapping) {
              const fieldId = mapping
              const fieldValue = row[excelColumns.value[index]]
              if (fieldValue !== undefined && fieldValue !== null && fieldValue !== '') {
                // 若为更新模式，匹配用的条件字段不写入（无论其类型），
                // 防止自动编号等受限字段导致写入失败；由其余字段完成更新/新增。
                if (upsertEnabled.value && fieldId === upsertFieldId) {
                  // 跳过写入条件字段
                } else {
                  fields[fieldId] = fieldValue
                }
              }
            }
          })

          if (upsertEnabled.value && upsertFieldId) {
            const keyValue = row[excelColumns.value[upsertKeyIndex.value]]
            const variants = makeKeyVariants(keyValue)
            let hitId
            for (const k of variants) {
              if (existingIndex && existingIndex.has(k)) { hitId = existingIndex.get(k); break }
            }
            if (hitId) {
              batchRecordsToUpdate.push({ recordId: hitId, fields })
            } else {
              batchRecordsToCreate.push({ fields })
            }
          } else {
            batchRecordsToCreate.push({ fields })
          }
        } catch (error) {
          console.error('准备记录失败:', error)
          errorCount++
        }
      }

      // 先执行更新（逐条）
      if (batchRecordsToUpdate.length > 0) {
        for (const item of batchRecordsToUpdate) {
          try {
            await table.setRecord(item.recordId, { fields: item.fields })
            successCount += 1
          } catch (err) {
            console.error('更新记录失败:', err)
            errorCount += 1
          }
        }
      }

      // 再执行新增（批量）
      if (batchRecordsToCreate.length > 0) {
        try {
          await table.addRecords(batchRecordsToCreate)
          successCount += batchRecordsToCreate.length
        } catch (error) {
          console.error('批量新增失败:', error)
          errorCount += batchRecordsToCreate.length
        }
      }

      ElMessage.success(`已处理 ${Math.min(i + batchSize, totalRows)}/${totalRows} 行（更新 ${batchRecordsToUpdate.length}，新增 ${batchRecordsToCreate.length}）`)
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
  background-color: #1890ff;
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
  background-color: #1890ff;
  color: white;
}

.step.completed .step-number {
  background-color: #1890ff;
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
  color: #1890ff;
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

.search-tip {
  margin-bottom: 20px;
}

.search-tip :deep(.el-alert__content) {
  font-size: 13px;
}

.search-tip :deep(.el-alert__content ul) {
  margin: 8px 0 0 0;
  padding-left: 20px;
}

.search-tip :deep(.el-alert__content li) {
  margin: 4px 0;
  color: #646a73;
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
  transition: all 0.3s ease;
}

.mapping-item.auto-matched {
  border-color: #1890ff;
  background-color: #f0f8ff;
  box-shadow: 0 2px 4px rgba(24, 144, 255, 0.1);
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
  color: #1890ff;
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
  color: #1890ff;
  margin-bottom: 12px;
}

.success-result h3 {
  color: #1890ff;
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
  border-color: #1890ff;
}

:deep(.el-button) {
  border-radius: 6px;
  font-weight: 500;
}

:deep(.el-button--primary) {
  background-color: #1890ff;
  border-color: #1890ff;
}

:deep(.el-button--primary:hover) {
  background-color: #40a9ff;
  border-color: #40a9ff;
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

