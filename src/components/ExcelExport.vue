<template>
  <div class="step-content">
    <h3>导出配置</h3>
    <p class="step-description">从当前多维表格导出为 Excel，支持相邻相同值合并与数字列合并策略。导出字段为当前视图的可见字段，支持自定义筛选条件进行数据过滤。</p>

    <div class="export-controls">
      <div class="view-info" v-if="viewInfo">
        <el-alert
          :title="`当前视图: ${viewInfo.name}`"
          :description="`可见字段 ${tableColumns.length} 个，数据记录 ${cachedRecords.length} 条${filterEnabled ? `，筛选后 ${filteredRecordCount} 条` : ''}`"
          type="info"
          :closable="false"
          show-icon
        />
      </div>
      <el-form label-width="100px" class="export-form">
        <el-form-item label="导出字段">
          <el-select v-model="selectedFieldIds" multiple collapse-tags filterable placeholder="选择导出字段" style="width: 420px">
            <el-option v-for="c in tableColumns" :key="c.id" :label="c.name" :value="c.id" />
          </el-select>
          <el-button class="ml8" @click="selectAll">全选</el-button>
          <el-button class="ml4" @click="clearAll">清空</el-button>
        </el-form-item>
        <el-form-item label="相同行合并">
          <el-switch v-model="mergeEnabled" />
        </el-form-item>
        <el-form-item v-if="mergeEnabled" label="合并字段">
          <el-select v-model="mergeFieldIds" multiple collapse-tags filterable placeholder="选择需要合并的字段" style="width: 420px">
            <el-option v-for="c in tableColumns" :key="c.id" :label="c.name" :value="c.id" />
          </el-select>
        </el-form-item>
        <el-form-item label="数字策略">
          <el-select v-model="numberMergeMode" style="width: 180px">
            <el-option label="保持首格值" value="fill" />
            <el-option label="取平均值" value="average" />
          </el-select>
        </el-form-item>
        <el-form-item label="数据筛选">
          <el-switch v-model="filterEnabled" />
          <span class="filter-tip">启用自定义筛选条件</span>
        </el-form-item>
        <div v-if="filterEnabled" class="filter-conditions">
          <div v-for="(condition, index) in filterConditions" :key="index" class="filter-condition">
            <el-select v-model="condition.fieldId" placeholder="选择字段" style="width: 150px">
              <el-option v-for="field in tableColumns" :key="field.id" :label="field.name" :value="field.id" />
            </el-select>
            <el-select v-model="condition.operator" placeholder="条件" style="width: 120px">
              <el-option label="等于" value="eq" />
              <el-option label="不等于" value="ne" />
              <el-option label="包含" value="contains" />
              <el-option label="不包含" value="not_contains" />
              <el-option label="大于" value="gt" />
              <el-option label="小于" value="lt" />
              <el-option label="大于等于" value="gte" />
              <el-option label="小于等于" value="lte" />
              <el-option label="为空" value="empty" />
              <el-option label="不为空" value="not_empty" />
            </el-select>
            <el-input 
              v-model="condition.value" 
              placeholder="匹配值" 
              style="width: 150px"
              :disabled="condition.operator === 'empty' || condition.operator === 'not_empty'"
            />
            <el-button type="danger" size="small" @click="removeFilterCondition(index)">删除</el-button>
          </div>
          <el-button type="primary" size="small" @click="addFilterCondition">添加筛选条件</el-button>
          <el-button type="info" size="small" @click="testFilter" :disabled="!filterConditions.length">测试筛选</el-button>
          
          <!-- 调试信息 -->
          <div v-if="filterConditions.length > 0" class="debug-info">
            <el-alert
              :title="`筛选状态: ${filterEnabled ? '已启用' : '已禁用'}`"
              :description="`原始数据: ${cachedRecords.length} 条，筛选后: ${filteredRecordCount} 条`"
              :type="filteredRecordCount > 0 ? 'success' : 'warning'"
              :closable="false"
              show-icon
            />
            <div v-if="filteredRecordCount === 0 && filterEnabled" class="debug-tips">
              <p><strong>筛选无结果可能的原因：</strong></p>
              <ul>
                <li>检查字段名称是否正确</li>
                <li>检查筛选条件是否合理</li>
                <li>检查匹配值是否正确</li>
                <li>尝试使用"包含"而不是"等于"</li>
              </ul>
              
              <!-- 显示字段值示例 -->
              <div v-if="cachedRecords.length > 0" class="field-samples">
                <p><strong>字段值示例（前3条记录）：</strong></p>
                <div v-for="(condition, index) in filterConditions" :key="index" class="field-sample">
                  <strong>{{ getFieldName(condition.fieldId) }}:</strong>
                  <div v-for="(record, recordIndex) in cachedRecords.slice(0, 3)" :key="recordIndex" class="sample-value">
                    记录{{ recordIndex + 1 }}: {{ formatFieldValue(record.fields[condition.fieldId]) }}
                  </div>
                </div>
              </div>
            </div>
          </div>
        </div>
        <el-form-item label="文件名">
          <el-input v-model="fileName" placeholder="导出文件名" style="width: 420px" />
        </el-form-item>
        <el-form-item>
          <el-button type="primary" @click="exportExcel" :disabled="selectedFieldIds.length === 0" :loading="exporting">导出Excel</el-button>
        </el-form-item>
      </el-form>
    </div>

    <div class="preview-table">
      <el-table :data="previewRows" style="width: 100%" max-height="420" v-loading="loading">
        <el-table-column v-for="f in selectedFieldIds" :key="f" :prop="f" :label="getFieldName(f)" />
      </el-table>
    </div>
  </div>
</template>

<script setup>
import { ref, computed, onMounted, watch } from 'vue'
import { bitable } from '@lark-base-open/js-sdk'
import * as XLSX from 'xlsx'

const tableColumns = ref([])
const selectedFieldIds = ref([])
const mergeEnabled = ref(true)
const numberMergeMode = ref('fill')
const fileName = ref('导出数据.xlsx')
const exporting = ref(false)
const loading = ref(false)
const cachedRecords = ref([]) // 原始记录缓存
const mergeFieldIds = ref([]) // 指定需要执行合并的字段（为空则默认所有选中字段）
const viewInfo = ref(null) // 当前视图信息
const filterEnabled = ref(false) // 是否启用筛选
const filterConditions = ref([]) // 筛选条件

const getFieldName = (id) => tableColumns.value.find(c => c.id === id)?.name || id
const getFieldType = (id) => tableColumns.value.find(c => c.id === id)?.type

// 格式化字段值用于显示
const formatFieldValue = (value) => {
  if (value === null || value === undefined) {
    return '(空)'
  }
  if (typeof value === 'object') {
    if (Array.isArray(value)) {
      // 处理飞书文本字段的段落数组
      return value.map(segment => {
        if (typeof segment === 'object') {
          // 文本段落
          if (segment.type === 'text' && segment.text) {
            return segment.text
          }
          // URL段落
          if (segment.type === 'url' && segment.text) {
            return segment.text
          }
          // 人员段落
          if (segment.mentionType === 'User' && segment.name) {
            return segment.name
          }
          // 文档段落
          if (segment.mentionType && segment.text) {
            return segment.text
          }
          // 其他对象
          if (segment.name) {
            return segment.name
          }
          if (segment.text) {
            return segment.text
          }
        }
        return String(segment)
      }).join('')
    } else if (value.name) {
      return value.name
    } else if (value.text) {
      return value.text
    } else {
      return JSON.stringify(value)
    }
  }
  return String(value)
}

const formatDate = (v) => {
  const d = typeof v === 'number' ? new Date(v) : new Date(String(v))
  if (isNaN(d.getTime())) return String(v)
  const yy = d.getFullYear()
  const mm = String(d.getMonth() + 1).padStart(2, '0')
  const dd = String(d.getDate()).padStart(2, '0')
  const hh = String(d.getHours()).padStart(2, '0')
  const mi = String(d.getMinutes()).padStart(2, '0')
  return `${yy}-${mm}-${dd} ${hh}:${mi}`
}

// 文本字段值（IOpenSegment[]）转纯文本
// 参考：文本字段返回的段落结构，包含 Text/Url/UserMention/DocumentMention 等
// 文档： https://lark-base-team.github.io/js-sdk-docs/zh/api/field/text
const renderTextSegments = (segments) => {
  if (!Array.isArray(segments)) return String(segments ?? '')
  return segments.map(seg => {
    if (!seg || typeof seg !== 'object') return ''
    // 文本/URL 段
    if (seg.text) return String(seg.text)
    // 人员/文档等 @ 段落
    if (seg.name) return String(seg.name)
    if (seg.link) return String(seg.link)
    return ''
  }).join('')
}

const extractText = (o) => {
  if (o == null) return ''
  if (typeof o !== 'object') return String(o)
  // 常见结构兜底键
  const keys = ['text', 'name', 'title', 'label', 'value', 'email', 'phone', 'url']
  for (const k of keys) {
    if (o[k] != null) return String(o[k])
  }
  // 位置/地址类
  if (o.address) return String(o.address)
  // 若包含 id + name
  if (o.id && o.name) return String(o.name)
  return ''
}

const normalizeForDisplay = (fieldId, val) => {
  if (val == null) return ''
  const t = getFieldType(fieldId)
  switch (t) {
    case 1: // 文本
      // 文本字段可能为 IOpenSegment[]
      if (Array.isArray(val)) return renderTextSegments(val)
      if (typeof val === 'object') {
        // 若 SDK 返回的是对象（极少见），尝试常见键
        return extractText(val)
      }
      return String(val)
    case 2: // 数字
    case 12: // 货币
      if (Array.isArray(val)) return val.join(', ')
      if (typeof val === 'object') return extractText(val)
      return val
    case 3: // 单选
      if (typeof val === 'object') return extractText(val)
      return String(val)
    case 4: // 多选
      if (Array.isArray(val)) return val.map(extractText).filter(Boolean).join(', ')
      return extractText(val)
    case 5: // 日期/时间
      return formatDate(val)
    case 7: // 复选框
      return val ? '是' : '否'
    case 8: // 人员
      if (Array.isArray(val)) return val.map(extractText).filter(Boolean).join(', ')
      return extractText(val)
    case 9: // 电话
    case 10: // 邮箱
    case 11: // URL
      return String(extractText(val))
    case 13: // 百分比
      if (typeof val === 'number') return `${(val * 100).toFixed(2)}%`
      return String(val)
    default:
      if (Array.isArray(val)) return val.map(extractText).filter(Boolean).join(', ')
      if (typeof val === 'object') return extractText(val)
      return String(val)
  }
}

const sleep = (ms) => new Promise(res => setTimeout(res, ms))

const getValidSelection = async (retry = 20, delay = 150) => {
  for (let i = 0; i < retry; i++) {
    try {
      const sel = await bitable.base.getSelection()
      if (sel && sel.tableId) return sel
    } catch {}
    await sleep(delay)
  }
  return null
}

const loadColumns = async () => {
  try {
    loading.value = true
    const inFeishu = !!(bitable && bitable.base)
    if (!inFeishu) {
      // mock
      tableColumns.value = [
        { id: 'fld_text', name: '文本', type: 1 },
        { id: 'fld_num', name: '数字', type: 2 },
        { id: 'fld_date', name: '日期', type: 5 },
      ]
      cachedRecords.value = []
      return
    }
    
    const selection = await getValidSelection()
    if (!selection || !selection.tableId) {
      console.warn('未获取到当前表选择，稍后重试或请在表中选中一个表。')
      return
    }
    
    const table = await bitable.base.getTableById(selection.tableId)
    
    // 获取当前视图的可见字段（未隐藏的字段）
    let visibleFieldIds = []
    let allFields = []
    
    if (selection.viewId) {
      try {
        // 获取当前视图
        const view = await table.getViewById(selection.viewId)
        // 获取视图元信息
        const viewMeta = await table.getViewMetaById(selection.viewId)
        viewInfo.value = {
          id: viewMeta.id,
          name: viewMeta.name,
          type: viewMeta.type
        }
        // 获取视图中的可见字段ID列表（有序）
        visibleFieldIds = await view.getVisibleFieldIdList()
        console.log('当前视图可见字段:', visibleFieldIds)
      } catch (viewError) {
        console.warn('获取视图字段失败，使用全部字段:', viewError)
        viewInfo.value = null
      }
    }
    
    // 获取所有字段元信息
    const fields = await table.getFieldMetaList()
    allFields = fields.map(f => ({ id: f.id, name: f.name, type: f.type }))
    
    // 如果有可见字段列表，则只显示可见字段；否则显示所有字段
    if (visibleFieldIds.length > 0) {
      tableColumns.value = allFields.filter(field => visibleFieldIds.includes(field.id))
    } else {
      tableColumns.value = allFields
    }
    
    // 默认选中所有可用字段
    if (!selectedFieldIds.value.length) {
      selectedFieldIds.value = tableColumns.value.map(c => c.id)
    }
    
    // 拉取当前视图的数据
    let records
    try {
      if (selection.viewId) {
        // 使用当前视图获取数据（包含筛选和排序）
        const view = await table.getViewById(selection.viewId)
        records = await view.getRecords({ pageSize: 1000 })
      } else {
        // 如果没有视图，获取全部记录
        records = await table.getRecords({ pageSize: 1000 })
      }
    } catch (viewError) {
      console.warn('获取视图数据失败，使用全部记录:', viewError)
      records = await table.getRecords({ pageSize: 1000 })
    }
    
    cachedRecords.value = records.records
    console.log('加载完成，字段数:', tableColumns.value.length, '记录数:', cachedRecords.value.length)
  } catch (e) {
    console.error('加载字段和数据失败:', e)
  } finally {
    loading.value = false
  }
}

const buildRowSpanMerges = (rows, fieldIds) => {
  const merges = []
  if (!rows.length) return merges
  const matrix = rows.map(r => fieldIds.map(fid => r[fid]))
  fieldIds.forEach((fid, c) => {
    let s = 0
    while (s < matrix.length) {
      let e = s
      const base = matrix[s][c]
      while (e + 1 < matrix.length && matrix[e + 1][c] === base) e++
      if (e > s) merges.push({ s: { r: s + 1, c }, e: { r: e + 1, c } })
      s = e + 1
    }
  })
  return merges
}

const applyNumberPolicy = (rows, fieldIds, mode) => {
  const out = rows.map(r => ({ ...r }))
  const isNumberField = (fid) => tableColumns.value.find(c => c.id === fid)?.type === 2
  fieldIds.forEach(fid => {
    if (!isNumberField(fid)) return
    let s = 0
    while (s < out.length) {
      let e = s
      const base = out[s][fid]
      while (e + 1 < out.length && out[e + 1][fid] === base) e++
      if (e > s) {
        if (mode === 'average') {
          let sum = 0, count = 0
          for (let i = s; i <= e; i++) {
            const v = parseFloat(String(out[i][fid]).replace(/[^\d.-]/g, ''))
            if (!isNaN(v)) { sum += v; count++ }
          }
          const avg = count ? sum / count : parseFloat(String(base).replace(/[^\d.-]/g, ''))
          for (let i = s; i <= e; i++) out[i][fid] = avg
        } else {
          for (let i = s + 1; i <= e; i++) out[i][fid] = out[s][fid]
        }
      }
      s = e + 1
    }
  })
  return out
}

const previewRows = computed(() => {
  if (!cachedRecords.value.length) return []
  
  // 应用筛选条件
  const filteredRecords = applyFilters(cachedRecords.value)
  
  return filteredRecords.slice(0, 200).map(r => {
    const obj = {}
    selectedFieldIds.value.forEach(fid => obj[fid] = normalizeForDisplay(fid, r.fields[fid]))
    return obj
  })
})

// 筛选后的记录数量
const filteredRecordCount = computed(() => {
  if (!cachedRecords.value.length) return 0
  return applyFilters(cachedRecords.value).length
})

watch(selectedFieldIds, () => {
  // 仅依赖 computed 自动联动，此处保留以便将来扩展（如统计、校验）
})

onMounted(() => {
  // 在飞书环境下，等待 base/selection 就绪
  loadColumns()
})

const selectAll = () => { selectedFieldIds.value = tableColumns.value.map(c => c.id) }
const clearAll = () => { selectedFieldIds.value = [] }

// 筛选条件管理
const addFilterCondition = () => {
  filterConditions.value.push({
    fieldId: '',
    operator: 'eq',
    value: ''
  })
}

const removeFilterCondition = (index) => {
  filterConditions.value.splice(index, 1)
}

// 测试筛选功能
const testFilter = () => {
  console.log('=== 开始测试筛选 ===')
  console.log('原始数据:', cachedRecords.value)
  console.log('筛选条件:', filterConditions.value)
  
  const testResult = applyFilters(cachedRecords.value)
  console.log('筛选结果:', testResult)
  
  // 显示前几条匹配的记录作为示例
  if (testResult.length > 0) {
    console.log('匹配的记录示例:')
    testResult.slice(0, 3).forEach((record, index) => {
      console.log(`记录 ${index + 1}:`, record)
    })
  } else {
    console.log('没有匹配的记录')
    
    // 显示一些原始数据作为参考
    console.log('原始数据示例:')
    cachedRecords.value.slice(0, 3).forEach((record, index) => {
      console.log(`原始记录 ${index + 1}:`, record)
    })
  }
  
  console.log('=== 筛选测试完成 ===')
}

// 筛选数据
const applyFilters = (records) => {
  if (!filterEnabled.value || !filterConditions.value.length) {
    return records
  }
  
  console.log('开始筛选，原始记录数:', records.length)
  console.log('筛选条件:', filterConditions.value)
  
  const filteredRecords = records.filter(record => {
    const matches = filterConditions.value.every(condition => {
      if (!condition.fieldId) {
        console.log('跳过空字段ID的条件')
        return true
      }
      
      const fieldValue = record.fields[condition.fieldId]
      const conditionValue = condition.value
      const operator = condition.operator
      
      console.log(`检查字段 ${condition.fieldId}: 值="${fieldValue}", 条件="${conditionValue}", 操作符="${operator}"`)
      
      // 处理空值情况
      if (operator === 'empty') {
        const isEmpty = fieldValue === null || fieldValue === undefined || fieldValue === ''
        console.log('空值检查结果:', isEmpty)
        return isEmpty
      }
      if (operator === 'not_empty') {
        const isNotEmpty = fieldValue !== null && fieldValue !== undefined && fieldValue !== ''
        console.log('非空值检查结果:', isNotEmpty)
        return isNotEmpty
      }
      
      // 如果条件值为空，跳过此条件
      if (!conditionValue) {
        console.log('条件值为空，跳过')
        return true
      }
      
      // 处理字段值，考虑不同的数据类型
      let fieldStr = ''
      if (fieldValue === null || fieldValue === undefined) {
        fieldStr = ''
      } else if (typeof fieldValue === 'object') {
        if (Array.isArray(fieldValue)) {
          // 处理飞书文本字段的段落数组
          fieldStr = fieldValue.map(segment => {
            if (typeof segment === 'object') {
              // 文本段落
              if (segment.type === 'text' && segment.text) {
                return segment.text
              }
              // URL段落
              if (segment.type === 'url' && segment.text) {
                return segment.text
              }
              // 人员段落
              if (segment.mentionType === 'User' && segment.name) {
                return segment.name
              }
              // 文档段落
              if (segment.mentionType && segment.text) {
                return segment.text
              }
              // 其他对象
              if (segment.name) {
                return segment.name
              }
              if (segment.text) {
                return segment.text
              }
            }
            return String(segment)
          }).join('')
        } else if (fieldValue.name) {
          fieldStr = fieldValue.name
        } else if (fieldValue.text) {
          fieldStr = fieldValue.text
        } else {
          fieldStr = String(fieldValue)
        }
      } else {
        fieldStr = String(fieldValue)
      }
      
      const conditionStr = String(conditionValue)
      
      // 转换为小写进行比较（仅对文本比较）
      const fieldStrLower = fieldStr.toLowerCase()
      const conditionStrLower = conditionStr.toLowerCase()
      
      let result = false
      switch (operator) {
        case 'eq':
          result = fieldStr === conditionStr
          break
        case 'ne':
          result = fieldStr !== conditionStr
          break
        case 'contains':
          result = fieldStrLower.includes(conditionStrLower)
          break
        case 'not_contains':
          result = !fieldStrLower.includes(conditionStrLower)
          break
        case 'gt':
          result = parseFloat(fieldValue) > parseFloat(conditionValue)
          break
        case 'lt':
          result = parseFloat(fieldValue) < parseFloat(conditionValue)
          break
        case 'gte':
          result = parseFloat(fieldValue) >= parseFloat(conditionValue)
          break
        case 'lte':
          result = parseFloat(fieldValue) <= parseFloat(conditionValue)
          break
        default:
          result = true
      }
      
      console.log(`筛选结果: "${fieldStr}" ${operator} "${conditionStr}" = ${result}`)
      return result
    })
    
    if (matches) {
      console.log('记录匹配:', record)
    }
    return matches
  })
  
  console.log('筛选完成，结果记录数:', filteredRecords.length)
  return filteredRecords
}

const isTextField = (fid) => tableColumns.value.find(c => c.id === fid)?.type === 1
const normalizeTextKey = (v) => String(v ?? '').trim().toLowerCase()

// 在执行合并前，如合并字段包含文本字段，则按这些文本字段顺序排序，便于相同值连续，从而合并
const sortRowsForTextMerge = (rows, mergeTargets) => {
  const textKeys = (mergeTargets || []).filter(isTextField)
  if (!textKeys.length) return rows
  const sorted = rows.slice().sort((a, b) => {
    for (const fid of textKeys) {
      const av = normalizeTextKey(a[fid])
      const bv = normalizeTextKey(b[fid])
      if (av < bv) return -1
      if (av > bv) return 1
    }
    return 0
  })
  return sorted
}

const exportExcel = async () => {
  if (!selectedFieldIds.value.length) return
  try {
    exporting.value = true
    // 拉取数据
    const inFeishu = typeof window !== 'undefined' && window?.bitable
    let rows = []
    if (inFeishu) {
      // 优先使用缓存数据，若为空则兜底请求
      let records = cachedRecords.value
      if (!records.length) {
        const selection = await bitable.base.getSelection()
        const table = await bitable.base.getTableById(selection.tableId)
        let res
        try {
          // 使用正确的SDK方法获取当前视图数据
          if (selection.viewId) {
            const view = await table.getViewById(selection.viewId)
            res = await view.getRecords({ pageSize: 1000 })
          } else {
            res = await table.getRecords({ pageSize: 1000 })
          }
        } catch (viewError) {
          console.warn('获取视图数据失败，使用全部记录:', viewError)
          res = await table.getRecords({ pageSize: 1000 })
        }
        records = res.records
      }
      // 应用筛选条件
      const filteredRecords = applyFilters(records)
      
      rows = filteredRecords.map(r => {
        const obj = {}
        selectedFieldIds.value.forEach(fid => obj[fid] = normalizeForDisplay(fid, r.fields[fid]))
        return obj
      })
    } else {
      rows = previewRows.value
    }

    // 仅对被选为合并的字段应用策略
    const mergeTargets = (mergeFieldIds.value && mergeFieldIds.value.length)
      ? selectedFieldIds.value.filter(fid => mergeFieldIds.value.includes(fid))
      : selectedFieldIds.value

    // 文本字段排序，确保相同文本相邻以便合并
    rows = sortRowsForTextMerge(rows, mergeTargets)

    // 数字字段合并策略
    rows = applyNumberPolicy(rows, mergeTargets, numberMergeMode.value)

    const header = selectedFieldIds.value.map(fid => getFieldName(fid))
    const aoa = [header]
    rows.forEach(r => aoa.push(selectedFieldIds.value.map(fid => r[fid])))
    const ws = XLSX.utils.aoa_to_sheet(aoa)
    if (mergeEnabled.value) {
      const mergeTargetsForMerges = (mergeFieldIds.value && mergeFieldIds.value.length)
        ? selectedFieldIds.value.map(fid => mergeFieldIds.value.includes(fid) ? fid : null).filter(Boolean)
        : selectedFieldIds.value
      const merges = buildRowSpanMerges(rows, mergeTargetsForMerges)
      if (merges.length) ws['!merges'] = merges
    }
    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, ws, '导出数据')
    XLSX.writeFile(wb, fileName.value || '导出数据.xlsx')
  } catch (e) {
    console.error('导出失败', e)
  } finally {
    exporting.value = false
  }
}
</script>

<style scoped>
.export-controls { 
  background: #fff; 
  border: 1px solid #e5e7eb; 
  border-radius: 12px; 
  padding: 20px; 
  margin-bottom: 16px;
  box-shadow: 0 2px 8px rgba(0, 0, 0, 0.06);
}

.view-info {
  margin-bottom: 16px;
}

.view-info :deep(.el-alert) {
  border-radius: 8px;
  border: 1px solid #e0f2fe;
  background: linear-gradient(135deg, #f0f8ff 0%, #e0f2fe 100%);
  box-shadow: 0 2px 8px rgba(59, 130, 246, 0.1);
}

.filter-tip {
  margin-left: 8px;
  color: #6b7280;
  font-size: 14px;
}

.filter-conditions {
  margin: 16px 0;
  padding: 16px;
  background: #f8fafc;
  border-radius: 8px;
  border: 1px solid #e5e7eb;
}

.filter-condition {
  display: flex;
  align-items: center;
  gap: 12px;
  margin-bottom: 12px;
  padding: 12px;
  background: white;
  border-radius: 6px;
  border: 1px solid #e5e7eb;
}

.filter-condition:last-child {
  margin-bottom: 0;
}

.debug-info {
  margin-top: 16px;
  padding: 12px;
  background: #f0f9ff;
  border-radius: 6px;
  border: 1px solid #bae6fd;
}

.debug-tips {
  margin-top: 8px;
  padding: 8px;
  background: #fef3c7;
  border-radius: 4px;
  border: 1px solid #f59e0b;
}

.debug-tips p {
  margin: 0 0 8px 0;
  color: #92400e;
  font-size: 14px;
}

.debug-tips ul {
  margin: 0;
  padding-left: 20px;
  color: #92400e;
  font-size: 13px;
}

.debug-tips li {
  margin: 4px 0;
}

.field-samples {
  margin-top: 12px;
  padding: 8px;
  background: #f3f4f6;
  border-radius: 4px;
  border: 1px solid #d1d5db;
}

.field-sample {
  margin: 8px 0;
  padding: 6px;
  background: white;
  border-radius: 4px;
  border: 1px solid #e5e7eb;
}

.sample-value {
  margin: 2px 0;
  padding: 2px 6px;
  background: #f9fafb;
  border-radius: 3px;
  font-size: 12px;
  color: #374151;
  font-family: monospace;
}

h3 {
  color: #1f2937;
  font-size: 18px;
  font-weight: 600;
  margin-bottom: 16px;
  display: flex;
  align-items: center;
  gap: 8px;
}

h3::before {
  content: '';
  width: 4px;
  height: 20px;
  background: linear-gradient(135deg, #3b82f6, #1d4ed8);
  border-radius: 2px;
  flex-shrink: 0;
}
.export-form :deep(.el-form-item__label) { color: #334155; }
.preview-table { margin-top: 8px; }
.step-description { 
  color: #6b7280; 
  margin: 6px 0 16px; 
  font-size: 14px;
  line-height: 1.6;
  background: #f8fafc;
  padding: 16px;
  border-radius: 8px;
  border-left: 4px solid #3b82f6;
  box-shadow: 0 1px 3px rgba(0, 0, 0, 0.1);
}
.step-actions { 
  margin-top: 16px; 
  display: flex; 
  justify-content: flex-end; 
  gap: 12px; 
}

.ml8 { margin-left: 8px; }
.ml4 { margin-left: 4px; }

/* 响应式设计 */
@media (max-width: 768px) {
  .export-controls {
    padding: 16px;
  }
  
  .export-form :deep(.el-form-item) {
    margin-bottom: 16px;
  }
  
  .export-form :deep(.el-form-item__label) {
    font-size: 14px;
    margin-bottom: 8px;
  }
  
  .step-actions {
    flex-direction: column;
    gap: 8px;
  }
  
  .step-actions .el-button {
    width: 100%;
  }
  
  .filter-condition {
    flex-direction: column;
    align-items: flex-start;
    gap: 8px;
  }
  
  .filter-condition .el-select,
  .filter-condition .el-input {
    width: 100% !important;
  }
}

@media (max-width: 480px) {
  .export-controls {
    padding: 12px;
  }
  
  h3 {
    font-size: 16px;
  }
  
  .step-description {
    padding: 12px;
    font-size: 13px;
  }
}
</style>


