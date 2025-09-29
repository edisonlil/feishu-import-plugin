<template>
  <div class="step-content">
    <h3>导出配置</h3>
    <p class="step-description">从当前多维表格导出为 Excel，支持相邻相同值合并与数字列合并策略。</p>

    <div class="export-controls">
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

    <div class="step-actions">
      <el-button type="primary" @click="exportExcel" :disabled="selectedFieldIds.length === 0" :loading="exporting">导出Excel</el-button>
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

const getFieldName = (id) => tableColumns.value.find(c => c.id === id)?.name || id
const getFieldType = (id) => tableColumns.value.find(c => c.id === id)?.type

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
    const fields = await table.getFieldMetaList()
    tableColumns.value = fields.map(f => ({ id: f.id, name: f.name, type: f.type }))
    // 默认选中全部字段
    if (!selectedFieldIds.value.length) {
      selectedFieldIds.value = tableColumns.value.map(c => c.id)
    }
    // 拉取前1000条作为预览/导出数据源
    // 优先使用当前视图的筛选结果，否则使用全部记录
    let records
    try {
      // 尝试获取当前视图的筛选记录
      if (selection.viewId) {
        records = await table.getRecordsByView(selection.viewId, { pageSize: 1000 })
      } else {
        records = await table.getRecords({ pageSize: 1000 })
      }
    } catch (viewError) {
      console.warn('获取视图筛选记录失败，使用全部记录:', viewError)
      records = await table.getRecords({ pageSize: 1000 })
    }
    cachedRecords.value = records.records
  } catch (e) {
    console.error(e)
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
  return cachedRecords.value.slice(0, 200).map(r => {
    const obj = {}
    selectedFieldIds.value.forEach(fid => obj[fid] = normalizeForDisplay(fid, r.fields[fid]))
    return obj
  })
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
          // 尝试获取当前视图的筛选记录
          if (selection.viewId) {
            res = await table.getRecordsByView(selection.viewId, { pageSize: 1000 })
          } else {
            res = await table.getRecords({ pageSize: 1000 })
          }
        } catch (viewError) {
          console.warn('获取视图筛选记录失败，使用全部记录:', viewError)
          res = await table.getRecords({ pageSize: 1000 })
        }
        records = res.records
      }
      rows = records.map(r => {
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


