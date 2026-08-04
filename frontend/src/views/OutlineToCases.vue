<template>
  <div class="case-container" :class="{ expanded: parsedData }">
    <el-tabs v-model="activeTab" class="case-tabs" stretch>
      <el-tab-pane label="生成案例" name="generate">
        <el-card class="upload-card" shadow="hover">
      <template #header>
        <div class="card-header">
          <span>上传XMind测试大纲</span>
        </div>
      </template>

      <el-upload
        ref="uploadRef"
        class="upload-demo"
        drag
        :auto-upload="false"
        :on-change="handleFileChange"
        :on-remove="handleFileRemove"
        :file-list="fileList"
        accept=".xmind,.json"
        :limit="1"
        :disabled="parsing || previewing"
      >
        <el-icon class="el-icon--upload"><upload-filled /></el-icon>
        <div class="el-upload__text">
          将文件拖到此处，或<em>点击上传</em>
        </div>
        <template #tip>
          <div class="el-upload__tip">支持 .xmind 和 .json 格式（仅支持单个文件）</div>
        </template>
      </el-upload>

      <div class="action-buttons">
        <el-button type="primary" :loading="parsing" :disabled="fileList.length === 0" @click="handleParse">
          解析
        </el-button>
        <el-button :disabled="!parsedData" @click="resetAll">重置</el-button>
      </div>
    </el-card>

    <el-card v-if="parsedData" class="overview-card" shadow="never">
      <template #header>
        <div class="card-header">
          <span>解析结果概览</span>
        </div>
      </template>
      <el-row :gutter="16">
        <el-col :span="8">
          <div class="stat-stack">
            <div class="stat-item">
              <div class="stat-title">需求名称</div>
              <el-tooltip :content="parsedData.requirement_name" placement="top">
                <div class="stat-value">{{ parsedData.requirement_name }}</div>
              </el-tooltip>
            </div>
            <div class="stat-item">
              <div class="stat-title">测试点总数</div>
              <div class="stat-value">{{ parsedData.stats?.total || 0 }}</div>
            </div>
          </div>
        </el-col>
        <el-col :span="8">
          <div class="stat-stack">
            <div class="stat-item">
              <div class="stat-title">流程测试点</div>
              <div class="stat-value">{{ parsedData.stats?.by_type?.process || 0 }}</div>
            </div>
            <div class="stat-item">
              <div class="stat-title">规则/页面测试点</div>
              <div class="stat-value">
                {{ parsedData.stats?.by_type?.rule || 0 }} / {{ parsedData.stats?.by_type?.page_control || 0 }}
              </div>
            </div>
          </div>
        </el-col>
        <el-col :span="8">
          <div class="stat-stack">
            <div class="stat-item">
              <div class="stat-title">优先级分布</div>
              <el-tooltip :content="prioritySummary" placement="top">
                <div class="stat-value">{{ prioritySummary }}</div>
              </el-tooltip>
            </div>
            <div class="stat-item">
              <div class="stat-title">正/反例分布</div>
              <el-tooltip :content="subtypeSummary" placement="top">
                <div class="stat-value">{{ subtypeSummary }}</div>
              </el-tooltip>
            </div>
          </div>
        </el-col>
      </el-row>
    </el-card>

    <el-card v-if="parsedData" class="preview-card" shadow="never">
      <template #header>
        <div class="card-header">
          <span>预生成策略</span>
        </div>
      </template>
      <div class="strategy-text">
        默认自动挑选3-5个测试点，覆盖流程/规则及正/反例，优先高优先级。
      </div>
      <el-button type="primary" :loading="previewing" @click="handlePreview">
        预生成
      </el-button>
    </el-card>

    <el-card v-if="previewCases.length" class="preview-table-card" shadow="never">
      <template #header>
        <div class="card-header">
          <span>预生成结果</span>
        </div>
      </template>
      <el-table :data="previewCases" border style="width: 100%">
        <el-table-column prop="point_type" label="类型" width="80">
          <template #default="{ row }">{{ typeMap[row.point_type] || row.point_type }}</template>
        </el-table-column>
        <el-table-column prop="subtype" label="子类型" width="90">
          <template #default="{ row }">{{ subtypeMap[row.subtype] || row.subtype }}</template>
        </el-table-column>
        <el-table-column prop="priority" label="优先级" width="80">
          <template #default="{ row }">{{ priorityMap[row.priority] || row.priority }}</template>
        </el-table-column>
        <el-table-column prop="text" label="测试点" min-width="220" />
        <el-table-column label="前提条件" min-width="200">
          <template #default="{ row }">
            <div class="multi-line">{{ row.preconditions.join('\n') }}</div>
          </template>
        </el-table-column>
        <el-table-column label="测试步骤" min-width="200">
          <template #default="{ row }">
            <div class="multi-line">{{ row.steps.join('\n') }}</div>
          </template>
        </el-table-column>
        <el-table-column label="预期结果" min-width="200">
          <template #default="{ row }">
            <div class="multi-line">{{ row.expected_results.join('\n') }}</div>
          </template>
        </el-table-column>
        <el-table-column label="操作" width="140">
          <template #default="{ row, $index }">
            <el-button link type="primary" @click="openEdit(row, $index)">编辑</el-button>
            <el-button link type="danger" @click="removePreview($index)">删除</el-button>
          </template>
        </el-table-column>
      </el-table>
    </el-card>

    <el-card v-if="parsedData" class="confirm-card" shadow="never">
      <template #header>
        <div class="card-header">
          <span>确认与批量生成</span>
        </div>
      </template>
      <div class="confirm-actions">
        <el-select v-model="strategy" placeholder="选择策略" style="width: 180px;">
          <el-option label="标准模式" value="standard" />
          <el-option label="快速模式" value="fast" />
        </el-select>
        <el-button type="primary" :disabled="!previewId" @click="handleConfirm">确认预生成</el-button>
        <el-button :disabled="!parsedData" @click="handleBulkGenerate">批量生成</el-button>
      </div>
    </el-card>

    <el-card v-if="generationTaskId" class="progress-card" shadow="never">
      <template #header>
        <div class="card-header">
          <span>生成进度</span>
        </div>
      </template>
      <el-progress :percentage="Math.round((generationStatus?.progress || 0) * 100)" />
      <div class="progress-meta">
        <div v-if="currentSessionId">SessionID：{{ currentSessionId }}</div>
        <div>已完成：{{ generationStatus?.completed || 0 }} / {{ generationStatus?.total || 0 }}</div>
        <div>Token消耗：{{ generationStatus?.token_usage || 0 }}</div>
      </div>
      <el-divider />
      <div class="log-panel">
        <div class="log-title">实时日志</div>
        <div class="log-body">
          <div v-for="(log, idx) in generationStatus?.logs || []" :key="idx" class="log-item">
            {{ log }}
          </div>
        </div>
      </div>
    </el-card>

    <el-card v-if="generationStatus?.status === 'completed'" class="export-card" shadow="never">
      <template #header>
        <div class="card-header">
          <span>结果导出</span>
        </div>
      </template>
      <div class="export-actions">
        <el-button type="success" :loading="exportLoading" @click="handleExport">导出XMind</el-button>
      </div>
    </el-card>

    <el-dialog v-model="editDialogVisible" title="编辑用例" width="60%">
      <el-form label-position="top">
        <el-form-item label="前提条件">
          <el-input v-model="editForm.preconditions" type="textarea" rows="3" />
        </el-form-item>
        <el-form-item label="测试步骤">
          <el-input v-model="editForm.steps" type="textarea" rows="4" />
        </el-form-item>
        <el-form-item label="预期结果">
          <el-input v-model="editForm.expected_results" type="textarea" rows="3" />
        </el-form-item>
      </el-form>
      <template #footer>
        <el-button @click="editDialogVisible = false">取消</el-button>
        <el-button type="primary" @click="saveEdit">保存</el-button>
      </template>
    </el-dialog>
      </el-tab-pane>

      <el-tab-pane label="导出案例" name="session">
        <el-card class="session-card" shadow="hover">
          <template #header>
            <div class="card-header">
              <span>通过 session_id 重新生成并导出 XMind</span>
            </div>
          </template>
          <div class="session-actions">
            <el-input
              v-model="sessionIdInput"
              placeholder="请输入 session_id"
              clearable
              class="session-input"
              :title="sessionIdInput"
            />
            <el-button type="primary" :loading="sessionExportLoading" @click="handleExportBySession">
              导出XMind
            </el-button>
          </div>
        </el-card>
      </el-tab-pane>

      <el-tab-pane label="案例统计" name="import">

        <!-- 模块1: 导入文件 -->
        <el-card class="import-file-card" shadow="hover">
          <template #header>
            <div class="card-header">
              <span>导入文件</span>
            </div>
          </template>
          <el-upload
            ref="importUploadRef"
            class="upload-demo"
            drag
            :auto-upload="false"
            :on-change="handleImportFileChange"
            :on-remove="handleImportFileRemove"
            :file-list="importFileList"
            accept=".xmind,.json"
            :limit="1"
            :disabled="importParsing"
          >
            <el-icon class="el-icon--upload"><upload-filled /></el-icon>
            <div class="el-upload__text">
              将文件拖到此处，或<em>点击上传</em>
            </div>
            <template #tip>
              <div class="el-upload__tip">支持 .xmind 和 .json 格式，用于导入案例并统计各功能/步骤的案例数</div>
            </template>
          </el-upload>
        </el-card>

        <!-- 模块2: 操作 -->
        <el-card class="import-action-card" shadow="hover">
          <template #header>
            <div class="card-header">
              <span>操作</span>
            </div>
          </template>
          <div class="action-buttons">
            <el-button type="primary" :loading="importParsing" :disabled="importFileList.length === 0" @click="handleImport">
              导入并统计
            </el-button>
            <el-button :disabled="importedPoints.length === 0" @click="resetImport">重置</el-button>
            <el-button v-if="importedPoints.length > 0" type="success" @click="handleExportCsv">
              导出 CSV
            </el-button>
          </div>
        </el-card>

        <!-- 模块3: 案例统计 -->
        <el-card v-if="importedPoints.length > 0" class="import-stats-card" shadow="never">
          <template #header>
            <div class="card-header">
              <span>案例统计</span>
            </div>
          </template>

          <div class="summary-row">
            <div class="summary-item">
              <div class="summary-label">案例总数</div>
              <div class="summary-value">{{ importedPoints.length }}</div>
            </div>
            <div class="summary-item">
              <div class="summary-label">功能/步骤数</div>
              <div class="summary-value">{{ importStats.length }}</div>
            </div>
            <div class="summary-item">
              <div class="summary-label">规则简称案例</div>
              <div class="summary-value">{{ importAliasTotal }}</div>
            </div>
            <div class="summary-item">
              <div class="summary-label">类型分布</div>
              <div class="summary-value">
                {{ importTypeCounts.process }}/{{ importTypeCounts.rule }}/{{ importTypeCounts.page_control }}
              </div>
            </div>
            <div class="summary-item">
              <div class="summary-label">自动化</div>
              <div class="summary-value">{{ importAutomationTotal }}</div>
            </div>
          </div>

          <el-table
            :data="importStats"
            row-key="rowKey"
            :tree-props="{ children: 'children' }"
            border
            stripe
            style="width: 100%; margin-top: 16px;"
          >
            <el-table-column prop="component" label="组件" width="120" align="center" v-if="false" />
            <el-table-column prop="function" label="功能/步骤" min-width="220">
              <template #default="{ row }">
                <span>{{ row.function }}</span>
                <el-tag v-if="row.category === 'rule_alias'" size="small" type="info" class="alias-tag">规则简称</el-tag>
              </template>
            </el-table-column>
            <el-table-column prop="count" label="案例数" width="100" align="center" />
            <el-table-column prop="process" label="流程" width="80" align="center" />
            <el-table-column prop="rule" label="规则" width="80" align="center" />
            <el-table-column prop="page_control" label="页面" width="100" align="center" />
            <el-table-column prop="positive" label="正例" width="80" align="center" />
            <el-table-column prop="negative" label="反例" width="80" align="center" />
            <el-table-column label="优先级 (高/中/低)" width="140" align="center">
              <template #default="{ row }">
                {{ row.priority1 }} / {{ row.priority2 }} / {{ row.priority3 }}
              </template>
            </el-table-column>
            <el-table-column prop="automated" label="自动化" width="80" align="center" />
          </el-table>
        </el-card>
      </el-tab-pane>

    </el-tabs>
  </div>
</template>

<script setup>
import { ref, computed, onBeforeUnmount } from 'vue'
import { ElMessage } from 'element-plus'
import { UploadFilled } from '@element-plus/icons-vue'
import {
  parseXmind,
  previewGenerate,
  confirmPreview,
  bulkGenerate,
  getGenerationStatus,
  getGenerationStatusBySession,
  exportCases,
  exportCasesBySession,
  exportCasesBySessionWithHeaders
} from '../utils/api'

const uploadRef = ref(null)
const fileList = ref([])
const parsing = ref(false)
const typeMap = { process: '流程', rule: '规则', page_control: '页面' }
const subtypeMap = { positive: '正例', negative: '反例' }
const priorityMap = { 1: '高', 2: '中', 3: '低', '1': '高', '2': '中', '3': '低' }

const previewing = ref(false)
const parsedData = ref(null)
const previewCases = ref([])
const previewId = ref('')
const generationTaskId = ref('')
const generationStatus = ref(null)
const currentSessionId = ref('')
const strategy = ref('standard')
const exportLoading = ref(false)
const activeTab = ref('generate')
const sessionIdInput = ref('')
const sessionExportLoading = ref(false)

let pollTimer = null

// 导入案例 tab 状态
const importUploadRef = ref(null)
const importFileList = ref([])
const importParsing = ref(false)
const importedPoints = ref([])
const importFileName = ref('')
const importRequirementName = ref('')

const typeLabelMap = {
  process: '业务流程',
  rule: '业务规则',
  page_control: '页面'
}

const TYPE_SEGMENTS = ['业务流程', '业务规则', '页面控制']

// 解析 context 路径，定位段落节点（业务流程/业务规则/页面控制）的位置
const resolveContextPath = (context) => {
  const parts = context ? String(context).split(' / ').filter(Boolean) : []
  let sectionIndex = -1
  for (let i = parts.length - 1; i >= 0; i--) {
    if (TYPE_SEGMENTS.includes(parts[i])) {
      sectionIndex = i
      break
    }
  }
  return { parts, sectionIndex }
}

// 规则简称识别：优先取后端解析字段 rule_alias；
// 兼容旧 JSON（无该字段）时，取段落节点后的第一个路径段作为简称
const resolveRuleAlias = (point) => {
  if (point.rule_alias) return point.rule_alias
  const { parts, sectionIndex } = resolveContextPath(point.context || '')
  if (sectionIndex >= 0 && sectionIndex + 1 < parts.length) {
    return parts[sectionIndex + 1]
  }
  return null
}

const newStatsRow = (rowKey, name, component, category) => ({
  rowKey,
  function: name,
  component: component || '',
  category,
  count: 0,
  process: 0, rule: 0, page_control: 0,
  positive: 0, negative: 0,
  priority1: 0, priority2: 0, priority3: 0,
  automated: 0
})

const accumulateStats = (row, point) => {
  row.count++
  if (row[point.point_type] !== undefined) row[point.point_type]++
  if (point.subtype === 'positive') row.positive++
  else if (point.subtype === 'negative') row.negative++
  if (point.priority === 1) row.priority1++
  else if (point.priority === 2) row.priority2++
  else if (point.priority === 3) row.priority3++
  if (point.is_automated) row.automated++
}

// 案例统计：功能/步骤为主行，规则简称（带"联系"标注）归属其下，
// 作为可展开的子行（tree 表格，默认收起），统计上与功能步骤分开计数
const importStats = computed(() => {
  const points = importedPoints.value
  if (!points.length) return []
  const groups = {}
  const groupOrder = []

  const ensureGroup = (groupKey, component) => {
    if (!groups[groupKey]) {
      const row = newStatsRow(`g_${groupOrder.length}_${groupKey}`, groupKey, component, 'function_step')
      row._aliases = {}
      groups[groupKey] = row
      groupOrder.push(groupKey)
    }
    return groups[groupKey]
  }

  for (const point of points) {
    const context = point.context || ''
    const alias = resolveRuleAlias(point)

    let groupKey
    let component = ''
    if (alias) {
      const { parts, sectionIndex } = resolveContextPath(context)
      if (sectionIndex >= 1) {
        groupKey = parts[sectionIndex - 1]
        component = parts[1] || ''
      } else if (sectionIndex === 0) {
        groupKey = parts[0]
        component = parts[1] || ''
      } else {
        groupKey = alias
      }
    } else {
      const pathParts = context ? context.split(' / ').filter(Boolean) : []
      if (pathParts.length >= 2) {
        const last = pathParts[pathParts.length - 1]
        if (TYPE_SEGMENTS.includes(last)) {
          groupKey = pathParts[pathParts.length - 2]
        } else {
          groupKey = last
        }
        component = pathParts[1] || ''
      } else if (pathParts.length === 1) {
        groupKey = pathParts[0]
      } else {
        groupKey = typeLabelMap[point.point_type] || point.point_type || '未分类'
      }
    }

    const g = ensureGroup(groupKey, component)
    // 主行（功能/步骤）合计全部案例（含其下规则简称），
    // 案例数与流程/规则/页面/正例/反例/优先级/自动化均为总数
    accumulateStats(g, point)
    if (!alias) {
      continue
    }

    // 规则简称子行只统计业务规则案例
    if (point.point_type !== 'rule') {
      continue
    }
    if (!g._aliases[alias]) {
      g._aliases[alias] = newStatsRow(
        `${g.rowKey}_a_${Object.keys(g._aliases).length}`,
        alias,
        g.component,
        'rule_alias'
      )
    }
    accumulateStats(g._aliases[alias], point)
  }

  return groupOrder.map((key) => {
    const g = groups[key]
    const children = Object.values(g._aliases)
    const row = { ...g }
    delete row._aliases
    if (children.length) {
      row.children = children
    }
    return row
  })
})

// 规则简称案例数（子行合计，仅业务规则案例）
const importAliasTotal = computed(() => {
  let total = 0
  for (const row of importStats.value) {
    for (const child of row.children || []) {
      total += child.count
    }
  }
  return total
})

const importTypeCounts = computed(() => {
  const points = importedPoints.value
  const counts = { process: 0, rule: 0, page_control: 0 }
  for (const point of points) {
    if (counts[point.point_type] !== undefined) counts[point.point_type]++
  }
  return counts
})

const importAutomationTotal = computed(() => {
  const points = importedPoints.value
  let automated = 0
  for (const point of points) {
    if (point.is_automated) automated++
  }
  return automated
})

const prioritySummary = computed(() => {
  const byPriority = parsedData.value?.stats?.by_priority || {}
  return `高 ${byPriority['1'] || 0} / 中 ${byPriority['2'] || 0} / 低 ${byPriority['3'] || 0}`
})

const subtypeSummary = computed(() => {
  const bySubtype = parsedData.value?.stats?.by_subtype || {}
  return `正例 ${bySubtype.positive || 0} / 反例 ${bySubtype.negative || 0}`
})

const formatProgressMessage = (status) => {
  if (!status) {
    return '生成处理中，请稍后再试'
  }
  const percentage = Math.round((status.progress || 0) * 100)
  const completed = status.completed || 0
  const total = status.total || 0
  return `生成处理中：${completed}/${total}（${percentage}%）`
}

const ensureSessionCompleted = async (sessionId) => {
  const status = await getGenerationStatusBySession(sessionId)
  generationStatus.value = status
  if (status?.status === 'completed') {
    return true
  }
  if (status?.status === 'failed') {
    ElMessage.error(status.error || '生成失败')
    return false
  }
  ElMessage.warning(formatProgressMessage(status))
  return false
}

const handleFileChange = (file, files) => {
  fileList.value = files || []
}

const handleFileRemove = () => {
  if (parsing.value || previewing.value) {
    return false
  }
}

const handleParse = async () => {
  const file = uploadRef.value?.fileList?.[0]?.raw || fileList.value?.[0]?.raw
  if (!file) {
    ElMessage.warning('请先上传XMind文件')
    return
  }
  parsing.value = true
  try {
    const res = await parseXmind(file)
    if (!res.success) {
      throw new Error(res.message || '解析失败')
    }
    parsedData.value = res.data
    previewCases.value = []
    previewId.value = ''
    generationTaskId.value = ''
    generationStatus.value = null
    currentSessionId.value = ''
    ElMessage.success('解析成功')
  } catch (error) {
    ElMessage.error(error.message || '解析失败')
  } finally {
    parsing.value = false
  }
}

const handlePreview = async () => {
  if (!parsedData.value?.parse_id) {
    ElMessage.warning('请先完成解析')
    return
  }
  previewing.value = true
  try {
    const res = await previewGenerate(parsedData.value.parse_id, 4)
    if (!res.success) {
      throw new Error(res.message || '预生成失败')
    }
    previewId.value = res.preview_id
    previewCases.value = res.cases || []
    ElMessage.success('预生成完成')
  } catch (error) {
    ElMessage.error(error.message || '预生成失败')
  } finally {
    previewing.value = false
  }
}

const handleConfirm = async () => {
  if (!previewId.value) {
    ElMessage.warning('请先完成预生成')
    return
  }
  try {
    const res = await confirmPreview(previewId.value, strategy.value)
    if (!res.success) {
      throw new Error(res.message || '任务提交失败')
    }
    generationTaskId.value = res.task_id
    currentSessionId.value = res.session_id || ''
    startPolling(res.task_id)
    ElMessage.success(`生成任务已提交${res.session_id ? `，session_id: ${res.session_id}` : ''}`)
  } catch (error) {
    ElMessage.error(error.message || '任务提交失败')
  }
}

const handleBulkGenerate = async () => {
  if (!parsedData.value?.parse_id) {
    ElMessage.warning('请先完成解析')
    return
  }
  try {
    const res = await bulkGenerate(parsedData.value.parse_id, strategy.value)
    if (!res.success) {
      throw new Error(res.message || '任务提交失败')
    }
    generationTaskId.value = res.task_id
    currentSessionId.value = res.session_id || ''
    startPolling(res.task_id)
    ElMessage.success(`生成任务已提交${res.session_id ? `，session_id: ${res.session_id}` : ''}`)
  } catch (error) {
    ElMessage.error(error.message || '任务提交失败')
  }
}

const startPolling = async (taskId) => {
  stopPolling()
  const poll = async () => {
    try {
      const status = currentSessionId.value
        ? await getGenerationStatusBySession(currentSessionId.value)
        : await getGenerationStatus(taskId)
      generationStatus.value = status
      if (status.session_id) {
        currentSessionId.value = status.session_id
      }
      if (status.status === 'completed' || status.status === 'failed') {
        stopPolling()
        return
      }
    } catch (error) {
      ElMessage.error(error.message || '查询任务失败')
      stopPolling()
      return
    }
    pollTimer = setTimeout(poll, 2000)
  }
  poll()
}

const stopPolling = () => {
  if (pollTimer) {
    clearTimeout(pollTimer)
    pollTimer = null
  }
}

const resetAll = () => {
  parsedData.value = null
  previewCases.value = []
  previewId.value = ''
  generationTaskId.value = ''
  generationStatus.value = null
  currentSessionId.value = ''
  fileList.value = []
  if (uploadRef.value) {
    uploadRef.value.clearFiles()
  }
  stopPolling()
}

const removePreview = (index) => {
  previewCases.value.splice(index, 1)
}

const editDialogVisible = ref(false)
const editForm = ref({
  preconditions: '',
  steps: '',
  expected_results: ''
})
let editingIndex = -1

const openEdit = (row, index) => {
  editingIndex = index
  editForm.value = {
    preconditions: (row.preconditions || []).join('\n'),
    steps: (row.steps || []).join('\n'),
    expected_results: (row.expected_results || []).join('\n')
  }
  editDialogVisible.value = true
}

const saveEdit = () => {
  if (editingIndex < 0) return
  const row = previewCases.value[editingIndex]
  row.preconditions = editForm.value.preconditions.split('\n').filter(v => v.trim())
  row.steps = editForm.value.steps.split('\n').filter(v => v.trim())
  row.expected_results = editForm.value.expected_results.split('\n').filter(v => v.trim())
  editDialogVisible.value = false
}

const handleExport = async () => {
  exportLoading.value = true
  try {
    if (currentSessionId.value) {
      const ready = await ensureSessionCompleted(currentSessionId.value)
      if (!ready) {
        return
      }
      const response = await exportCasesBySessionWithHeaders(currentSessionId.value)
      const blob = response.data
      const filename = resolveDownloadName(
        response.headers?.['content-disposition'],
        parsedData.value?.requirement_name
      )
      const url = window.URL.createObjectURL(blob)
      const link = document.createElement('a')
      link.href = url
      link.download = filename
      document.body.appendChild(link)
      link.click()
      document.body.removeChild(link)
      window.URL.revokeObjectURL(url)
      return
    }
    if (!generationStatus.value?.cases?.length) {
      ElMessage.warning('暂无可导出的用例')
      return
    }
    const blob = await exportCases(parsedData.value?.requirement_name || '测试用例', generationStatus.value.cases)
    const url = window.URL.createObjectURL(blob)
    const link = document.createElement('a')
    link.href = url
    link.download = `${parsedData.value?.requirement_name || '测试用例'}.xmind`
    document.body.appendChild(link)
    link.click()
    document.body.removeChild(link)
    window.URL.revokeObjectURL(url)
  } catch (error) {
    ElMessage.error(error.message || '导出失败')
  } finally {
    exportLoading.value = false
  }
}

const handleExportBySession = async () => {
  if (!sessionIdInput.value) {
    ElMessage.warning('请输入 session_id')
    return
  }
  sessionExportLoading.value = true
  try {
    const sessionId = sessionIdInput.value.trim()
    const ready = await ensureSessionCompleted(sessionId)
    if (!ready) {
      return
    }
    const response = await exportCasesBySessionWithHeaders(sessionId)
    const blob = response.data
    const filename = resolveDownloadName(
      response.headers?.['content-disposition'],
      parsedData.value?.requirement_name
    )
    const url = window.URL.createObjectURL(blob)
    const link = document.createElement('a')
    link.href = url
    link.download = filename
    document.body.appendChild(link)
    link.click()
    document.body.removeChild(link)
    window.URL.revokeObjectURL(url)
  } catch (error) {
    ElMessage.error(error.message || '导出失败')
  } finally {
    sessionExportLoading.value = false
  }
}

const handleImportFileChange = (file, files) => {
  importFileList.value = files || []
}

const handleImportFileRemove = () => {
  if (importParsing.value) return false
}

const handleImport = async () => {
  const file = importUploadRef.value?.fileList?.[0]?.raw || importFileList.value?.[0]?.raw
  if (!file) {
    ElMessage.warning('请先选择文件')
    return
  }
  importParsing.value = true
  try {
    const name = file.name.toLowerCase()
    if (name.endsWith('.json')) {
      const text = await file.text()
      const data = JSON.parse(text)
      if (Array.isArray(data)) {
        importedPoints.value = data
      } else if (data?.test_points) {
        importedPoints.value = data.test_points
      } else if (data?.cases) {
        importedPoints.value = data.cases
      } else {
        ElMessage.error('无法识别的 JSON 格式，请提供案例数组或包含 test_points/cases 字段的对象')
        return
      }
      importFileName.value = file.name
      importRequirementName.value = file.name.replace(/\.json$/i, '')
      ElMessage.success(`成功导入 ${importedPoints.value.length} 条案例`)
    } else if (name.endsWith('.xmind')) {
      const res = await parseXmind(file)
      if (!res.success) {
        throw new Error(res.message || '解析失败')
      }
      importedPoints.value = res.data?.test_points || []
      importFileName.value = file.name
      importRequirementName.value = res.data?.requirement_name || ''
      if (res.data?.stats) {
        ElMessage.success(`成功解析 ${res.data.stats.total || importedPoints.value.length} 个测试点`)
      } else {
        ElMessage.success(`成功解析 ${importedPoints.value.length} 个测试点`)
      }
    } else {
      ElMessage.error('不支持的文件格式，请上传 .xmind 或 .json 文件')
    }
  } catch (error) {
    ElMessage.error(error.message || '导入失败')
  } finally {
    importParsing.value = false
  }
}

const resetImport = () => {
  importedPoints.value = []
  importFileName.value = ''
  importRequirementName.value = ''
  importFileList.value = []
  if (importUploadRef.value) {
    importUploadRef.value.clearFiles()
  }
}


// ---------- CSV / ZIP 导出工具 ----------
const csvEscape = (v) => {
  const s = String(v)
  if (s.includes(',') || s.includes('"') || s.includes('\n')) {
    return '"' + s.replace(/"/g, '""') + '"'
  }
  return s
}

// withAliases=false：仅功能/步骤行；true：功能/步骤行 + 规则简称子行（全展开）
const buildStatsCsv = (rows, withAliases) => {
  const statHeaders = ['案例数', '流程', '规则', '页面', '正例', '反例', '优先级(高)', '优先级(中)', '优先级(低)', '自动化']
  const headers = withAliases
    ? ['组件', '功能/步骤', '类别', ...statHeaders]
    : ['组件', '功能/步骤', ...statHeaders]
  const statValues = (row) => [
    row.count, row.process, row.rule, row.page_control,
    row.positive, row.negative,
    row.priority1, row.priority2, row.priority3,
    row.automated
  ]
  const lines = [headers.join(',')]
  for (const row of rows) {
    const base = [row.component || '', row.function || '']
    lines.push((withAliases ? [...base, '功能步骤', ...statValues(row)] : [...base, ...statValues(row)]).map(csvEscape).join(','))
    if (withAliases) {
      for (const child of row.children || []) {
        lines.push([
          row.component || '',
          `${row.function || ''} / ${child.function || ''}`,
          '规则简称',
          ...statValues(child)
        ].map(csvEscape).join(','))
      }
    }
  }
  return '\ufeff' + lines.join('\n')
}

const CRC_TABLE = (() => {
  const table = new Uint32Array(256)
  for (let n = 0; n < 256; n++) {
    let c = n
    for (let k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1)
    table[n] = c >>> 0
  }
  return table
})()

const crc32 = (bytes) => {
  let crc = 0xFFFFFFFF
  for (let i = 0; i < bytes.length; i++) {
    crc = (CRC_TABLE[(crc ^ bytes[i]) & 0xFF] ^ (crc >>> 8)) >>> 0
  }
  return (crc ^ 0xFFFFFFFF) >>> 0
}

// 生成 ZIP（store 方式，不压缩）：files = [{ name, data: Uint8Array }]
const buildZip = (files) => {
  const encoder = new TextEncoder()
  const now = new Date()
  const dosTime = ((now.getHours() << 11) | (now.getMinutes() << 5) | (now.getSeconds() >> 1)) & 0xFFFF
  const dosDate = (((now.getFullYear() - 1980) << 9) | ((now.getMonth() + 1) << 5) | now.getDate()) & 0xFFFF
  const chunks = []
  const centralChunks = []
  let offset = 0
  let centralSize = 0

  for (const file of files) {
    const nameBytes = encoder.encode(file.name)
    const data = file.data
    const crc = crc32(data)

    const local = new DataView(new ArrayBuffer(30))
    local.setUint32(0, 0x04034B50, true)
    local.setUint16(4, 20, true)          // 解压所需版本
    local.setUint16(6, 0x0800, true)      // 文件名 UTF-8
    local.setUint16(8, 0, true)           // store 不压缩
    local.setUint16(10, dosTime, true)
    local.setUint16(12, dosDate, true)
    local.setUint32(14, crc, true)
    local.setUint32(18, data.length, true)
    local.setUint32(22, data.length, true)
    local.setUint16(26, nameBytes.length, true)
    local.setUint16(28, 0, true)
    chunks.push(new Uint8Array(local.buffer), nameBytes, data)

    const central = new DataView(new ArrayBuffer(46))
    central.setUint32(0, 0x02014B50, true)
    central.setUint16(4, 20, true)
    central.setUint16(6, 20, true)
    central.setUint16(8, 0x0800, true)
    central.setUint16(10, 0, true)
    central.setUint16(12, dosTime, true)
    central.setUint16(14, dosDate, true)
    central.setUint32(16, crc, true)
    central.setUint32(20, data.length, true)
    central.setUint32(24, data.length, true)
    central.setUint16(28, nameBytes.length, true)
    central.setUint16(30, 0, true)
    central.setUint16(32, 0, true)
    central.setUint16(34, 0, true)
    central.setUint16(36, 0, true)
    central.setUint32(38, 0, true)
    central.setUint32(42, offset, true)
    centralChunks.push(new Uint8Array(central.buffer), nameBytes)

    offset += 30 + nameBytes.length + data.length
    centralSize += 46 + nameBytes.length
  }

  const eocd = new DataView(new ArrayBuffer(22))
  eocd.setUint32(0, 0x06054B50, true)
  eocd.setUint16(4, 0, true)
  eocd.setUint16(6, 0, true)
  eocd.setUint16(8, files.length, true)
  eocd.setUint16(10, files.length, true)
  eocd.setUint32(12, centralSize, true)
  eocd.setUint32(16, offset, true)
  eocd.setUint16(20, 0, true)

  const out = new Uint8Array(offset + centralSize + 22)
  let pos = 0
  for (const chunk of [...chunks, ...centralChunks, new Uint8Array(eocd.buffer)]) {
    out.set(chunk, pos)
    pos += chunk.length
  }
  return out
}

const downloadBlob = (blob, filename) => {
  const url = URL.createObjectURL(blob)
  const link = document.createElement('a')
  link.href = url
  link.download = filename
  document.body.appendChild(link)
  link.click()
  document.body.removeChild(link)
  URL.revokeObjectURL(url)
}

const handleExportCsv = () => {
  const rows = importStats.value
  if (!rows.length) {
    ElMessage.warning('暂无统计数据可导出')
    return
  }
  const baseName = importRequirementName.value || importFileName.value.replace(/\.[^.]+$/, '') || '案例统计'
  const ts = formatTimestamp()
  const encoder = new TextEncoder()
  const hasAlias = rows.some((row) => (row.children || []).length > 0)

  // 无规则简称：单份 CSV
  if (!hasAlias) {
    const blob = new Blob([buildStatsCsv(rows, false)], { type: 'text/csv;charset=utf-8;' })
    downloadBlob(blob, `${baseName}_案例统计_${ts}.csv`)
    ElMessage.success('CSV 导出成功')
    return
  }

  // 有规则简称：两份 CSV 打包为 ZIP
  //  1) 仅功能/步骤（数量为含规则简称的合计）
  //  2) 全展开（功能/步骤 + 规则简称子行）
  const zipBytes = buildZip([
    { name: `${baseName}_仅功能步骤_${ts}.csv`, data: encoder.encode(buildStatsCsv(rows, false)) },
    { name: `${baseName}_含规则简称全展开_${ts}.csv`, data: encoder.encode(buildStatsCsv(rows, true)) }
  ])
  downloadBlob(new Blob([zipBytes], { type: 'application/zip' }), `${baseName}_案例统计_${ts}.zip`)
  ElMessage.success('已导出压缩包：含 仅功能步骤 与 含规则简称全展开 两份 CSV')
}

onBeforeUnmount(() => {
  stopPolling()
})

const resolveDownloadName = (contentDisposition, requirementName) => {
  const filenameFromHeader = parseFilename(contentDisposition)
  if (filenameFromHeader) {
    return filenameFromHeader
  }
  const name = requirementName || '测试用例'
  return `测试用例_${name}_${formatTimestamp()}.xmind`
}

const parseFilename = (contentDisposition) => {
  if (!contentDisposition) return ''
  const utf8Match = contentDisposition.match(/filename\*\=UTF-8''([^;]+)/i)
  if (utf8Match && utf8Match[1]) {
    try {
      return decodeURIComponent(utf8Match[1])
    } catch {
      return utf8Match[1]
    }
  }
  const plainMatch = contentDisposition.match(/filename=([^;]+)/i)
  if (plainMatch && plainMatch[1]) {
    return plainMatch[1].replace(/\"/g, '')
  }
  return ''
}

const formatTimestamp = () => {
  const now = new Date()
  const pad = (v) => String(v).padStart(2, '0')
  return `${now.getFullYear()}${pad(now.getMonth() + 1)}${pad(now.getDate())}${pad(now.getHours())}${pad(now.getMinutes())}${pad(now.getSeconds())}`
}
</script>

<style scoped>
.case-container {
  width: 92%;
  max-width: 1200px;
  margin: 0 auto;
}

.case-container.expanded {
  width: 98%;
  max-width: 1400px;
}

.case-tabs {
  width: 100%;
}

/* 统一 tab 头部样式：等宽、居中、一致的间距 */
:deep(.case-tabs .el-tabs__header) {
  margin-bottom: 24px;
}

:deep(.case-tabs .el-tabs__nav) {
  width: 100%;
  display: flex;
}

:deep(.case-tabs .el-tabs__item) {
  flex: 1;
  text-align: center;
  font-size: 15px;
  font-weight: 500;
  padding: 0 !important;
  height: 44px;
  line-height: 44px;
  transition: color 0.2s ease;
}

:deep(.case-tabs .el-tabs__active-bar) {
  height: 3px;
  border-radius: 2px;
}

.upload-card,
.overview-card,
.preview-card,
.preview-table-card,
.confirm-card,
.progress-card,
.export-card,
.session-card,
.import-file-card,
.import-action-card {
  margin-bottom: 20px;
}

.card-header {
  font-weight: 600;
}

.action-buttons {
  margin-top: 16px;
  display: flex;
  gap: 10px;
}

.stat-item {
  padding: 12px;
  background: #f5f7fa;
  border-radius: 6px;
  display: flex;
  flex-direction: column;
}

.stat-stack {
  display: flex;
  flex-direction: column;
  gap: 12px;
}

.stat-title {
  font-size: 13px;
  color: #909399;
  margin-bottom: 6px;
}

.stat-value {
  font-size: 14px;
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
}

.strategy-text {
  margin-bottom: 12px;
  color: #606266;
}

.confirm-actions {
  display: flex;
  align-items: center;
  gap: 12px;
}

.session-actions {
  display: flex;
  align-items: center;
  gap: 12px;
  flex-wrap: wrap;
}

.session-input {
  flex: 0 0 420px;
  max-width: 420px;
}

.session-input :deep(.el-input__inner) {
  overflow: hidden;
  text-overflow: ellipsis;
}

.progress-meta {
  margin-top: 10px;
  display: flex;
  gap: 20px;
  flex-wrap: wrap;
}

.log-panel {
  max-height: 200px;
  overflow-y: auto;
}

.log-title {
  font-weight: 600;
  margin-bottom: 8px;
}

.log-item {
  font-size: 12px;
  color: #606266;
  margin-bottom: 4px;
}

.multi-line {
  white-space: pre-line;
  font-size: 12px;
}

.import-stats-card {
  margin-bottom: 20px;
}

.import-action-card :deep(.el-card__body) {
  padding: 16px 20px;
}

.import-action-card .action-buttons {
  margin-top: 0;
}

.summary-row {
  display: flex;
  gap: 24px;
  flex-wrap: wrap;
}

.summary-item {
  padding: 12px 20px;
  background: #f5f7fa;
  border-radius: 6px;
  flex: 1;
  min-width: 120px;
}

.summary-label {
  font-size: 13px;
  color: #909399;
  margin-bottom: 6px;
  white-space: nowrap;
}

.summary-value {
  font-size: 18px;
  font-weight: 600;
  color: #303133;
  word-break: break-all;
}

.alias-tag {
  margin-left: 8px;
}

</style>
