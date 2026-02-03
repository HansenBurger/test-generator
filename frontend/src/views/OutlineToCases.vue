<template>
  <div class="case-container" :class="{ expanded: parsedData }">
    <el-tabs v-model="activeTab" class="case-tabs">
      <el-tab-pane label="生成用例" name="generate">
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
        accept=".xmind"
        :limit="1"
        :disabled="parsing || previewing"
      >
        <el-icon class="el-icon--upload"><upload-filled /></el-icon>
        <div class="el-upload__text">
          将XMind拖到此处，或<em>点击上传</em>
        </div>
        <template #tip>
          <div class="el-upload__tip">仅支持 .xmind 格式，命名规范：需求名_V版本_时间（仅支持单个文件）</div>
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
              <div class="stat-value">
                高 {{ parsedData.stats?.by_priority?.['1'] || 0 }} /
                中 {{ parsedData.stats?.by_priority?.['2'] || 0 }} /
                低 {{ parsedData.stats?.by_priority?.['3'] || 0 }}
              </div>
            </div>
            <div class="stat-item">
              <div class="stat-title">正/反例分布</div>
              <div class="stat-value">
                正例 {{ parsedData.stats?.by_subtype?.positive || 0 }} /
                反例 {{ parsedData.stats?.by_subtype?.negative || 0 }}
              </div>
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
        <el-table-column prop="point_type" label="类型" width="80" />
        <el-table-column prop="subtype" label="子类型" width="90" />
        <el-table-column prop="priority" label="优先级" width="80" />
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

      <el-tab-pane label="按Session导出" name="session">
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
              style="max-width: 420px"
            />
            <el-button type="primary" :loading="sessionExportLoading" @click="handleExportBySession">
              导出XMind
            </el-button>
          </div>
        </el-card>
      </el-tab-pane>
    </el-tabs>
  </div>
</template>

<script setup>
import { ref, onBeforeUnmount } from 'vue'
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
    const response = await exportCasesBySessionWithHeaders(sessionIdInput.value.trim())
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

.upload-card,
.overview-card,
.preview-card,
.preview-table-card,
.confirm-card,
.progress-card,
.export-card {
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
</style>
