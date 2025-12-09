<template>
  <div class="home-container">
    <el-card class="upload-card" shadow="hover">
      <template #header>
        <div class="card-header">
          <span>上传Word文档</span>
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
        :limit="5"
        accept=".doc,.docx"
        :disabled="loading || generating || hasProcessed"
        multiple
      >
        <el-icon class="el-icon--upload"><upload-filled /></el-icon>
        <div class="el-upload__text">
          将文件拖到此处，或<em>点击上传</em>
        </div>
        <template #tip>
          <div class="el-upload__tip">
            支持 .doc 和 .docx 格式的Word文档，最多可同时上传5个文件
            <span v-if="hasProcessed" style="color: #f56c6c;">（已处理，请先取消后再导入）</span>
          </div>
        </template>
      </el-upload>
      
      <!-- 处理进度 -->
      <div v-if="processingFiles.length > 0" class="processing-status">
        <el-progress
          :percentage="processingProgress"
          :status="processingStatus"
        />
        <p class="processing-text">{{ processingText }}</p>
      </div>
    </el-card>
    
    <!-- 操作按钮区域 - 放在卡片外部 -->
    <div class="action-buttons-container">
      <el-button
        type="primary"
        size="large"
        :loading="loading || generating"
        :disabled="fileCount === 0 || loading || generating || hasProcessed"
        @click="handleParseAndGenerate"
      >
        <el-icon><Document /></el-icon>
        解析并生成测试大纲
      </el-button>
      <el-button
        v-if="parsedDataList.length > 0"
        size="large"
        @click="handleDebug"
      >
        <el-icon><View /></el-icon>
        调试信息
      </el-button>
      <el-button
        v-if="hasProcessed"
        size="large"
        @click="handleCancel"
      >
        <el-icon><Close /></el-icon>
        取消
      </el-button>
    </div>

    <!-- 解析结果预览弹窗 -->
    <el-dialog
      v-model="previewDialogVisible"
      title="解析结果预览"
      width="60%"
      :close-on-click-modal="false"
      :close-on-press-escape="false"
      :show-close="false"
    >
      <div class="preview-dialog-content">
        <el-alert
          v-if="previewError"
          :title="previewError"
          type="error"
          :closable="true"
          @close="previewError = ''"
          show-icon
          style="margin-bottom: 20px;"
        />
        
        <el-collapse 
          v-for="(item, idx) in previewDataList" 
          :key="idx" 
          class="file-result-collapse"
          :model-value="[`file-${idx}`]"
        >
          <el-collapse-item :name="`file-${idx}`" :title="`文件 ${idx + 1}: ${item.file}`">
            <!-- 建模需求显示 -->
            <template v-if="item.data.document_type !== 'non_modeling'">
              <el-descriptions :column="2" border>
                <el-descriptions-item label="版本编号">
                  {{ item.data.version || '未提取' }}
                </el-descriptions-item>
                <el-descriptions-item label="用例名称">
                  {{ item.data.requirement_info?.case_name || '未提取' }}
                </el-descriptions-item>
                <el-descriptions-item label="客户（C）">
                  {{ item.data.requirement_info?.customer ?? '未提取' }}
                </el-descriptions-item>
                <el-descriptions-item label="产品（P）">
                  {{ item.data.requirement_info?.product ?? '未提取' }}
                </el-descriptions-item>
                <el-descriptions-item label="渠道（C）">
                  {{ item.data.requirement_info?.channel ?? '未提取' }}
                </el-descriptions-item>
                <el-descriptions-item label="合作方（P）">
                  {{ item.data.requirement_info?.partner ?? '未提取' }}
                </el-descriptions-item>
                <el-descriptions-item label="活动数量" :span="2">
                  {{ item.data.activities?.length || 0 }}
                </el-descriptions-item>
              </el-descriptions>

              <!-- 活动详情 -->
              <el-collapse 
                v-if="item.data.activities?.length" 
                class="activities-collapse"
                :model-value="getActivityKeys(item.data.activities, idx)"
              >
                <el-collapse-item
                  v-for="(activity, index) in item.data.activities"
                  :key="index"
                  :name="`file-${idx}-activity-${index}`"
                  :title="`活动 ${index + 1}: ${activity.name}`"
                >
                  <div
                    v-for="(component, cIndex) in activity.components"
                    :key="cIndex"
                    class="component-item"
                  >
                    <h4>组件: {{ component.name }}</h4>
                    <div
                      v-for="(task, tIndex) in component.tasks"
                      :key="tIndex"
                      class="task-item"
                    >
                      <h5>任务: {{ task.name }}</h5>
                      <div
                        v-for="(step, sIndex) in task.steps"
                        :key="sIndex"
                        class="step-item"
                      >
                        <p><strong>步骤:</strong> {{ step.name }}</p>
                        <p v-if="step.input_elements?.length">
                          <strong>输入要素:</strong> {{ step.input_elements.length }} 个
                        </p>
                        <p v-if="step.output_elements?.length">
                          <strong>输出要素:</strong> {{ step.output_elements.length }} 个
                        </p>
                      </div>
                    </div>
                  </div>
                </el-collapse-item>
              </el-collapse>
            </template>

            <!-- 非建模需求显示 -->
            <template v-else>
              <el-descriptions :column="2" border>
                <el-descriptions-item label="文档类型">
                  非建模需求
                </el-descriptions-item>
                <el-descriptions-item label="文件编号">
                  {{ item.data.file_number || '未提取' }}
                </el-descriptions-item>
                <el-descriptions-item label="文件名称" :span="2">
                  {{ item.data.file_name || '未提取' }}
                </el-descriptions-item>
                <el-descriptions-item label="需求名称" :span="2">
                  {{ item.data.requirement_name || '未提取' }}
                </el-descriptions-item>
                <el-descriptions-item label="设计者" :span="2">
                  {{ item.data.designer || '未提取' }}
                </el-descriptions-item>
                <el-descriptions-item label="功能数量" :span="2">
                  {{ item.data.functions?.length || 0 }}
                </el-descriptions-item>
              </el-descriptions>

              <!-- 功能详情 -->
              <el-collapse 
                v-if="item.data.functions?.length" 
                class="activities-collapse"
                :model-value="getFunctionKeys(item.data.functions, idx)"
              >
                <el-collapse-item
                  v-for="(func, index) in item.data.functions"
                  :key="index"
                  :name="`file-${idx}-function-${index}`"
                  :title="`功能 ${index + 1}: ${func.name}`"
                >
                  <div class="function-item">
                    <p>
                      <strong>输入要素:</strong> {{ func.input_elements?.length || 0 }} 个
                    </p>
                    <p>
                      <strong>输出要素:</strong> {{ func.output_elements?.length || 0 }} 个
                    </p>
                  </div>
                </el-collapse-item>
              </el-collapse>
            </template>
          </el-collapse-item>
        </el-collapse>
      </div>
      
      <template #footer>
        <div class="dialog-footer">
          <el-button @click="handlePreviewCancel">取消</el-button>
          <el-button type="primary" :loading="generating" @click="handlePreviewConfirm">
            确认并生成
          </el-button>
        </div>
      </template>
    </el-dialog>
  </div>
</template>

<script setup>
import { ref, computed } from 'vue'
import { ElMessage } from 'element-plus'
import { UploadFilled, Document, Close, View } from '@element-plus/icons-vue'
import { parseDocument, generateOutlineFromJson } from '../utils/api'

const uploadRef = ref(null)
const fileList = ref([])
const loading = ref(false)
const generating = ref(false)
const parsedDataList = ref([]) // 存储多个文件的解析结果（用于调试信息）
const previewDataList = ref([]) // 预览数据列表（用于弹窗显示）
const previewDialogVisible = ref(false) // 预览弹窗显示状态
const previewError = ref('') // 预览错误信息
const hasProcessed = ref(false) // 是否已处理过文件
const processingFiles = ref([]) // 正在处理的文件列表
const currentProcessingIndex = ref(0) // 当前处理的文件索引
const pendingFiles = ref([]) // 待生成的文件数据

// 计算当前文件数量（用于按钮状态）
const fileCount = computed(() => {
  return uploadRef.value?.fileList?.length || fileList.value.length || 0
})

// 计算处理进度
const processingProgress = computed(() => {
  if (processingFiles.value.length === 0) return 0
  return Math.round((currentProcessingIndex.value / processingFiles.value.length) * 100)
})

const processingStatus = computed(() => {
  if (loading.value || generating.value) return null
  return 'success'
})

const processingText = computed(() => {
  if (processingFiles.value.length === 0) return ''
  const current = currentProcessingIndex.value + 1
  const total = processingFiles.value.length
  const currentFile = processingFiles.value[currentProcessingIndex.value]?.name || ''
  if (loading.value) {
    return `正在解析: ${currentFile} (${current}/${total})`
  } else if (generating.value) {
    return `正在生成: ${currentFile} (${current}/${total})`
  } else {
    return `处理完成: ${current}/${total}`
  }
})

const handleFileChange = (file, fileListParam) => {
  if (hasProcessed.value) {
    ElMessage.warning('请先取消当前处理，再导入新文件')
    // 延迟移除，避免立即触发
    setTimeout(() => {
      if (uploadRef.value) {
        uploadRef.value.handleRemove(file)
      }
    }, 100)
    return false
  }
  // 更新本地fileList引用，确保响应式更新
  fileList.value = fileListParam || []
  // 检查文件数量限制
  if (fileListParam && fileListParam.length > 5) {
    ElMessage.warning('最多只能上传5个文件')
    // 移除超出限制的文件
    setTimeout(() => {
      if (uploadRef.value) {
        const filesToRemove = fileListParam.slice(5)
        filesToRemove.forEach(f => {
          uploadRef.value.handleRemove(f)
        })
      }
    }, 100)
    return false
  }
  // 验证文件格式
  const validFiles = fileListParam.filter(f => f.raw && (f.name.endsWith('.doc') || f.name.endsWith('.docx')))
  if (validFiles.length !== fileListParam.length) {
    ElMessage.warning('只支持 .doc 和 .docx 格式的文件')
  }
}

const handleFileRemove = () => {
  if (hasProcessed.value) {
    ElMessage.warning('请先取消当前处理，再删除文件')
    return false
  }
}

// 解析并显示预览
const handleParseAndGenerate = async () => {
  // 获取当前文件列表（从upload组件）
  const currentFiles = uploadRef.value?.fileList || fileList.value
  if (!currentFiles || currentFiles.length === 0) {
    ElMessage.warning('请先上传文件')
    return
  }

  // 过滤有效文件（确保有raw属性）
  const validFiles = currentFiles.filter(f => {
    const file = f.raw || f
    return file && (file.name?.endsWith('.doc') || file.name?.endsWith('.docx'))
  })
  
  if (validFiles.length === 0) {
    ElMessage.warning('请上传有效的Word文档')
    return
  }

  hasProcessed.value = true
  processingFiles.value = validFiles
  previewDataList.value = []
  previewError.value = ''
  currentProcessingIndex.value = 0
  pendingFiles.value = []

  // 逐个解析文件，显示预览
  loading.value = true
  let hasError = false
  let firstErrorFile = ''
  let firstErrorMessage = ''
  
  for (let i = 0; i < validFiles.length; i++) {
    currentProcessingIndex.value = i
    const file = validFiles[i].raw || validFiles[i]

    try {
      // 解析文档
      const response = await parseDocument(file)
      
      if (!response.success) {
        throw new Error(response.message || '解析失败')
      }

      // 保存解析结果用于预览和后续生成
      const parsedItem = {
        file: validFiles[i].name,
        data: response.data
      }
      previewDataList.value.push(parsedItem)
      pendingFiles.value.push({
        file: validFiles[i].name,
        data: response.data
      })
    } catch (error) {
      // 记录第一个错误
      if (!hasError) {
        hasError = true
        firstErrorFile = validFiles[i].name
        firstErrorMessage = error.message || '解析失败'
      }
      
      // 如果发生错误，立即停止处理其他文件
      break
    }
  }

  loading.value = false

  // 如果发生错误，清除文件并显示错误信息
  if (hasError) {
    // 清空所有状态
    previewDataList.value = []
    pendingFiles.value = []
    hasProcessed.value = false
    processingFiles.value = []
    currentProcessingIndex.value = 0
    
    // 清除文件列表
    fileList.value = []
    if (uploadRef.value) {
      uploadRef.value.clearFiles()
    }
    
    // 显示错误信息（后端已经格式化为"文件名 解析失败：错误原因"的格式）
    // 如果后端返回的信息已经包含文件名，直接使用；否则添加文件名
    let errorMsg = firstErrorMessage
    if (!errorMsg.includes(firstErrorFile)) {
      errorMsg = `${firstErrorFile} 解析失败：${firstErrorMessage}`
    }
    
    // 使用 ElMessage 悬浮显示错误信息
    ElMessage.error({
      message: errorMsg,
      duration: 5000, // 5秒后自动消失
      showClose: true // 显示关闭按钮
    })
    
    // 如果有多个文件，提示其他文件未处理
    if (validFiles.length > 1) {
      ElMessage.warning({
        message: '由于第一个文件解析失败，其他文件已停止处理',
        duration: 5000,
        showClose: true
      })
    }
  } else if (previewDataList.value.length > 0) {
    // 所有文件解析成功，显示预览弹窗
    previewDialogVisible.value = true
  } else {
    // 没有文件解析成功（理论上不应该到这里，因为hasError会捕获）
    hasProcessed.value = false
    ElMessage.error('所有文件解析失败，请检查文件格式')
  }
}

// 预览确认，生成XMind文件
const handlePreviewConfirm = async () => {
  if (pendingFiles.value.length === 0) {
    ElMessage.warning('没有可生成的文件')
    return
  }

  generating.value = true
  let successCount = 0
  let failCount = 0

  // 生成并下载XMind文件
  for (const item of pendingFiles.value) {
    try {
      const blob = await generateOutlineFromJson(item.data)
      
      // 下载文件
      const url = window.URL.createObjectURL(blob)
      const link = document.createElement('a')
      link.href = url
      
      // 根据文档类型生成文件名
      let filename = '测试大纲.xmind'
      if (item.data.document_type === 'non_modeling') {
        // 非建模需求：需求名称-时间戳
        const requirementName = item.data.requirement_name || '测试大纲'
        const timestamp = new Date().toISOString().replace(/[-:]/g, '').replace(/\..+/, '').replace('T', '_').slice(0, 15)
        filename = `${requirementName}-${timestamp}.xmind`
      } else {
        // 建模需求：用例名称-版本号
        const caseName = item.data.requirement_info?.case_name || '测试大纲'
        const version = item.data.version || ''
        if (version) {
          filename = `${caseName}-${version}.xmind`
        } else {
          filename = `${caseName}.xmind`
        }
      }
      link.download = filename
      
      document.body.appendChild(link)
      link.click()
      document.body.removeChild(link)
      window.URL.revokeObjectURL(url)
      
      successCount++
    } catch (error) {
      failCount++
      ElMessage.error(`${item.file} 生成失败: ${error.message || '生成失败'}`)
    }
  }

  generating.value = false

  // 保存解析结果用于调试信息
  parsedDataList.value = [...pendingFiles.value]

  // 关闭弹窗，清空预览数据
  previewDialogVisible.value = false
  previewDataList.value = []
  pendingFiles.value = []
  
  // 清空文件列表
  fileList.value = []
  if (uploadRef.value) {
    uploadRef.value.clearFiles()
  }

  if (successCount > 0) {
    ElMessage.success(`成功生成 ${successCount} 个文件${failCount > 0 ? `，失败 ${failCount} 个` : ''}`)
  }
}

// 预览取消
const handlePreviewCancel = () => {
  previewDialogVisible.value = false
  previewDataList.value = []
  pendingFiles.value = []
  hasProcessed.value = false
  processingFiles.value = []
  currentProcessingIndex.value = 0
  previewError.value = ''
  // 清空文件列表
  fileList.value = []
  if (uploadRef.value) {
    uploadRef.value.clearFiles()
  }
  ElMessage.info('已取消生成')
}

// 取消处理，允许重新导入
const handleCancel = () => {
  hasProcessed.value = false
  parsedDataList.value = []
  previewDataList.value = []
  pendingFiles.value = []
  processingFiles.value = []
  currentProcessingIndex.value = 0
  previewError.value = ''
  loading.value = false
  generating.value = false
  previewDialogVisible.value = false
  // 清空文件列表
  fileList.value = []
  if (uploadRef.value) {
    uploadRef.value.clearFiles()
  }
  ElMessage.info('已取消，可以重新导入文件')
}

// 获取活动折叠面板的默认展开keys
const getActivityKeys = (activities, fileIndex) => {
  return activities.map((_, index) => `file-${fileIndex}-activity-${index}`)
}

// 获取功能折叠面板的默认展开keys
const getFunctionKeys = (functions, fileIndex) => {
  return functions.map((_, index) => `file-${fileIndex}-function-${index}`)
}

// 调试信息
const handleDebug = () => {
  if (parsedDataList.value.length === 0) {
    ElMessage.warning('没有解析数据')
    return
  }
  
  const debugInfo = JSON.stringify(parsedDataList.value, null, 2)
  console.log('=== 解析数据调试信息 ===')
  console.log(debugInfo)
  
  // 创建新窗口显示调试信息
  const newWindow = window.open('', '_blank')
  if (newWindow) {
    newWindow.document.write('<!DOCTYPE html><html><head><meta charset="UTF-8"><title>调试信息</title><style>body { font-family: monospace; padding: 20px; background: #f5f5f5; } pre { background: white; padding: 15px; border-radius: 5px; overflow: auto; }</style></head><body><h2>解析数据调试信息</h2><pre>' + debugInfo.replace(/</g, '&lt;').replace(/>/g, '&gt;') + '</pre></body></html>')
    newWindow.document.close()
  } else {
    ElMessage.warning('无法打开新窗口，请查看浏览器控制台')
  }
}
</script>

<style scoped>
.home-container {
  width: 100%;
  max-width: 100%;
  box-sizing: border-box;
  overflow-x: hidden;
}

.upload-card {
  margin-bottom: 20px;
  width: 100%;
  max-width: 100%;
  box-sizing: border-box;
  overflow-x: hidden;
}

/* 确保卡片内容不会超出 */
:deep(.upload-card .el-card__body) {
  width: 100%;
  max-width: 100%;
  box-sizing: border-box;
  overflow-x: hidden;
  display: flex;
  flex-direction: column;
}

.card-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  font-weight: 600;
  font-size: 16px;
}

.upload-demo {
  width: 100%;
}

/* 限制文件名显示长度，避免窗体大小变化 */
:deep(.el-upload-list__item) {
  width: 100%;
  max-width: 100%;
}

:deep(.el-upload-list__item-name) {
  max-width: 400px;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
  display: inline-block;
  vertical-align: middle;
}

:deep(.el-upload-list) {
  width: 100%;
}

/* 操作按钮容器 - 外部居中 */
.action-buttons-container {
  margin-top: 20px;
  display: flex;
  justify-content: center;
  gap: 15px;
  width: 100%;
  max-width: 100%;
  box-sizing: border-box;
  flex-wrap: wrap;
}

.preview-card {
  margin-top: 20px;
}

.activities-collapse {
  margin-top: 20px;
}

.component-item {
  margin: 15px 0;
  padding: 10px;
  background: #f5f7fa;
  border-radius: 4px;
}

.component-item h4 {
  margin: 0 0 10px 0;
  color: #409eff;
}

.task-item {
  margin: 10px 0 10px 20px;
  padding: 10px;
  background: white;
  border-radius: 4px;
}

.task-item h5 {
  margin: 0 0 8px 0;
  color: #67c23a;
}

.step-item {
  margin: 8px 0 8px 20px;
  padding: 8px;
  background: #fafafa;
  border-radius: 4px;
  font-size: 14px;
}

.step-item p {
  margin: 4px 0;
}

/* 错误提示 - 放在卡片内部顶部 */
.processing-status {
  margin-top: 20px;
  padding: 15px;
  background: #f5f7fa;
  border-radius: 4px;
}

.processing-text {
  margin-top: 10px;
  text-align: center;
  color: #606266;
  font-size: 14px;
}

.file-result-collapse {
  margin-top: 10px;
}

.preview-dialog-content {
  max-height: 70vh;
  overflow-y: auto;
}

.dialog-footer {
  display: flex;
  justify-content: flex-end;
  gap: 10px;
}
</style>

