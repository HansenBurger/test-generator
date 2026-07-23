<template>
  <el-dialog
    :model-value="modelValue"
    title="模型配置"
    width="520px"
    :close-on-click-modal="false"
    @update:model-value="$emit('update:modelValue', $event)"
    @open="load"
  >
    <div v-loading="loading">
      <el-form label-position="top">
        <el-form-item>
          <template #label>
            <span class="field-label">当前生效模型
              <el-tooltip placement="top" :show-after="150">
                <template #content><div class="tip">本次生成实际调用的模型。选择「auto」时，会在可选模型间自动切换。</div></template>
                <el-icon class="tip-icon"><QuestionFilled /></el-icon>
              </el-tooltip>
            </span>
          </template>
          <el-tag :type="modelMode === 'auto' ? 'warning' : 'success'" size="large">{{ effectiveLabel }}</el-tag>
        </el-form-item>

        <el-form-item>
          <template #label>
            <span class="field-label">模型
              <el-tooltip placement="top" :show-after="150">
                <template #content><div class="tip">选择具体模型则固定使用；选择 auto 则在可选模型间自动切换（某个失败或额度用尽时换下一个）；清空则使用默认模型。</div></template>
                <el-icon class="tip-icon"><QuestionFilled /></el-icon>
              </el-tooltip>
            </span>
          </template>
          <el-select v-model="modelPick" placeholder="使用默认模型" clearable style="width: 100%">
            <el-option label="auto（自动轮转）" value="__auto__" />
            <el-option v-for="m in availableModels" :key="m" :label="m" :value="m" />
          </el-select>
        </el-form-item>

        <el-form-item>
          <template #label>
            <span class="field-label">启用模型思考（thinking）
              <el-tooltip placement="top" :show-after="150">
                <template #content><div class="tip">开启后模型会先思考再作答，回答通常更细致，但更慢、更耗额度；关闭则更快。</div></template>
                <el-icon class="tip-icon"><QuestionFilled /></el-icon>
              </el-tooltip>
            </span>
          </template>
          <el-switch v-model="form.enable_thinking" />
        </el-form-item>

        <el-form-item>
          <template #label>
            <span class="field-label">思考 Token 余量
              <el-tooltip placement="top" :show-after="150">
                <template #content><div class="tip">开启思考时，模型会先用一部分额度思考。这里为思考预留额外空间，避免思考占满额度后没有正式回答。</div></template>
                <el-icon class="tip-icon"><QuestionFilled /></el-icon>
              </el-tooltip>
            </span>
          </template>
          <el-input-number
            v-model="form.thinking_token_buffer"
            :min="0"
            :max="32000"
            :step="100"
            controls-position="right"
          />
        </el-form-item>

        <el-form-item>
          <template #label>
            <span class="field-label">生成温度（temperature）
              <el-tooltip placement="top" :show-after="150">
                <template #content><div class="tip">控制回答的随机性：越低越稳定、越高越多样，仅影响单条用例的生成。不填则使用默认值（standard 0.2 / 其它 0.6）。</div></template>
                <el-icon class="tip-icon"><QuestionFilled /></el-icon>
              </el-tooltip>
            </span>
          </template>
          <div class="temp-row">
            <el-switch v-model="customTemp" active-text="自定义" inactive-text="默认" />
            <el-input-number
              v-if="customTemp"
              v-model="form.temperature"
              :min="0"
              :max="2"
              :step="0.1"
              :precision="1"
              controls-position="right"
              style="margin-left: 12px"
            />
          </div>
        </el-form-item>
      </el-form>
    </div>

    <template #footer>
      <el-button @click="resetDefaults">恢复默认</el-button>
      <el-button @click="$emit('update:modelValue', false)">取消</el-button>
      <el-button type="primary" :loading="saving" @click="save">保存</el-button>
    </template>
  </el-dialog>
</template>

<script setup>
import { reactive, ref, computed } from 'vue'
import { ElMessage } from 'element-plus'
import { getModelConfig, updateModelConfig } from '../utils/api'

defineProps({
  modelValue: { type: Boolean, default: false }
})
const emit = defineEmits(['update:modelValue', 'saved'])

const AUTO_MARK = '__auto__'

const loading = ref(false)
const saving = ref(false)
const availableModels = ref([])
const effectiveModel = ref('')
const defaultModel = ref('')
const modelMode = ref('default')
const customTemp = ref(false)
const defaults = reactive({
  enable_thinking: false,
  thinking_token_buffer: 1500,
  current_model: null,
  temperature: null,
  model_mode: 'default'
})

const form = reactive({
  enable_thinking: false,
  thinking_token_buffer: 1500,
  current_model: null,
  temperature: null,
  model_mode: 'default'
})

const effectiveLabel = computed(() => {
  if (modelMode.value === 'auto') return 'auto · 轮转'
  return effectiveModel.value || '-'
})

// 模型下拉三态映射：__auto__ / 具体模型 / 空(默认)
const modelPick = computed({
  get() {
    if (form.model_mode === 'auto') return AUTO_MARK
    if (form.model_mode === 'fixed' && form.current_model) return form.current_model
    return null
  },
  set(v) {
    if (v === AUTO_MARK) {
      form.model_mode = 'auto'
      form.current_model = null
    } else if (v === null || v === undefined || v === '') {
      form.model_mode = 'default'
      form.current_model = null
    } else {
      form.model_mode = 'fixed'
      form.current_model = v
    }
  }
})

const fillForm = (cfg) => {
  form.enable_thinking = !!cfg.enable_thinking
  form.thinking_token_buffer = cfg.thinking_token_buffer
  form.current_model = cfg.current_model || null
  form.temperature = cfg.temperature ?? null
  form.model_mode = cfg.model_mode || 'default'
  customTemp.value = cfg.temperature !== null && cfg.temperature !== undefined
}

const load = async () => {
  loading.value = true
  try {
    const data = await getModelConfig()
    availableModels.value = data.available_models || []
    effectiveModel.value = data.effective_model || ''
    defaultModel.value = data.default_model || ''
    modelMode.value = data.model_mode || 'default'
    if (data.defaults) {
      defaults.enable_thinking = !!data.defaults.enable_thinking
      defaults.thinking_token_buffer = data.defaults.thinking_token_buffer
      defaults.current_model = data.defaults.current_model || null
      defaults.temperature = data.defaults.temperature ?? null
      defaults.model_mode = data.defaults.model_mode || 'default'
    }
    fillForm(data)
  } catch (e) {
    ElMessage.error('加载模型配置失败：' + (e.message || e))
  } finally {
    loading.value = false
  }
}

const resetDefaults = () => {
  fillForm(defaults)
  ElMessage.info('已填入默认值，点击保存后生效')
}

const save = async () => {
  saving.value = true
  try {
    const data = await updateModelConfig({
      enable_thinking: !!form.enable_thinking,
      thinking_token_buffer: Number(form.thinking_token_buffer) || 0,
      current_model: form.model_mode === 'fixed' ? (form.current_model || null) : null,
      temperature: customTemp.value && form.temperature !== null && form.temperature !== undefined
        ? Number(form.temperature)
        : null,
      model_mode: form.model_mode
    })
    availableModels.value = data.available_models || []
    effectiveModel.value = data.effective_model || ''
    modelMode.value = data.model_mode || 'default'
    fillForm(data)
    ElMessage.success('模型配置已保存并生效')
    emit('saved', data)
  } catch (e) {
    ElMessage.error('保存失败：' + (e.message || e))
  } finally {
    saving.value = false
  }
}
</script>

<style scoped>
.field-label {
  display: inline-flex;
  align-items: center;
  gap: 4px;
}
.tip-icon {
  color: #909399;
  cursor: help;
  font-size: 14px;
  transition: color 0.2s;
}
.tip-icon:hover {
  color: #409eff;
}
.tip {
  max-width: 260px;
  line-height: 1.6;
  font-size: 12px;
}
.temp-row {
  display: flex;
  align-items: center;
}
</style>
