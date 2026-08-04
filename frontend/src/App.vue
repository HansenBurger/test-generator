<template>
  <el-container class="app-container">
    <el-header class="app-header">
      <div class="header-content">
        <div class="header-left">
          <h1 class="site-title" @click="$router.push('/')">测试大纲与用例生成器</h1>
        </div>
        <div class="header-actions">
          <nav v-if="showNav" class="header-nav">
            <el-button
              text
              :class="{ 'nav-active': $route.path === '/' }"
              @click="$router.push('/')"
            >
              首页
            </el-button>
            <el-button
              text
              :class="{ 'nav-active': $route.path === '/outline-to-cases' }"
              @click="$router.push('/outline-to-cases')"
            >
              大纲转用例
            </el-button>
            <el-button
              text
              :class="{ 'nav-active': $route.path === '/outline-generation' }"
              @click="$router.push('/outline-generation')"
            >
              需求生成大纲
            </el-button>
          </nav>
          <el-tag v-if="currentModel" class="model-tag" effect="dark" round>
            <el-icon class="model-tag-icon"><Cpu /></el-icon>{{ currentModel === 'auto' ? 'auto · 轮转' : currentModel }}
          </el-tag>
          <el-button
            class="config-btn"
            :icon="Setting"
            circle
            :disabled="modelConfigLocked"
            :title="modelConfigLocked ? lockReason : '模型配置'"
            @click="openConfig"
          />
        </div>
      </div>
    </el-header>
    <el-main>
      <router-view />
    </el-main>
    <ModelConfigDialog v-model="configVisible" @saved="loadCurrentModel" />
  </el-container>
</template>

<script setup>
import { computed, ref, onMounted } from 'vue'
import { useRoute } from 'vue-router'
import { Setting, Cpu } from '@element-plus/icons-vue'
import { getModelConfig } from './utils/api'
import { modelConfigLocked, lockReason } from './utils/generationLock'
import ModelConfigDialog from './components/ModelConfigDialog.vue'

const route = useRoute()
const showNav = computed(() => route.path !== '/')

const configVisible = ref(false)
const currentModel = ref('')

const openConfig = () => {
  // 预生成/批量生成调用模型期间禁止修改配置
  if (modelConfigLocked.value) {
    return
  }
  configVisible.value = true
}

const loadCurrentModel = async () => {
  try {
    const data = await getModelConfig()
    currentModel.value = data.effective_model || ''
  } catch (e) {
    // 顶部模型展示非关键，静默失败
  }
}

onMounted(loadCurrentModel)
</script>

<style scoped>
.app-container {
  min-height: 100vh;
}

.app-header {
  background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
  color: white;
  padding: 0 24px;
  box-shadow: 0 2px 8px rgba(0, 0, 0, 0.1);
  height: auto !important;
  min-height: 60px;
}

.header-content {
  display: flex;
  align-items: center;
  justify-content: space-between;
  max-width: 1400px;
  margin: 0 auto;
  padding: 12px 0;
}

.header-left {
  display: flex;
  align-items: center;
}

.site-title {
  margin: 0;
  font-size: 22px;
  font-weight: 600;
  cursor: pointer;
  user-select: none;
  white-space: nowrap;
  transition: opacity 0.2s;
}

.site-title:hover {
  opacity: 0.9;
}

.header-actions {
  display: flex;
  align-items: center;
  gap: 12px;
}

.header-nav {
  display: flex;
  align-items: center;
  gap: 4px;
}

.header-nav .el-button {
  color: rgba(255, 255, 255, 0.85);
  font-size: 14px;
  padding: 8px 16px;
  border-radius: 6px;
  transition: all 0.2s;
}

.header-nav .el-button:hover {
  color: white;
  background: rgba(255, 255, 255, 0.15);
}

.header-nav .el-button.nav-active {
  color: white;
  background: rgba(255, 255, 255, 0.2);
  font-weight: 500;
}

.model-tag {
  display: inline-flex;
  align-items: center;
  max-width: 240px;
}

.model-tag :deep(.el-tag__content) {
  display: inline-flex;
  align-items: center;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.model-tag-icon {
  margin-right: 4px;
}

.config-btn {
  color: #fff;
  background: rgba(255, 255, 255, 0.15);
  border-color: rgba(255, 255, 255, 0.4);
}

.config-btn:hover,
.config-btn:focus {
  color: #fff;
  background: rgba(255, 255, 255, 0.28);
  border-color: rgba(255, 255, 255, 0.7);
}

.config-btn.is-disabled,
.config-btn.is-disabled:hover,
.config-btn.is-disabled:focus {
  color: rgba(255, 255, 255, 0.45);
  background: rgba(255, 255, 255, 0.08);
  border-color: rgba(255, 255, 255, 0.2);
  cursor: not-allowed;
}

.el-main {
  padding: 40px;
  max-width: 1400px;
  margin: 0 auto;
  width: 100%;
}
</style>

<style>
* {
  margin: 0;
  padding: 0;
  box-sizing: border-box;
}

body {
  font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif;
  background: #f5f7fa;
}
</style>
