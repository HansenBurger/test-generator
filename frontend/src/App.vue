<template>
  <el-container class="app-container">
    <el-header class="app-header">
      <div class="header-content">
        <div class="header-left">
          <h1 class="site-title" @click="$router.push('/')">测试大纲与用例生成器</h1>
        </div>
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
      </div>
    </el-header>
    <el-main>
      <router-view />
    </el-main>
  </el-container>
</template>

<script setup>
import { computed } from 'vue'
import { useRoute } from 'vue-router'

const route = useRoute()
const showNav = computed(() => route.path !== '/')
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
