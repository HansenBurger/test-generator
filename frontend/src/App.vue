<template>
  <el-container class="app-container">
    <el-header class="app-header">
      <h1>测试大纲与用例生成器</h1>
    </el-header>
    <el-main>
      <router-view />
    </el-main>
  </el-container>
</template>

<script setup>
// 主应用组件
</script>

<style scoped>
.app-container {
  min-height: 100vh;
}

.app-header {
  background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
  color: white;
  display: flex;
  flex-direction: column;
  justify-content: center;
  align-items: center;
  padding: 20px;
  box-shadow: 0 2px 8px rgba(0, 0, 0, 0.1);
}

.app-header h1 {
  margin: 0;
  font-size: 28px;
  font-weight: 600;
}

.subtitle {
  margin: 8px 0 0 0;
  font-size: 14px;
  opacity: 0.9;
}

.el-main {
  padding: 40px;
  max-width: 1200px;
  margin: 0 auto;
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
}
</style>

