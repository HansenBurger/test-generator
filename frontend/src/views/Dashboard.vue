<template>
  <div class="dashboard-container">
    <el-row :gutter="20">
      <el-col :span="12">
        <el-card class="entry-card" shadow="hover" @click="goTo('/outline-to-cases')">
          <div class="entry-title">大纲转用例</div>
          <div class="entry-desc">上传XMind大纲，生成测试用例并导出</div>
        </el-card>
      </el-col>
      <el-col :span="12">
        <el-card class="entry-card" shadow="hover" @click="goTo('/outline-generation')">
          <div class="entry-title">需求生成大纲</div>
          <div class="entry-desc">上传Word文档，生成测试大纲XMind</div>
        </el-card>
      </el-col>
    </el-row>
  </div>
</template>

<script setup>
import { useRouter } from 'vue-router'

const router = useRouter()
const goTo = (path) => {
  router.push(path)
}
</script>

<style scoped>
.dashboard-container {
  padding: 10px 0;
}

.entry-card {
  cursor: pointer;
  transition: transform 0.2s ease;
  min-height: 140px;
}

.entry-card:hover {
  transform: translateY(-4px);
}

.entry-title {
  font-size: 18px;
  font-weight: 600;
  margin-bottom: 10px;
}

.entry-desc {
  color: #606266;
  font-size: 14px;
}
</style>
