import { reactive, computed } from 'vue'

// 生成锁：预生成/批量生成调用大模型期间，禁止修改模型配置（顶栏配置入口禁用）。
// 按持有者计数，避免预生成与批量生成并发时互相解锁。
const lockOwners = reactive(new Map()) // owner -> reason

export const modelConfigLocked = computed(() => lockOwners.size > 0)
export const lockReason = computed(() => lockOwners.values().next().value || '')

export const acquireModelConfigLock = (owner, reason = '') => {
  lockOwners.set(owner, reason)
}

export const releaseModelConfigLock = (owner) => {
  lockOwners.delete(owner)
}

export const releaseAllModelConfigLocks = () => {
  lockOwners.clear()
}
