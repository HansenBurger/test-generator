import axios from 'axios'

const api = axios.create({
    baseURL: '/api',
    timeout: 300000, // 300秒（5分钟）超时，因为大文件转换可能需要较长时间
    headers: {
        'Content-Type': 'application/json'
    }
})

// 请求拦截器
api.interceptors.request.use(
    config => {
        return config
    },
    error => {
        return Promise.reject(error)
    }
)

// 响应拦截器
api.interceptors.response.use(
    response => {
        return response.data
    },
    error => {
        // 处理不同类型的错误
        let message = '请求失败'

        if (error.response) {
            // 服务器返回了错误响应
            const status = error.response.status
            const data = error.response.data

            if (data?.detail) {
                message = data.detail
            } else if (data?.message) {
                message = data.message
            } else if (status === 404) {
                message = '请求的资源不存在，请检查API路径是否正确'
            } else if (status === 500) {
                message = data?.detail || '服务器内部错误'
            } else {
                message = `请求失败 (状态码: ${status})`
            }
        } else if (error.request) {
            // 请求已发出但没有收到响应
            message = '无法连接到服务器，请检查网络连接或服务器是否运行'
        } else {
            // 其他错误
            message = error.message || '请求失败'
        }

        return Promise.reject(new Error(message))
    }
)

/**
 * 上传并解析Word文档（已废弃，请使用 parseDocumentAsync）
 * @deprecated 请使用 parseDocumentAsync 和 pollTaskUntilComplete
 * @param {File} file - Word文档文件
 * @returns {Promise} 解析结果
 */
export const parseDocument = (file) => {
    // 为了向后兼容保留，但内部使用异步方式
    console.warn('parseDocument 已废弃，请使用 parseDocumentAsync')
    return parseDocumentAsync(file).then(taskResponse => {
        if (!taskResponse.success) {
            throw new Error(taskResponse.message)
        }
        // 轮询直到完成（最多60秒）
        return pollTaskUntilComplete(taskResponse.task_id, null, 1000, 60000)
    })
}

/**
 * 生成XMind测试大纲
 * @param {Object} parsedData - 解析后的文档数据
 * @returns {Promise} 文件下载
 */
export const generateOutline = (parsedData) => {
    return api.post('/generate-outline', {
        parsed_data: parsedData
    }, {
        responseType: 'blob'
    })
}

/**
 * 从JSON数据生成XMind测试大纲
 * @param {Object} parsedData - 解析后的文档数据
 * @returns {Promise} 文件下载
 */
export const generateOutlineFromJson = (parsedData) => {
    return api.post('/generate-outline-from-json', parsedData, {
        responseType: 'blob'
    })
}

/**
 * 异步上传并解析Word文档
 * @param {File} file - Word文档文件
 * @returns {Promise} 任务创建响应，包含task_id
 */
export const parseDocumentAsync = (file) => {
    const formData = new FormData()
    formData.append('file', file)

    return api.post('/parse-doc-async', formData, {
        headers: {
            'Content-Type': 'multipart/form-data'
        }
    })
}

/**
 * 查询任务状态
 * @param {string} taskId - 任务ID
 * @returns {Promise} 任务状态
 */
export const getTaskStatus = (taskId) => {
    return api.get(`/task/${taskId}`)
}

/**
 * 轮询任务状态直到完成
 * @param {string} taskId - 任务ID
 * @param {Function} onProgress - 进度回调函数 (progress) => {}
 * @param {number} interval - 轮询间隔（毫秒），默认1000ms
 * @param {number} timeout - 超时时间（毫秒），默认300000ms（5分钟）
 * @returns {Promise} 任务结果
 */
export const pollTaskUntilComplete = async (taskId, onProgress = null, interval = 1000, timeout = 300000) => {
    const startTime = Date.now()

    while (true) {
        // 检查超时
        if (Date.now() - startTime > timeout) {
            throw new Error('任务轮询超时')
        }

        try {
            const status = await getTaskStatus(taskId)

            // 调用进度回调
            if (onProgress) {
                onProgress(status.progress, status.status)
            }

            // 检查任务状态
            if (status.status === 'completed') {
                if (status.result && status.result.success) {
                    return status.result
                } else {
                    throw new Error(status.result?.message || status.error || '任务完成但结果失败')
                }
            } else if (status.status === 'failed') {
                throw new Error(status.error || '任务处理失败')
            }

            // 等待后继续轮询
            await new Promise(resolve => setTimeout(resolve, interval))

        } catch (error) {
            // 如果是404，任务不存在
            if (error.response && error.response.status === 404) {
                throw new Error('任务不存在')
            }
            throw error
        }
    }
}

/**
 * 上传并解析XMind测试大纲
 * @param {File} file - xmind文件
 * @returns {Promise}
 */
export const parseXmind = (file) => {
    const formData = new FormData()
    formData.append('file', file)
    return api.post('/parse-xmind', formData, {
        headers: {
            'Content-Type': 'multipart/form-data'
        }
    })
}

/**
 * 预生成测试用例
 * @param {string} parseId
 * @param {number} count
 */
export const previewGenerate = (parseId, count = 4) => {
    return api.post('/preview-generate', {
        parse_id: parseId,
        count
    })
}

/**
 * 确认预生成
 * @param {string} previewId
 * @param {string} strategy
 */
export const confirmPreview = (previewId, strategy = 'standard') => {
    return api.post('/confirm-preview', {
        preview_id: previewId,
        strategy
    })
}

/**
 * 批量生成
 * @param {string} parseId
 * @param {string} strategy
 */
export const bulkGenerate = (parseId, strategy = 'standard') => {
    return api.post('/bulk-generate', {
        parse_id: parseId,
        strategy
    })
}

/**
 * 查询生成任务状态
 * @param {string} taskId
 */
export const getGenerationStatus = (taskId) => {
    return api.get(`/generation-status?task_id=${taskId}`)
}

/**
 * 重新生成
 * @param {string} taskId
 * @param {string} strategy
 */
export const retryGeneration = (taskId, strategy = 'standard') => {
    return api.post('/retry-generation', {
        task_id: taskId,
        strategy
    })
}

/**
 * 导出测试用例
 * @param {string} requirementName
 * @param {Array} cases
 */
export const exportCases = (requirementName, cases) => {
    return api.post('/export-cases', {
        requirement_name: requirementName,
        cases
    }, {
        responseType: 'blob'
    })
}

/**
 * 根据session_id导出测试用例
 * @param {string} sessionId
 */
export const exportCasesBySession = (sessionId) => {
    return api.get(`/export-cases-by-session?session_id=${sessionId}`, {
        responseType: 'blob'
    })
}

/**
 * 根据session_id导出测试用例（包含响应头）
 * @param {string} sessionId
 */
export const exportCasesBySessionWithHeaders = (sessionId) => {
    return axios.get(`/api/export-cases-by-session?session_id=${sessionId}`, {
        responseType: 'blob'
    })
}

export default api



