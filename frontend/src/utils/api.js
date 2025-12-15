import axios from 'axios'

const api = axios.create({
    baseURL: '/api',
    timeout: 60000, // 60秒超时，因为文档解析可能需要较长时间
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
 * 上传并解析Word文档
 * @param {File} file - Word文档文件
 * @returns {Promise} 解析结果
 */
export const parseDocument = (file) => {
    const formData = new FormData()
    formData.append('file', file)

    return api.post('/parse-doc', formData, {
        headers: {
            'Content-Type': 'multipart/form-data'
        }
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

export default api



