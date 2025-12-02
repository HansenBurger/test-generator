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
        const message = error.response?.data?.detail || error.message || '请求失败'
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



