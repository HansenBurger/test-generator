# test-generator — 测试大纲生成器

将 Word 格式需求文档转换为 XMind 格式测试大纲。

## 技术栈

- 后端：Python 3 + FastAPI + python-docx
- 前端：Vue 3 + Element Plus + Vite
- XMind 生成：直接生成 XMind XML 格式（非第三方库）
- 部署：docker-compose（backend + frontend）

## 快速启动

```bash
# 后端
cd backend && python main.py

# 前端
cd frontend && npm run dev
```

或使用 `scripts/dev.sh` 一键启动。

## 项目特有约定

- XMind 输出为原生 XML 格式，不依赖 xmind SDK
- 文档解析核心在 `backend/app/services/xmind_parser.py`
- 前端核心视图：`frontend/src/views/OutlineToCases.vue`

## 部署

```bash
./scripts/deploy.sh deploy
```

## 注意事项

- Word 文档格式多样，解析逻辑需兼容非标准排版
- 生成 XMind 后建议用 XMind 软件验证打开是否正常
