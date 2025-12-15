"""
FastAPI 主应用入口
"""
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from app.api import router
from app.utils.middleware import LoggingMiddleware
from app.utils.logger import logger

app = FastAPI(
    title="测试大纲生成器",
    description="将Word格式需求文档转换为XMind格式的测试大纲",
    version="1.0.0"
)

# 配置日志中间件（需要在CORS之前添加）
app.add_middleware(LoggingMiddleware)

# 配置CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # 生产环境应限制具体域名
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# 注册路由
app.include_router(router, prefix="/api")

# 记录应用启动
logger.info("FastAPI应用启动")

# 启动LibreOffice守护进程（用于加速.doc文件转换）
try:
    from app.utils.libreoffice_daemon import libreoffice_daemon
    logger.info("LibreOffice守护进程管理器已初始化")
except Exception as e:
    logger.warning(f"LibreOffice守护进程初始化失败: {str(e)}，将使用普通模式")

@app.get("/")
async def root():
    return {"message": "测试大纲生成器API服务运行中"}

if __name__ == "__main__":
    import uvicorn
    import sys
    # 默认端口8001，可以通过命令行参数指定：python main.py --port 8001
    port = 8001
    if len(sys.argv) > 1 and "--port" in sys.argv:
        try:
            port_index = sys.argv.index("--port")
            port = int(sys.argv[port_index + 1])
        except (ValueError, IndexError):
            pass
    uvicorn.run(app, host="0.0.0.0", port=port)

