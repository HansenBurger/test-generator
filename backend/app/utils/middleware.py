"""
FastAPI中间件 - 记录API请求和响应日志
"""
import time
from fastapi import Request
from starlette.middleware.base import BaseHTTPMiddleware
from app.utils.logger import api_logger


class LoggingMiddleware(BaseHTTPMiddleware):
    """记录API请求和响应的中间件"""
    
    async def dispatch(self, request: Request, call_next):
        """处理请求并记录日志"""
        start_time = time.time()
        
        # 记录请求信息
        client_ip = request.client.host if request.client else "unknown"
        method = request.method
        url = str(request.url)
        path = request.url.path
        query_params = dict(request.query_params)
        
        # 记录请求开始（包括所有路径，即使是404）
        api_logger.info(
            f"请求开始 - IP: {client_ip}, 方法: {method}, 路径: {path}, "
            f"查询参数: {query_params}, URL: {url}"
        )
        
        # 处理请求
        try:
            response = await call_next(request)
            
            # 计算处理时间
            process_time = time.time() - start_time
            
            # 记录响应信息
            status_code = response.status_code
            
            # 对于404错误，记录更详细的信息
            if status_code == 404:
                api_logger.warning(
                    f"请求404 - IP: {client_ip}, 方法: {method}, 路径: {path}, "
                    f"查询参数: {query_params}, URL: {url}, 处理时间: {process_time:.3f}秒"
                )
            else:
                api_logger.info(
                    f"请求完成 - IP: {client_ip}, 方法: {method}, 路径: {path}, "
                    f"状态码: {status_code}, 处理时间: {process_time:.3f}秒"
                )
            
            return response
            
        except Exception as e:
            # 计算处理时间
            process_time = time.time() - start_time
            
            # 记录异常
            api_logger.error(
                f"请求异常 - IP: {client_ip}, 方法: {method}, 路径: {path}, "
                f"查询参数: {query_params}, URL: {url}, "
                f"错误: {str(e)}, 处理时间: {process_time:.3f}秒",
                exc_info=True
            )
            raise


