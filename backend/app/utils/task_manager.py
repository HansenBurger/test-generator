"""
异步任务管理器
使用内存存储任务状态，支持任务提交、状态查询和结果获取
"""
import uuid
import threading
import time
from enum import Enum
from typing import Dict, Optional, Any
from dataclasses import dataclass, field
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor

from app.utils.logger import api_logger


class TaskStatus(str, Enum):
    """任务状态"""
    PENDING = "pending"  # 等待中
    PROCESSING = "processing"  # 处理中
    COMPLETED = "completed"  # 已完成
    FAILED = "failed"  # 失败


@dataclass
class Task:
    """任务对象"""
    task_id: str
    filename: str
    status: TaskStatus = TaskStatus.PENDING
    created_at: datetime = field(default_factory=datetime.now)
    started_at: Optional[datetime] = None
    completed_at: Optional[datetime] = None
    result: Optional[Any] = None
    error: Optional[str] = None
    progress: float = 0.0  # 0.0 - 1.0


class TaskManager:
    """任务管理器（单例模式）"""
    _instance = None
    _lock = threading.Lock()
    
    def __new__(cls):
        if cls._instance is None:
            with cls._lock:
                if cls._instance is None:
                    cls._instance = super().__new__(cls)
                    cls._instance._initialized = False
        return cls._instance
    
    def __init__(self):
        if hasattr(self, '_initialized') and self._initialized:
            return
        
        self._tasks: Dict[str, Task] = {}
        self._lock = threading.Lock()
        self._executor = ThreadPoolExecutor(max_workers=3, thread_name_prefix="doc_parser")
        self._initialized = True
        
        # 启动清理线程，定期清理过期任务（24小时）
        cleanup_thread = threading.Thread(target=self._cleanup_expired_tasks, daemon=True)
        cleanup_thread.start()
    
    def create_task(self, filename: str) -> str:
        """创建新任务，返回任务ID"""
        task_id = str(uuid.uuid4())
        task = Task(task_id=task_id, filename=filename)
        
        with self._lock:
            self._tasks[task_id] = task
        
        api_logger.info(f"创建任务 - 任务ID: {task_id}, 文件名: {filename}")
        return task_id
    
    def submit_task(self, task_id: str, task_func, *args, **kwargs):
        """提交任务到线程池执行"""
        def task_wrapper():
            task = self.get_task(task_id)
            if not task:
                return
            
            try:
                task.status = TaskStatus.PROCESSING
                task.started_at = datetime.now()
                api_logger.info(f"任务开始处理 - 任务ID: {task_id}")
                
                # 执行任务
                result = task_func(*args, **kwargs)
                
                task.status = TaskStatus.COMPLETED
                task.completed_at = datetime.now()
                task.result = result
                task.progress = 1.0
                api_logger.info(f"任务完成 - 任务ID: {task_id}")
                
            except Exception as e:
                task.status = TaskStatus.FAILED
                task.completed_at = datetime.now()
                task.error = str(e)
                api_logger.error(f"任务失败 - 任务ID: {task_id}, 错误: {str(e)}", exc_info=True)
        
        self._executor.submit(task_wrapper)
    
    def get_task(self, task_id: str) -> Optional[Task]:
        """获取任务"""
        with self._lock:
            return self._tasks.get(task_id)
    
    def update_task_progress(self, task_id: str, progress: float):
        """更新任务进度"""
        with self._lock:
            task = self._tasks.get(task_id)
            if task:
                task.progress = max(0.0, min(1.0, progress))
    
    def _cleanup_expired_tasks(self):
        """清理过期任务（后台线程）"""
        while True:
            try:
                time.sleep(3600)  # 每小时检查一次
                now = datetime.now()
                expired_tasks = []
                
                with self._lock:
                    for task_id, task in list(self._tasks.items()):
                        # 清理24小时前的已完成或失败任务
                        if task.completed_at or task.status == TaskStatus.FAILED:
                            age = (now - task.created_at).total_seconds()
                            if age > 86400:  # 24小时
                                expired_tasks.append(task_id)
                    
                    for task_id in expired_tasks:
                        del self._tasks[task_id]
                        api_logger.info(f"清理过期任务 - 任务ID: {task_id}")
            
            except Exception as e:
                api_logger.error(f"清理过期任务时出错: {str(e)}", exc_info=True)


# 全局任务管理器实例
task_manager = TaskManager()

