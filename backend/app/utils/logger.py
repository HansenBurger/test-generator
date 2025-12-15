"""
日志模块 - 提供统一的日志记录功能
支持FastAPI接口日志、解析日志和生成日志
日志轮转：当日log存储在app.log，次日log归档到yyyy/MM/yyyy-MM-dd.log
"""
import logging
import os
from pathlib import Path
from datetime import datetime, timedelta
from logging.handlers import TimedRotatingFileHandler
from typing import Optional


class DailyRotatingFileHandler(TimedRotatingFileHandler):
    """
    自定义日志轮转处理器
    当日log存储在app.log，次日log归档到yyyy/MM/yyyy-MM-dd.log
    """
    
    def __init__(self, log_dir: str, filename: str = "app.log"):
        """
        初始化日志处理器
        
        Args:
            log_dir: 日志目录路径
            filename: 日志文件名（默认app.log）
        """
        self.log_dir = Path(log_dir)
        self.log_dir.mkdir(parents=True, exist_ok=True)
        
        self.filename = filename
        self.base_path = self.log_dir / filename
        
        # 使用TimedRotatingFileHandler，每天午夜轮转
        super().__init__(
            filename=str(self.base_path),
            when='midnight',
            interval=1,
            backupCount=0,  # 不使用backupCount，我们自己处理归档
            encoding='utf-8',
            delay=False
        )
        
        # 检查是否需要归档昨天的日志
        self._archive_previous_log()
    
    def _archive_previous_log(self):
        """归档之前的日志文件"""
        if not self.base_path.exists():
            return
        
        # 获取文件修改时间
        file_mtime = datetime.fromtimestamp(self.base_path.stat().st_mtime)
        today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
        
        # 如果文件是今天创建的，不需要归档
        if file_mtime >= today:
            return
        
        # 获取文件创建日期（使用修改时间作为参考）
        file_date = file_mtime.date()
        
        # 创建归档路径：yyyy/MM/yyyy-MM-dd.log
        archive_dir = self.log_dir / file_date.strftime("%Y") / file_date.strftime("%m")
        archive_dir.mkdir(parents=True, exist_ok=True)
        
        archive_filename = file_date.strftime("%Y-%m-%d.log")
        archive_path = archive_dir / archive_filename
        
        # 如果归档文件已存在，追加内容
        if archive_path.exists():
            with open(self.base_path, 'r', encoding='utf-8') as src:
                content = src.read()
            with open(archive_path, 'a', encoding='utf-8') as dst:
                dst.write(content)
            self.base_path.unlink()
        else:
            # 移动文件到归档目录
            self.base_path.rename(archive_path)
    
    def doRollover(self):
        """
        执行日志轮转
        当日志需要轮转时，将当前日志归档到yyyy/MM/yyyy-MM-dd.log
        """
        if self.stream:
            self.stream.close()
            self.stream = None
        
        # 归档昨天的日志
        yesterday = (datetime.now() - timedelta(days=1)).date()
        archive_dir = self.log_dir / yesterday.strftime("%Y") / yesterday.strftime("%m")
        archive_dir.mkdir(parents=True, exist_ok=True)
        
        archive_filename = yesterday.strftime("%Y-%m-%d.log")
        archive_path = archive_dir / archive_filename
        
        # 如果当前日志文件存在，移动到归档目录
        if self.base_path.exists():
            if archive_path.exists():
                # 如果归档文件已存在，追加内容
                with open(self.base_path, 'r', encoding='utf-8') as src:
                    content = src.read()
                with open(archive_path, 'a', encoding='utf-8') as dst:
                    dst.write(content)
                self.base_path.unlink()
            else:
                # 移动文件到归档目录
                self.base_path.rename(archive_path)
        
        # 创建新的日志文件
        if not self.delay:
            self.stream = self._open()


def setup_logger(
    name: str = "app",
    log_dir: Optional[str] = None,
    level: int = logging.INFO,
    format_string: Optional[str] = None
) -> logging.Logger:
    """
    设置并返回logger实例
    
    Args:
        name: logger名称
        log_dir: 日志目录路径，如果为None则使用项目根目录下的log目录
        level: 日志级别
        format_string: 日志格式字符串
    
    Returns:
        配置好的logger实例
    """
    if log_dir is None:
        # 默认使用项目根目录下的log目录
        project_root = Path(__file__).parent.parent.parent.parent
        log_dir = project_root / "log"
    
    if format_string is None:
        format_string = '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    
    logger = logging.getLogger(name)
    logger.setLevel(level)
    
    # 避免重复添加handler
    if logger.handlers:
        return logger
    
    # 创建日志目录
    log_dir = Path(log_dir)
    log_dir.mkdir(parents=True, exist_ok=True)
    
    # 添加控制台处理器
    console_handler = logging.StreamHandler()
    console_handler.setLevel(level)
    console_formatter = logging.Formatter(format_string)
    console_handler.setFormatter(console_formatter)
    logger.addHandler(console_handler)
    
    # 添加文件处理器（带轮转）
    file_handler = DailyRotatingFileHandler(str(log_dir), "app.log")
    file_handler.setLevel(level)
    file_formatter = logging.Formatter(format_string)
    file_handler.setFormatter(file_formatter)
    logger.addHandler(file_handler)
    
    return logger


# 创建默认logger实例
logger = setup_logger("app")
api_logger = setup_logger("app.api")
parser_logger = setup_logger("app.parser")
generator_logger = setup_logger("app.generator")


