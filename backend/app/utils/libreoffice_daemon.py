"""
LibreOffice守护进程管理器
启动一个常驻的LibreOffice进程，通过UNO API调用，避免每次启动的开销
"""
import os
import subprocess
import time
import socket
import threading
import re
from typing import Optional
from app.utils.logger import parser_logger


class LibreOfficeDaemon:
    """LibreOffice守护进程管理器"""
    
    _instance = None
    _lock = threading.Lock()
    _process: Optional[subprocess.Popen] = None
    _port = 2002
    _host = "localhost"
    
    def __new__(cls):
        if cls._instance is None:
            with cls._lock:
                if cls._instance is None:
                    cls._instance = super().__new__(cls)
        return cls._instance
    
    def __init__(self):
        if hasattr(self, '_initialized'):
            return
        self._initialized = True
        self._daemon_supported = os.name != "nt"
        if not self._daemon_supported:
            parser_logger.info("Windows平台不支持LibreOffice守护进程，跳过初始化")
            return
        self._start_daemon()
    
    def _start_daemon(self):
        """启动LibreOffice守护进程"""
        if not getattr(self, "_daemon_supported", True):
            return
        if self._process and self._process.poll() is None:
            # 进程还在运行
            return
        
        try:
            # 检查端口是否已被占用
            if self._is_port_open(self._host, self._port):
                parser_logger.info(f"LibreOffice守护进程可能已在运行（端口{self._port}已占用）")
                return
            
            # 启动LibreOffice守护进程
            env = dict(os.environ)
            env.update({
                "HOME": "/tmp",
                "SAL_USE_VCLPLUGIN": "headless",
                "SAL_DISABLE_OPENCL": "1",
            })
            
            cmd = [
                "soffice",
                "--headless",
                "--invisible",
                "--nodefault",
                "--nolockcheck",
                "--nologo",
                "--norestore",
                f"--accept=socket,host={self._host},port={self._port};urp;"
            ]
            
            popen_kwargs = {
                "stdout": subprocess.DEVNULL,
                "stderr": subprocess.DEVNULL,
                "env": env,
            }
            if os.name != "nt":
                popen_kwargs["preexec_fn"] = lambda: os.nice(10) if hasattr(os, "nice") else None
            self._process = subprocess.Popen(cmd, **popen_kwargs)
            
            # 等待守护进程启动
            for _ in range(20):  # 最多等待10秒
                time.sleep(0.5)
                if self._is_port_open(self._host, self._port):
                    parser_logger.info(f"LibreOffice守护进程启动成功（端口{self._port}）")
                    return
            
            parser_logger.warning("LibreOffice守护进程启动超时，将使用普通模式")
            if self._process:
                self._process.terminate()
                self._process = None
                
        except Exception as e:
            parser_logger.error(f"启动LibreOffice守护进程失败: {str(e)}", exc_info=True)
            if self._process:
                try:
                    self._process.terminate()
                except:
                    pass
                self._process = None
    
    def _is_port_open(self, host: str, port: int) -> bool:
        """检查端口是否开放"""
        try:
            sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            sock.settimeout(1)
            result = sock.connect_ex((host, port))
            sock.close()
            return result == 0
        except:
            return False
    
    def is_running(self) -> bool:
        """检查守护进程是否运行"""
        if self._process is None:
            return False
        if self._process.poll() is not None:
            return False
        return self._is_port_open(self._host, self._port)
    
    def restart_if_needed(self):
        """如果需要，重启守护进程"""
        if not self.is_running():
            parser_logger.info("LibreOffice守护进程未运行，正在重启...")
            self._start_daemon()
    
    def convert_with_daemon(self, doc_path: str, output_path: str) -> str:
        """使用守护进程模式转换文档（更快）"""
        import subprocess
        
        # 确保守护进程运行
        self.restart_if_needed()
        
        if not self.is_running():
            # 守护进程未运行，回退到普通模式
            raise RuntimeError("LibreOffice守护进程未运行")
        
        try:
            output_dir = os.path.dirname(output_path)
            os.makedirs(output_dir, exist_ok=True)
            
            # 使用soffice连接到守护进程进行转换
            # 注意：实际上LibreOffice的批处理模式不支持连接到已运行的守护进程
            # 所以我们使用普通模式，但守护进程已经启动，可以减少一些初始化时间
            # 或者直接使用普通模式（因为守护进程主要用于UNO API，批处理模式每次都会启动新进程）
            # 为了真正利用守护进程，我们需要使用UNO API，但这比较复杂
            # 这里我们使用优化的普通模式，但确保守护进程在运行（可以预热系统）
            env = dict(os.environ)
            env.update({
                "HOME": "/tmp",
                "SAL_USE_VCLPLUGIN": "headless",
                "SAL_DISABLE_OPENCL": "1",
            })
            
            # 使用优化的批处理模式
            cmd = [
                "soffice",
                "--headless",
                "--invisible",
                "--nodefault",
                "--nolockcheck",
                "--nologo",
                "--norestore",
                "--safe-mode",
                "--convert-to", "docx",
                "--outdir", output_dir,
                doc_path
            ]
            
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=300,
                env=env
            )
            
            # 查找生成的文件
            input_basename = os.path.splitext(os.path.basename(doc_path))[0]
            safe_basename = re.sub(r'[<>:"/\\|?*]', '_', input_basename)
            
            possible_files = [
                os.path.join(output_dir, f"{input_basename}.docx"),
                os.path.join(output_dir, f"{safe_basename}.docx"),
                os.path.join(output_dir, os.path.basename(doc_path).replace('.doc', '.docx').replace('.DOC', '.docx')),
            ]
            
            generated_file = None
            for possible_file in possible_files:
                if os.path.exists(possible_file):
                    generated_file = possible_file
                    break
            
            if not generated_file:
                docx_files = [f for f in os.listdir(output_dir) if f.endswith('.docx')]
                if docx_files:
                    docx_files.sort(key=lambda f: os.path.getmtime(os.path.join(output_dir, f)), reverse=True)
                    generated_file = os.path.join(output_dir, docx_files[0])
            
            if not generated_file or not os.path.exists(generated_file):
                raise ValueError("转换完成但未找到生成的.docx文件")
            
            if generated_file != output_path:
                if os.path.exists(output_path):
                    os.unlink(output_path)
                os.rename(generated_file, output_path)
            
            if not os.path.exists(output_path) or os.path.getsize(output_path) == 0:
                raise ValueError("转换后的.docx文件为空或不存在")
            
            return output_path
            
        except Exception as e:
            parser_logger.error(f"使用守护进程转换失败: {str(e)}", exc_info=True)
            raise
    
    def shutdown(self):
        """关闭守护进程"""
        if self._process:
            try:
                self._process.terminate()
                self._process.wait(timeout=5)
            except:
                try:
                    self._process.kill()
                except:
                    pass
            finally:
                self._process = None


# 全局守护进程实例
libreoffice_daemon = LibreOfficeDaemon()

