"""
API路由定义
"""
import io
import os
import re
import tempfile
import asyncio
from fastapi import APIRouter, UploadFile, File, HTTPException
from fastapi.responses import FileResponse, StreamingResponse
from typing import Optional

from app.models.schemas import (
    ParseResponse, ParsedDocument, GenerateOutlineRequest,
    TaskCreateResponse, TaskStatusResponse
)
from app.services.doc_parser import DocumentParser
from app.services.xmind_generator import XMindGenerator
from app.utils.logger import api_logger, parser_logger, generator_logger
from app.utils.task_manager import task_manager, TaskStatus
from datetime import datetime

router = APIRouter()

@router.get("/health")
async def health_check():
    """健康检查端点"""
    return {"status": "ok", "message": "API服务正常运行"}


def sanitize_error_message(error_msg: str, filename: str) -> str:
    """
    清理错误信息，移除临时路径等内部信息
    只保留用户友好的错误描述和文件名
    返回格式：文件名 解析失败：错误原因
    """
    # 移除临时路径（Windows和Unix路径）
    # 匹配类似 "Package not found at 'C:\Users\...\tmpyug_3_c4.docx'" 的模式
    error_msg = re.sub(r"Package not found at ['\"][^'\"]+['\"]", "文件格式错误或文件已损坏", error_msg)
    error_msg = re.sub(r"at ['\"][^'\"]*[Tt]emp[^'\"]*['\"]", "", error_msg)
    error_msg = re.sub(r"['\"][A-Za-z]:\\[^'\"]*['\"]", "", error_msg)  # Windows绝对路径
    error_msg = re.sub(r"['\"]/[^'\"]*['\"]", "", error_msg)  # Unix绝对路径
    
    # 移除常见的临时文件路径模式
    error_msg = re.sub(r"tmp[a-z0-9_]+\.(docx?|zip)", "", error_msg, flags=re.IGNORECASE)
    
    # 移除文件名（如果错误信息中已经包含，避免重复）
    if filename:
        error_msg = error_msg.replace(filename, "").strip()
    
    # 清理多余的空格和标点
    error_msg = re.sub(r"\s+", " ", error_msg).strip()
    error_msg = re.sub(r"^\s*[:：]\s*", "", error_msg)  # 移除开头的冒号
    error_msg = re.sub(r"^\s*解析失败\s*[:：]\s*", "", error_msg)  # 移除开头的"解析失败："
    
    # 如果错误信息为空或只包含技术细节，提供通用错误信息
    if not error_msg or len(error_msg) < 3:
        error_msg = "文档格式错误或文件已损坏，请检查文件是否正确"
    
    # 统一格式：文件名 解析失败：错误原因
    if filename:
        return f"{filename} 解析失败：{error_msg}"
    else:
        return f"解析失败：{error_msg}"


@router.post("/parse-doc", response_model=ParseResponse)
async def parse_document(file: UploadFile = File(...)):
    """
    上传并解析Word文档（兼容接口，内部使用异步方式）
    建议使用 /api/parse-doc-async 接口
    """
    # 为了向后兼容，保留此接口，但内部使用异步方式
    api_logger.info(f"收到同步文档解析请求（将转为异步） - 文件名: {file.filename}")
    
    # 验证文件类型
    if not file.filename.endswith(('.doc', '.docx')):
        api_logger.warning(f"不支持的文件类型 - 文件名: {file.filename}")
        return ParseResponse(
            success=False,
            message="不支持的文件类型，请上传Word文档（.doc或.docx）"
        )
    
    # 保存临时文件
    file_ext = '.doc' if file.filename.endswith('.doc') else '.docx'
    tmp_path = None
    tmp_file = None
    
    try:
        # 读取文件内容
        content = await file.read()
        if not content:
            api_logger.warning(f"上传的文件为空 - 文件名: {file.filename}")
            return ParseResponse(
                success=False,
                message="上传的文件为空，请检查文件是否正确"
            )
        
        # 创建任务
        task_id = task_manager.create_task(file.filename)
        
        # 创建临时文件并写入
        tmp_file = tempfile.NamedTemporaryFile(delete=False, suffix=file_ext)
        tmp_path = tmp_file.name
        try:
            tmp_file.write(content)
            tmp_file.flush()
            os.fsync(tmp_file.fileno())
        finally:
            tmp_file.close()
        
        # 验证文件
        if not os.path.exists(tmp_path) or os.path.getsize(tmp_path) == 0:
            raise ValueError("临时文件创建失败")
        
        api_logger.info(f"临时文件保存成功 - 任务ID: {task_id}, 文件名: {file.filename}, 路径: {tmp_path}")
        
        # 提交异步任务
        task_manager.submit_task(task_id, _process_document_async, tmp_path, file.filename)
        
        # 等待任务完成（最多等待50秒，留10秒缓冲给前端）
        import time
        start_time = time.time()
        timeout = 50  # 50秒，留10秒缓冲
        
        while True:
            task = task_manager.get_task(task_id)
            if not task:
                raise ValueError("任务不存在")
            
            if task.status == TaskStatus.COMPLETED:
                if task.result:
                    return task.result
                else:
                    raise ValueError(task.error or "任务完成但结果为空")
            
            if task.status == TaskStatus.FAILED:
                raise ValueError(task.error or "任务处理失败")
            
            # 检查超时（50秒内未完成，返回任务ID让前端轮询）
            elapsed = time.time() - start_time
            if elapsed > timeout:
                # 返回特殊响应，包含任务ID
                api_logger.info(f"同步接口超时，返回任务ID - 任务ID: {task_id}, 已等待: {elapsed:.2f}秒")
                return ParseResponse(
                    success=False,
                    message=f"处理时间较长，请使用任务ID查询结果。任务ID: {task_id}",
                    data=None
                )
            
            # 等待500ms后继续检查
            await asyncio.sleep(0.5)
    
    except HTTPException:
        raise
    except Exception as e:
        api_logger.error(f"处理文档解析请求失败 - 文件名: {file.filename}, 错误: {str(e)}", exc_info=True)
        # 清理临时文件
        if tmp_path and os.path.exists(tmp_path):
            try:
                os.unlink(tmp_path)
            except:
                pass
        return ParseResponse(
            success=False,
            message=f"处理失败: {str(e)}"
        )


def _process_document_async(tmp_path: str, filename: str) -> ParseResponse:
    """异步处理文档的函数"""
    import time
    start_time = time.time()
    
    try:
        # 解析文档
        parser = DocumentParser(tmp_path)
        parsed_doc = parser.parse()
        
        elapsed_time = time.time() - start_time
        api_logger.info(f"文档解析成功 - 文件名: {filename}, 耗时: {elapsed_time:.3f}秒")
        
        return ParseResponse(
            success=True,
            message="文档解析成功",
            data=parsed_doc
        )
    except ValueError as e:
        # ValueError 通常是业务逻辑错误
        elapsed_time = time.time() - start_time
        error_msg = sanitize_error_message(str(e), filename)
        api_logger.warning(f"文档解析失败（业务错误） - 文件名: {filename}, 错误: {error_msg}, 耗时: {elapsed_time:.3f}秒")
        return ParseResponse(
            success=False,
            message=error_msg
        )
    except Exception as e:
        # 其他异常
        elapsed_time = time.time() - start_time
        error_msg = sanitize_error_message(str(e), filename)
        api_logger.error(f"文档解析失败（系统错误） - 文件名: {filename}, 错误: {error_msg}, 耗时: {elapsed_time:.3f}秒", exc_info=True)
        return ParseResponse(
            success=False,
            message=error_msg
        )
    finally:
        # 清理临时文件
        if os.path.exists(tmp_path):
            try:
                os.unlink(tmp_path)
            except:
                pass


@router.post("/parse-doc-async", response_model=TaskCreateResponse)
async def parse_document_async(file: UploadFile = File(...)):
    """
    异步上传并解析Word文档
    立即返回任务ID，通过 /api/task/{task_id} 查询状态
    """
    api_logger.info(f"收到异步文档解析请求 - 文件名: {file.filename}, 文件大小: {file.size if hasattr(file, 'size') else 'unknown'}")
    
    # 验证文件类型
    if not file.filename.endswith(('.doc', '.docx')):
        api_logger.warning(f"不支持的文件类型 - 文件名: {file.filename}")
        return TaskCreateResponse(
            success=False,
            task_id="",
            message="不支持的文件类型，请上传Word文档（.doc或.docx）"
        )
    
    # 保存临时文件
    file_ext = '.doc' if file.filename.endswith('.doc') else '.docx'
    tmp_path = None
    tmp_file = None
    
    try:
        # 读取文件内容
        content = await file.read()
        if not content:
            api_logger.warning(f"上传的文件为空 - 文件名: {file.filename}")
            return TaskCreateResponse(
                success=False,
                task_id="",
                message="上传的文件为空，请检查文件是否正确"
            )
        
        # 创建任务
        task_id = task_manager.create_task(file.filename)
        
        # 创建临时文件并写入
        tmp_file = tempfile.NamedTemporaryFile(delete=False, suffix=file_ext)
        tmp_path = tmp_file.name
        try:
            tmp_file.write(content)
            tmp_file.flush()
            os.fsync(tmp_file.fileno())
        finally:
            tmp_file.close()
        
        # 验证文件
        if not os.path.exists(tmp_path) or os.path.getsize(tmp_path) == 0:
            raise ValueError("临时文件创建失败")
        
        api_logger.info(f"临时文件保存成功 - 任务ID: {task_id}, 文件名: {file.filename}, 路径: {tmp_path}")
        
        # 提交异步任务
        task_manager.submit_task(task_id, _process_document_async, tmp_path, file.filename)
        
        return TaskCreateResponse(
            success=True,
            task_id=task_id,
            message="任务已提交，请使用任务ID查询状态"
        )
        
    except Exception as e:
        api_logger.error(f"创建异步任务失败 - 文件名: {file.filename}, 错误: {str(e)}", exc_info=True)
        # 清理临时文件
        if tmp_path and os.path.exists(tmp_path):
            try:
                os.unlink(tmp_path)
            except:
                pass
        return TaskCreateResponse(
            success=False,
            task_id="",
            message=f"创建任务失败: {str(e)}"
        )


@router.get("/task/{task_id}", response_model=TaskStatusResponse)
async def get_task_status(task_id: str):
    """查询任务状态"""
    task = task_manager.get_task(task_id)
    
    if not task:
        raise HTTPException(status_code=404, detail="任务不存在")
    
    return TaskStatusResponse(
        task_id=task.task_id,
        status=task.status.value,
        filename=task.filename,
        progress=task.progress,
        created_at=task.created_at.isoformat(),
        started_at=task.started_at.isoformat() if task.started_at else None,
        completed_at=task.completed_at.isoformat() if task.completed_at else None,
        error=task.error,
        result=task.result
    )


@router.post("/generate-outline")
async def generate_outline(request: GenerateOutlineRequest):
    """
    生成XMind测试大纲
    """
    import time
    start_time = time.time()
    
    doc_type = request.parsed_data.document_type
    doc_name = request.parsed_data.requirement_name if doc_type == "non_modeling" else (
        request.parsed_data.requirement_info.case_name if request.parsed_data.requirement_info else "未知"
    )
    api_logger.info(f"收到生成大纲请求 - 文档类型: {doc_type}, 名称: {doc_name}")
    
    try:
        # 生成XMind文件
        generator = XMindGenerator(request.parsed_data)
        xmind_bytes = generator.generate()
        
        # 生成文件名：统一格式为需求名称-时间戳
        if request.parsed_data.document_type == "non_modeling":
            requirement_name = request.parsed_data.requirement_name or "测试大纲"
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"{requirement_name}-{timestamp}.xmind"
        else:
            case_name = (request.parsed_data.requirement_info.case_name 
                        if request.parsed_data.requirement_info 
                        else None) or "测试大纲"
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"{case_name}-{timestamp}.xmind"
        
        elapsed_time = time.time() - start_time
        api_logger.info(f"大纲生成成功 - 文件名: {filename}, 耗时: {elapsed_time:.3f}秒")
        
        # 返回文件流
        return StreamingResponse(
            io.BytesIO(xmind_bytes),
            media_type="application/xmind",
            headers={
                "Content-Disposition": f"attachment; filename={filename}"
            }
        )
    except Exception as e:
        elapsed_time = time.time() - start_time
        api_logger.error(f"大纲生成失败 - 文档名称: {doc_name}, 耗时: {elapsed_time:.3f}秒, 错误: {str(e)}", exc_info=True)
        raise HTTPException(status_code=500, detail=f"生成大纲失败：{str(e)}")


@router.post("/generate-outline-from-json")
async def generate_outline_from_json(parsed_data: ParsedDocument):
    """
    从JSON数据生成XMind测试大纲（便捷接口）
    """
    import time
    start_time = time.time()
    
    doc_type = parsed_data.document_type
    doc_name = parsed_data.requirement_name if doc_type == "non_modeling" else (
        parsed_data.requirement_info.case_name if parsed_data.requirement_info else "未知"
    )
    api_logger.info(f"收到从JSON生成大纲请求 - 文档类型: {doc_type}, 名称: {doc_name}")
    
    try:
        # 生成XMind文件
        generator = XMindGenerator(parsed_data)
        xmind_bytes = generator.generate()
        
        # 生成文件名：统一格式为需求名称-时间戳
        if parsed_data.document_type == "non_modeling":
            requirement_name = parsed_data.requirement_name or "测试大纲"
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"{requirement_name}-{timestamp}.xmind"
        else:
            case_name = (parsed_data.requirement_info.case_name 
                        if parsed_data.requirement_info 
                        else None) or "测试大纲"
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"{case_name}-{timestamp}.xmind"
        
        # 对文件名进行URL编码，确保中文正确显示
        import urllib.parse
        encoded_filename = urllib.parse.quote(filename.encode('utf-8'))
        
        elapsed_time = time.time() - start_time
        api_logger.info(f"从JSON生成大纲成功 - 文件名: {filename}, 耗时: {elapsed_time:.3f}秒")
        
        # 返回文件流
        return StreamingResponse(
            io.BytesIO(xmind_bytes),
            media_type="application/xmind",
            headers={
                "Content-Disposition": f"attachment; filename*=UTF-8''{encoded_filename}",
                "Content-Type": "application/xmind"
            }
        )
    except Exception as e:
        import traceback
        elapsed_time = time.time() - start_time
        error_detail = f"生成大纲失败：{str(e)}\n{traceback.format_exc()}"
        api_logger.error(f"从JSON生成大纲失败 - 文档名称: {doc_name}, 耗时: {elapsed_time:.3f}秒, 错误: {str(e)}", exc_info=True)
        raise HTTPException(status_code=500, detail=error_detail)

