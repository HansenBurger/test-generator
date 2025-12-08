"""
API路由定义
"""
import io
import os
import re
import tempfile
from fastapi import APIRouter, UploadFile, File, HTTPException
from fastapi.responses import FileResponse, StreamingResponse
from typing import Optional

from app.models.schemas import ParseResponse, ParsedDocument, GenerateOutlineRequest
from app.services.doc_parser import DocumentParser
from app.services.xmind_generator import XMindGenerator
from datetime import datetime

router = APIRouter()


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
    上传并解析Word文档
    """
    # 验证文件类型
    if not file.filename.endswith(('.doc', '.docx')):
        return ParseResponse(
            success=False,
            message="不支持的文件类型，请上传Word文档（.doc或.docx）"
        )
    
    # 保存临时文件
    with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp_file:
        content = await file.read()
        tmp_file.write(content)
        tmp_path = tmp_file.name
    
    try:
        # 解析文档
        parser = DocumentParser(tmp_path)
        parsed_doc = parser.parse()
        
        return ParseResponse(
            success=True,
            message="文档解析成功",
            data=parsed_doc
        )
    except ValueError as e:
        # ValueError 通常是业务逻辑错误，直接返回
        error_msg = sanitize_error_message(str(e), file.filename)
        return ParseResponse(
            success=False,
            message=error_msg
        )
    except Exception as e:
        # 其他异常，清理错误信息
        error_msg = sanitize_error_message(str(e), file.filename)
        return ParseResponse(
            success=False,
            message=error_msg
        )
    finally:
        # 清理临时文件
        if os.path.exists(tmp_path):
            os.unlink(tmp_path)


@router.post("/generate-outline")
async def generate_outline(request: GenerateOutlineRequest):
    """
    生成XMind测试大纲
    """
    try:
        # 生成XMind文件
        generator = XMindGenerator(request.parsed_data)
        xmind_bytes = generator.generate()
        
        # 生成文件名：用例名称-版本号
        case_name = request.parsed_data.requirement_info.case_name or "测试大纲"
        version = request.parsed_data.version or ""
        if version:
            filename = f"{case_name}-{version}.xmind"
        else:
            filename = f"{case_name}.xmind"
        
        # 返回文件流
        return StreamingResponse(
            io.BytesIO(xmind_bytes),
            media_type="application/xmind",
            headers={
                "Content-Disposition": f"attachment; filename={filename}"
            }
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"生成大纲失败：{str(e)}")


@router.post("/generate-outline-from-json")
async def generate_outline_from_json(parsed_data: ParsedDocument):
    """
    从JSON数据生成XMind测试大纲（便捷接口）
    """
    try:
        # 生成XMind文件
        generator = XMindGenerator(parsed_data)
        xmind_bytes = generator.generate()
        
        # 生成文件名
        if parsed_data.document_type == "non_modeling":
            # 非建模需求：需求名称-时间戳
            requirement_name = parsed_data.requirement_name or "测试大纲"
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"{requirement_name}-{timestamp}.xmind"
        else:
            # 建模需求：用例名称-版本号
            case_name = parsed_data.requirement_info.case_name or "测试大纲"
            version = parsed_data.version or ""
            if version:
                filename = f"{case_name}-{version}.xmind"
            else:
                filename = f"{case_name}.xmind"
        
        # 对文件名进行URL编码，确保中文正确显示
        import urllib.parse
        encoded_filename = urllib.parse.quote(filename.encode('utf-8'))
        
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
        error_detail = f"生成大纲失败：{str(e)}\n{traceback.format_exc()}"
        raise HTTPException(status_code=500, detail=error_detail)

