"""
API路由定义
"""
import io
import os
import tempfile
from fastapi import APIRouter, UploadFile, File, HTTPException
from fastapi.responses import FileResponse, StreamingResponse
from typing import Optional

from app.models.schemas import ParseResponse, ParsedDocument, GenerateOutlineRequest
from app.services.doc_parser import DocumentParser
from app.services.xmind_generator import XMindGenerator
from datetime import datetime

router = APIRouter()


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
        return ParseResponse(
            success=False,
            message=str(e)
        )
    except Exception as e:
        return ParseResponse(
            success=False,
            message=f"解析失败：{str(e)}"
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
        
        # 生成文件名：用例名称-版本号（处理中文编码）
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

