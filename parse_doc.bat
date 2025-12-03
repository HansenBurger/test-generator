@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion

:: 设置后端服务地址和端口
set API_URL=http://localhost:8000
set API_PORT=8000

:: 检查是否有文件被拖入
if "%~1"=="" (
    echo 请将Word文档拖拽到此脚本上，或直接运行脚本后输入文件路径
    echo.
    set /p FILE_PATH="请输入Word文档路径: "
    if "!FILE_PATH!"=="" (
        echo 未指定文件，退出
        pause
        exit /b 1
    )
) else (
    set FILE_PATH=%~1
)

:: 检查文件是否存在
if not exist "!FILE_PATH!" (
    echo 错误：文件不存在 - !FILE_PATH!
    pause
    exit /b 1
)

:: 检查文件扩展名
for %%F in ("!FILE_PATH!") do set EXT=%%~xF
if not "!EXT!"==".docx" if not "!EXT!"==".doc" (
    echo 错误：不支持的文件类型，请使用 .doc 或 .docx 文件
    pause
    exit /b 1
)

echo ========================================
echo 测试大纲生成器 - 文档解析工具
echo ========================================
echo.

:: 检查后端服务是否运行
echo [1/3] 检查后端服务状态...
powershell -Command "$response = try { Invoke-WebRequest -Uri '%API_URL%/' -Method Get -TimeoutSec 2 -UseBasicParsing } catch { $null }; if ($response -and $response.StatusCode -eq 200) { exit 0 } else { exit 1 }" >nul 2>&1
if %errorlevel% neq 0 (
    echo 后端服务未运行，正在启动...
    echo.
    
    :: 检查虚拟环境是否存在
    if not exist "backend\.venv" (
        echo 错误：未找到虚拟环境 backend\.venv
        echo 请先运行: cd backend ^&^& python -m venv .venv ^&^& .venv\Scripts\Activate.ps1 ^&^& pip install -r requirements.txt
        pause
        exit /b 1
    )
    
    :: 启动后端服务（后台运行）
    echo 正在启动后端服务...
    start /B "" powershell -NoExit -Command "cd '%CD%\backend'; .\.venv\Scripts\Activate.ps1; python main.py"
    
    :: 等待服务启动
    echo 等待服务启动...
    set /a RETRY_COUNT=0
    :WAIT_SERVER
    timeout /t 1 /nobreak >nul
    powershell -Command "$response = try { Invoke-WebRequest -Uri '%API_URL%/' -Method Get -TimeoutSec 1 -UseBasicParsing } catch { $null }; if ($response -and $response.StatusCode -eq 200) { exit 0 } else { exit 1 }" >nul 2>&1
    if %errorlevel% neq 0 (
        set /a RETRY_COUNT+=1
        if !RETRY_COUNT! geq 30 (
            echo 错误：服务启动超时，请检查后端服务
            pause
            exit /b 1
        )
        goto WAIT_SERVER
    )
    echo 后端服务已启动
) else (
    echo 后端服务运行中
)
echo.

:: 解析文档并生成XMind
for %%F in ("!FILE_PATH!") do set FILE_NAME=%%~nxF
echo [2/3] 正在解析文档并生成XMind测试大纲: !FILE_NAME!

powershell -ExecutionPolicy Bypass -File "%~dp0parse_doc_helper.ps1" -FilePath "!FILE_PATH!" -ApiUrl "%API_URL%"

if %errorlevel% neq 0 (
    echo.
    echo 处理失败，请检查错误信息
    pause
    exit /b 1
)

echo.
echo ========================================
echo 处理完成！
echo ========================================
echo XMind文件已保存到文档所在目录
echo.

pause

