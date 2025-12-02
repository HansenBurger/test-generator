@echo off
chcp 65001 >nul

if not exist "backend\.venv" (
    cd backend
    python -m venv .venv
    call .venv\Scripts\activate.bat
    pip install --upgrade pip >nul 2>&1
    pip install --only-binary :all: pydantic
    pip install fastapi uvicorn[standard] python-multipart python-docx
    cd ..
)

if not exist "frontend\node_modules" (
    cd frontend
    call npm install
    cd ..
)

cd backend
start "Backend" cmd /k "call .venv\Scripts\activate.bat && python main.py"

timeout /t 2 /nobreak >nul

cd ..\frontend
start "Frontend" cmd /k "npm run dev"
