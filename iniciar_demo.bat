@echo off
rem Gera a planilha fictícia e o dashboard e abre no navegador (dois cliques).
cd /d "%~dp0"
chcp 65001 >nul

if not exist ".venv\Scripts\python.exe" (
    echo Primeira vez: criando o ambiente e instalando o openpyxl...
    python -m venv .venv
    if errorlevel 1 (
        echo.
        echo Nao foi possivel criar o ambiente. Instale o Python em https://www.python.org/downloads/
        echo e marque "Add python.exe to PATH" na instalacao.
        pause
        exit /b 1
    )
    ".venv\Scripts\python.exe" -m pip install -q openpyxl
    if errorlevel 1 (
        echo.
        echo Falha ao instalar o openpyxl. Se a mensagem falar de "Long Path", mova a pasta para algo
        echo mais curto, como C:\projetos\dashboard-interativo, e tente de novo.
        pause
        exit /b 1
    )
)

".venv\Scripts\python.exe" gerar_planilha_demo.py
".venv\Scripts\python.exe" gerar_dashboard.py
if errorlevel 1 (
    pause
    exit /b 1
)
start "" "%~dp0dashboard.html"
