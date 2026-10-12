@echo off
chcp 936 >nul
setlocal
title DOE 工艺验证报告 - 拖拽启动器

rem ==================================================================
rem  用法：把 DOE_data.xlsx（或项目文件夹）拖到本文件的图标上
rem        双击 = 用 parameter.py 里的默认配置
rem ==================================================================

set "HERE=%~dp0"
set "SCRIPT=%HERE%launcher.py"

if not exist "%SCRIPT%" (
    echo [错误] 找不到 launcher.py，请确认它和 run.bat 在同一个目录。
    echo         当前目录：%HERE%
    echo.
    goto :done
)

rem ---- 找一个真正装了依赖的 Python ----
rem PATH 上第一个 python 不一定是能用的那个（本机就出现过空环境），
rem 所以这里逐个实测 import，而不是只看它存不存在。
set "PYEXE="
set "PYARGS="

py -3 -c "import pptx,matplotlib,pandas,openpyxl" >nul 2>&1
if not errorlevel 1 (
    set "PYEXE=py"
    set "PYARGS=-3"
)

if not defined PYEXE (
    python -c "import pptx,matplotlib,pandas,openpyxl" >nul 2>&1
    if not errorlevel 1 set "PYEXE=python"
)

if not defined PYEXE (
    python3 -c "import pptx,matplotlib,pandas,openpyxl" >nul 2>&1
    if not errorlevel 1 set "PYEXE=python3"
)

if not defined PYEXE if exist "%LOCALAPPDATA%\Programs\Python\Python313\python.exe" (
    "%LOCALAPPDATA%\Programs\Python\Python313\python.exe" -c "import pptx,matplotlib,pandas,openpyxl" >nul 2>&1
    if not errorlevel 1 set "PYEXE="%LOCALAPPDATA%\Programs\Python\Python313\python.exe""
)

if not defined PYEXE if exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" (
    "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" -c "import pptx,matplotlib,pandas,openpyxl" >nul 2>&1
    if not errorlevel 1 set "PYEXE="%LOCALAPPDATA%\Programs\Python\Python312\python.exe""
)

if not defined PYEXE (
    echo.
    echo [错误] 没有找到能用的 Python 环境。
    echo        需要 python-pptx / matplotlib / pandas / openpyxl。
    echo        在项目目录执行： pip install -r requirements.txt
    echo.
    goto :done
)

rem ---- 干活 ----
cd /d "%HERE%"

rem 固定 Python 的输出编码，和上面的 chcp 936 对齐。
rem 不写这句的话，万一系统环境里存在 PYTHONUTF8=1，
rem Python 会改用 UTF-8 输出，中文在 GBK 控制台上就成了乱码。
set "PYTHONIOENCODING=gbk"

echo [Python] %PYEXE% %PYARGS%
echo.

%PYEXE% %PYARGS% "%SCRIPT%" %*

if errorlevel 1 (
    echo.
    echo [提示] 上面有错误，请向上翻看日志。
)

:done
echo.
if not "%DOE_NO_PAUSE%"=="1" pause
