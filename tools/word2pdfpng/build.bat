@echo off
setlocal EnableExtensions

rem 一键用 MSVC 编译 word2pdfpng.exe（静态 CRT /MT，无控制台窗口）
cd /d "%~dp0"

set "VSWHERE=%ProgramFiles(x86)%\Microsoft Visual Studio\Installer\vswhere.exe"
if not exist "%VSWHERE%" (
    echo 未找到 vswhere，请安装 Visual Studio 并勾选「使用 C++ 的桌面开发」。
    exit /b 1
)

for /f "usebackq delims=" %%i in (`"%VSWHERE%" -latest -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 -property installationPath`) do set "VSINSTALL=%%i"
if not defined VSINSTALL (
    echo 未找到带 VC 工具的 Visual Studio 安装。
    exit /b 1
)

call "%VSINSTALL%\VC\Auxiliary\Build\vcvars64.bat" >nul
if errorlevel 1 (
    echo vcvars64 初始化失败。
    exit /b 1
)

set SRC=main.cpp word_export.cpp pdf_to_png.cpp util.cpp
set OUT=word2pdfpng.exe

echo 正在编译 %OUT% ...
cl /nologo /EHsc /O2 /MT /std:c++17 /permissive- /utf-8 ^
    /DUNICODE /D_UNICODE ^
    %SRC% ^
    /Fe:%OUT% ^
    /link /SUBSYSTEM:WINDOWS /ENTRY:wWinMainCRTStartup ^
    ole32.lib oleaut32.lib uuid.lib comdlg32.lib shlwapi.lib windowsapp.lib

if errorlevel 1 (
    echo 编译失败。
    exit /b 1
)

echo 成功：%CD%\%OUT%
del /q *.obj 2>nul
exit /b 0
