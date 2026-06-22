@echo off
chcp 65001 >nul
echo ============================================
echo   期末复习 - 本地 HTTP 服务器
echo   讲义目录: _extracted
echo   访问地址: http://localhost:8888
echo ============================================
echo.

cd /d "%~dp0..\..\..\_extracted"

:: 安装 MathJax（如果还没有）
if not exist "node_modules\mathjax\es5\tex-svg.js" (
    echo [正在安装 MathJax 本地渲染引擎...]
    call npm install
    echo [MathJax 安装完成]
    echo.
)

echo 按 Ctrl+C 停止服务器
echo ============================================
echo.

python -m http.server 8888
