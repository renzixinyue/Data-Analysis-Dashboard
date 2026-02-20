@echo off
echo ========================================================
echo 正在启动数据可视化大屏服务器...
echo ========================================================
echo.
echo 请在浏览器中访问以下地址：
echo http://localhost:8000/
echo.
echo (按 Ctrl+C 可停止服务器)
echo.

cd dashboard
python -m http.server 8000
pause
