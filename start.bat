@echo off
REM ============================================================================
REM  GSIS-P / SMIR Reconciliation  -  launch on a shared PC (browser access)
REM  Other machines on the LAN open  http://THIS-PC-IP:8502  in a browser.
REM  Fully offline: no internet is used at runtime.
REM ============================================================================
title GSIS-P / SMIR Reconciliation
cd /d "%~dp0"

REM Never phone home to Hugging Face at runtime (models are bundled locally).
set HF_HUB_OFFLINE=1
set TRANSFORMERS_OFFLINE=1
set HF_HUB_DISABLE_SYMLINKS=1
REM Disable Streamlit's file watcher — it crashes when it introspects PyTorch
REM (torch._classes). Required whenever torch is installed.
set STREAMLIT_SERVER_FILE_WATCHER_TYPE=none

echo ============================================
echo    GSIS-P / SMIR Reconciliation
echo ============================================
echo   On this PC:      http://localhost:8502
for /f "tokens=2 delims=:" %%a in ('ipconfig ^| findstr /c:"IPv4"') do echo   On the network:  http://%%a:8502
echo.
echo   Press Ctrl+C to stop.
echo.

python -m streamlit run reconcile.py --server.port 8502 --server.address 0.0.0.0 --server.headless true --server.fileWatcherType none
pause
