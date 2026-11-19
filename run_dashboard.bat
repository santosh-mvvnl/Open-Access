@echo off
title Open Access Consumers & Billing Revenue Analysis Dashboard
echo ===============================================================================
echo Starting Open Access Consumers & Billing Revenue Analysis Dashboard...
echo ===============================================================================
cd /d "%~dp0"
python -m streamlit run app.py --server.port 8503 --server.headless false
pause
