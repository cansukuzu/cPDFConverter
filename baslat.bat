@echo off
title cPDF Converter
echo ===================================================
echo           cPDF Converter Baslatiliyor...
echo ===================================================
echo.
python -m streamlit run app.py --server.port 8501
pause
