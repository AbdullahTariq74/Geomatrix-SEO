@echo off
cd /d "%~dp0"
echo Installing dependencies...
pip install -r requirements.txt -q
echo.
echo Starting GeoMatrix SEO Generator...
streamlit run app.py --server.port 8501
pause
