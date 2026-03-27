@echo off
echo ============================================================
echo   EXR Owner Financials Extractor - Web App
echo ============================================================
echo.
echo Installing / verifying dependencies...
python -m pip install -r requirements.txt --quiet
echo.
echo Starting Streamlit...  The app will open in your browser.
echo Press Ctrl+C in this window to stop the app.
echo.

python -m streamlit run app.py

pause
