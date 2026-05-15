@echo off
echo Installing required packages...
pip install -r requirements.txt

echo.
echo Starting CSP Exit Tracker Dashboard...
echo Open your browser at: http://localhost:8501
echo Press Ctrl+C to stop.
echo.
python -m streamlit run dashboard.py
pause
