@echo off
echo Installing required packages...
pip install Flask PyMuPDF pdfplumber
echo.
echo Creating folders...
mkdir templates 2>nul
mkdir uploads 2>nul
mkdir outputs 2>nul
echo.
echo Setup complete!
echo.
echo To run the application, type: python app.py
pause