@echo off
echo ========================================================
echo PARETO - Kredi Tahsis ve ESG Degerlendirme Sistemi
echo ========================================================
echo Gerekli kutuphaneler kontrol ediliyor ve yukleniyor...
pip install -r requirements.txt
echo.
echo PARETO Sistemi Baslatiliyor... Lutfen bekleyin...
streamlit run pareto.py
pause
