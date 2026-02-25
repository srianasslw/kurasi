@echo off
cd /d "%~dp0"

:: 1. Bersihkan sisa-sisa lama
taskkill /f /im python.exe >nul 2>&1

:: 2. Perintahkan Windows untuk membuka browser (setelah jeda 6 detik)
:: Kita taruh ini di depan supaya tidak terhalang oleh macetnya streamlit
start /b cmd /c "timeout /t 6 /nobreak >nul && start http://127.0.0.1:8501"

:: 3. Jalankan mesin Python (ini yang bikin macet, tapi tidak masalah karena browser sudah dijadwalkan)
echo Menyalakan mesin... silakan tunggu browser terbuka otomatis...
python -m streamlit run app.py --server.port 8501 --server.headless true