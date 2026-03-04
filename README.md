# KURASI (Smart Accuracy System)

**KURASI** adalah platform cerdas berbasis Python yang dirancang untuk mengoptimalkan pengelolaan data **SPM** melalui proses penyaringan otomatis. Sistem ini secara efektif mengeliminasi data ganda (duplikat) dan memastikan akurasi informasi secara instan, transparan, dan akuntabel.

---

## 🚀 Langkah-Langkah Instalasi & Menjalankan

Ikuti panduan ini untuk menyiapkan lingkungan kerja di perangkat Anda.

### 1. Instalasi Git
* Unduh Git di [git-scm.com](https://git-scm.com/downloads).
* Jalankan installer dan klik **Next** hingga selesai.
* Verifikasi dengan mengetik `git --version` di Terminal/CMD.

### 2. Instalasi Python
* Unduh Python 3.12 atau versi terbaru di [python.org](https://www.python.org/downloads/).
* **PENTING:** Saat instalasi, centang kotak **"Add Python to PATH"**.
* Verifikasi dengan mengetik `python --version` atau `py --version` di Terminal.

### 3. Persiapan Proyek
Buka Terminal atau PowerShell, lalu arahkan ke folder tujuan (misal: htdocs) dan lakukan cloning:
```bash
cd C:\xampp\htdocs
git clone [https://github.com/srianasslw/kurasi.git](https://github.com/srianasslw/kurasi.git)
cd kurasi

### Jalankan perintah berikut untuk menginstal semua kebutuhan aplikasi:
py -m pip install streamlit pandas openpyxl pywin32

### Jalankan perintah berikut untuk membuka aplikasi di browser:
py -m streamlit run app.py
