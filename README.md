# 🏁 Tugas Akhir (TA) - Final Project

**Nama Mahasiswa**: Muhammad Razan Athallah  
**NRP**: 5025211008  
**Judul TA**: PENGEMBANGAN SIMULASI AKTIVITAS HARIAN PENGGUNA UNTUK KOMPUTER _SANDBOX_ YANG MENGEKSEKUSI _MALWARE_  
**Dosen Pembimbing**: Baskoro Adi Pratomo, S.Kom., M.Kom., Ph.D.    
**Dosen Ko-pembimbing**: Hudan Studiawan, S.Kom., M.Kom., Ph.D.

---

## 📺 Demo Program  

[![Demo Program](https://github.com/user-attachments/assets/0a9846ff-7b9d-4c73-a0e2-7c73a5ead899)](https://www.youtube.com/watch?v=Z1EMP3ihMXA)  
*Klik gambar di atas untuk menonton demo*

---

## 🛠 Panduan Instalasi & Menjalankan Program  

### Prasyarat  
Program dapat dijalankan pada komputer dengan sistem operasi Windows 11 dengan resolusi layar berikut:
- 1920×1080 (skala 100%, 125%, dan 150%)
- 1366×768 (skala 100% dan 125%)
- 1280×720 (skala 100%)

Dependensi yang dibutuhkan:
- Python 3.10+

Selain aplikasi bawaan yang sudah terinstal, berikut adalah aplikasi tambahan yang perlu diinstal secara manual:
- Google Chrome
- Mozilla Firefox
- Microsoft Excel
- Microsoft Word

### Langkah-langkah  
1. **Clone Repositori**  
   ```bash
   git clone https://github.com/Informatics-ITS/ta-athraz.git
   ```
2. **Buat Lingkungan Virtual dan Instal Dependensi**
   ```bash
   python -m venv <venv_name>

   <venv_name>\Scripts\activate       # For Windows (Command Prompt)
   <venv_name>\Scripts\Activate.ps1   # For Windows (PowerShell)

   pip install -r requirements.txt
   ```
3. **Konfigurasi Simulasi**  
   Konfigurasi terletak pada file `src/configs/scenarios.yaml`, contohnya sebagai berikut:  
   ```yaml
   execution_mode: "sequential"
   repeat: false
   exit_on_error: true
   scenarios:
     - scenario:
       - name: "Gmail"
         browser: "Google Chrome"
         methods:
           - method: "open"
             delay: 2
           - method: "next_email"
             delay: 2
             args:
               count: 5
     - scenario:
       - name: "Google Forms"
         browser: "Google Chrome"
         methods:
          - method: "fill_form"
            delay: 2
            args:
              url: "https://forms.gle/1wCpoSWBjQqVvtVv5"
              answers: ["short answer", "lorem ipsum", "paragraph answer"]
   ```
   Terdapat 4 properti dengan kegunaanya dalam mengatur simulasi sebagai berikut:
   <table>
      <thead>
         <tr>
            <th>Properti</th>
            <th>Nilai yang dapat diisikan</th>
            <th>Dampak</th>
            <th>Nilai Bawaan</th>
         </tr>
      </thead>
      <tbody>
         <tr>
            <td rowspan=2>execution_mode</td>
            <td>"random"</td>
            <td>Program menjalankan skenario-skenario secara acak</td>
            <td rowspan=2>"sequential"</td>
         </tr>
         <tr>
            <td>"sequential"</td>
            <td>Program menjalakan skenario sesuai urutan, dari atas ke bawah</td>
         </tr>
         <tr>
            <td rowspan=2>repeat</td>
            <td>true</td>
            <td>Program akan mengulang eksekusi seluruh skenario setelah semuanya selesai dijalankan</td>
            <td rowspan=2>false</td>
         </tr>
         <tr>
            <td>false</td>
            <td>Program hanya menjalankan seluruh skenario satu kali</td>
         </tr>
         <tr>
            <td rowspan=2>exit_on_error</td>
            <td>true</td>
            <td>Program langsung berhenti jika terjadi kesalahan</td>
            <td rowspan=2>true</td>
         </tr>
         <tr>
            <td>false</td>
            <td>Program tetap melanjutkan eksekusi meskipun terjadi kesalahan</td>
         </tr>
         <tr>
            <td>scenarios</td>
            <td>scenario</td>
            <td>Berisi skenario-skenario aktivitas yang akan disimulasikan</td>
            <td>-</td>
         </tr>
      </tbody>
   </table>
   Setiap skenario dapat terdiri dari satu atau lebih aplikasi atau aktivitas, dimana setiap aplikasi atau aktivitas memiliki beberapa properti seperti nama, metode-metode yang akan dijalankan, delay antar metode, dan argumen-argumen metode. Berikut aplikasi dan aktivitas yang dapat
   disimulasikan:
   <table>
      <thead>
         <tr>
            <th>No</th>
            <th>Kelompok Aplikasi atau Aktivitas</th>
            <th>Nama Aplikasi atau Aktivitas</th>
         </tr>
      </thead>
      <tbody>
         <tr>
            <td rowspan=3>1.</td>
            <td rowspan=3>Aplikasi Berbasis Website</td>
            <td>Gmail</td>
         </tr>
         <tr>
            <td>Google Forms</td>
         </tr>
         <tr>
            <td>YouTube</td>
         </tr>
         <tr>
            <td rowspan=2>2.</td>
            <td rowspan=2>Browser</td>
            <td>Google Chrome</td>
         </tr>
         <tr>
            <td>Mozilla Firefox</td>
         </tr>
         <tr>
            <td rowspan=4>3.</td>
            <td rowspan=4>Aplikasi Native</td>
            <td>Microsoft Excel</td>
         </tr>
         <tr>
            <td>Microsoft Paint</td>
         </tr>
         </tr>
         <tr>
            <td>Microsoft Word</td>
         </tr>
         </tr>
         <tr>
            <td>Notepad</td>
         </tr>
         <tr>
            <td rowspan=2>4.</td>
            <td rowspan=2>Aplikasi Sistem atau File</td>
            <td>Command Prompt</td>
         </tr>
         <tr>
            <td>File Explorer</td>
         </tr>
         <tr>
            <td rowspan=2>5.</td>
            <td rowspan=2>Otomasi</td>
            <td>Fungsi AutoIt</td>
         </tr>
         <tr>
            <td>Skrip Selenium</td>
         </tr>
      </tbody>
   </table>  
4. **Jalankan Program**
   ```bash
   cd src
   python main.py
   ```

---

## ✅ Validasi

Pastikan proyek memenuhi kriteria berikut sebelum submit:
- Source code dapat di-build/run tanpa error
- Video demo jelas menampilkan fitur utama
- README lengkap dan terupdate
- Tidak ada data sensitif (password, API key) yang ter-expose

---

## ⁉️ Pertanyaan?

Hubungi:
- Penulis: [athrazan2004@gmail.com]
- Pembimbing Utama: [baskoro@its.ac.id]
