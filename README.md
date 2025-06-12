# Mulmed AI v2 (AI Decision Edition)

Mulmed AI v2 adalah versi lanjutan dari sistem otomatis pengubah slide PowerPoint berbasis suara. Pada versi ini, keputusan untuk mengganti slide tidak hanya berdasarkan deteksi teks, tetapi diputuskan secara langsung oleh AI (Google Gemini atau ChatGPT) berdasarkan kesesuaian ucapan dengan isi slide.

## Fitur Unggulan

1. Keputusan utama diambil oleh Gemini (bisa diganti ke ChatGPT)

2. Transkripsi real-time dengan Google Speech Recognition

3. Evaluasi kecocokan ucapan dengan isi slide oleh AI

4. Kontrol PowerPoint melalui COM API (Windows only)

## Persyaratan (Requirements)

### Sistem:

* Sistem Operasi Windows

* PowerPoint terinstall (harus aktif dalam mode Slide Show)

# Instalasi

1. Clone repositori ini:
```
git clone https://github.com/Georgearr/Mulmed-AI-v2.git
cd Mulmed-AI
```

2. **(Opsional)** Buat dan aktifkan virtual environment:

```
python -m venv venv
venv\Scripts\activate
```

3. Install semua requirements:
```
pip install -r requirements.txt
```
4. Tambahkan API key:
```
# file .env\OPENAI_API_KEY=[Masukkan API Open AI]
GOOGLE_API_KEY=[Masukkan API Google Gemini]
```

# Konfigurasi API

* Anda dapat menggunakan **Google Gemini (default)** atau **OpenAI Chat GPT**.

* API key harus disimpan di file .env agar otomatis dikenali.

# Cara Menjalankan

1. Buka PowerPoint, mulai `Slide Show (F5)`.

2. Jalankan script:
```
python main.py
```
3. AI akan mengevaluasi dan mengganti slide secara otomatis.

# Mekanisme Kerja

1. Mikrofon menangkap ucapan.

2. Google Speech mengubahnya jadi teks.

3. Slide saat ini diambil dan teksnya dibaca.

4. AI (Gemini) menganalisis kecocokan antara ucapan dan teks slide.

5. Jika cocok (menurut AI), slide akan berpindah.

# Catatan

1. Pastikan PowerPoint aktif di jendela utama.

2. Gunakan input device berkualitas **baik**.

3. Gemini bisa diganti dengan Chat GPT jika diinginkan.

## Lisensi

Proyek ini open-source dan dapat dimodifikasi sesuai kebutuhan. AI ini sebenarnya dibuat untuk kebutuhan Multimedia, tetapi AI ini juga bebas digunakan untuk edukasi, presentasi, demo AI, dan lain-lain.

## Kontak

Dikembangkan oleh **George A. T.**
Untuk pertanyaan, bisa email georgearrev@gmail.com, atau DM Instagram dengan akun [@george_arrev_turnip](https://www.instagram.com/george_arrev_turnip)

Terimakasih.

Copyright © 2025 - George A. T.
