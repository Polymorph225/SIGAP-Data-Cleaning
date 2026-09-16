# 🧹 Data Cleaning SIGAP Bojonegoro

Aplikasi Streamlit untuk **membersihkan** dan **menggabungkan** berkas Rekam Medis Elektronik (RME) Puskesmas, lalu mengubahnya menjadi berkas siap unggah ke [SIGAP-Bojonegoro](https://github.com/Polymorph225/SIGAP-Bojonegoro).

Dipakai di UPT Puskesmas Purwosari, Kabupaten Bojonegoro.

---

## ✨ Apa yang bisa dilakukan

| Fitur | Keterangan |
|---|---|
| 🔗 **Gabung banyak berkas** | Unggah beberapa berkas bulanan sekaligus, hasilnya menjadi satu berkas |
| 📅 **Urut menurut tanggal** | Diurutkan berdasarkan tanggal kunjungan di dalam berkas, bukan nama berkasnya |
| 🧾 **Baca .xls dari RME** | Ekspor RME berekstensi `.xls` yang isinya sebenarnya teks berpemisah tab tetap terbaca |
| 🗑️ **Buang baris kembar** | Baris yang persis sama antarberkas dibuang, jumlahnya tetap dilaporkan |
| ✂️ **Trailing koma** | Hapus ` ,` di akhir sel (No RM, NIK, No Penjamin) |
| 🔠 **Normalisasi teks** | UPPERCASE Nama, Title Case Desa |
| 📋 **Isi nilai kosong** | Kolom kosong diisi nilai pengganti yang bisa diatur sendiri |
| ❌ **Hapus kolom** | Pilih kolom mana yang tidak perlu ikut |
| 📤 **Dua bentuk unduhan** | Arsip lengkap, atau 8 kolom siap unggah ke SIGAP |

---

## 📤 Dua bentuk keluaran

**📚 Arsip lengkap** — seluruh kolom apa adanya, termasuk nama, NIK, dan alamat pasien. Simpan di komputer puskesmas, jangan dibawa keluar.

**🏥 Siap unggah ke SIGAP** — hanya delapan kolom, tanpa satu pun kolom identitas:

```
tanggal_kunjungan, no_rm, umur, jenis_kelamin, poli, diagnosa, pembiayaan, desa
```

Pemetaan kolomnya:

| Kolom RME | Kolom SIGAP |
|---|---|
| `Tgl` | `tanggal_kunjungan` |
| `No RM` | `no_rm` |
| `Usia` | `umur` (diambil angka tahunnya saja) |
| `L/LP` | `jenis_kelamin` |
| `Unit` | `poli` |
| `Diagnosis` | `diagnosa` |
| `Cara Bayar` | `pembiayaan` |
| `Desa` | `desa` |

Keduanya tersedia sebagai **.xlsx** maupun **.csv**.

---

## 📁 Format yang didukung

- **Masukan:** `.xlsx`, `.xls` (Excel asli maupun teks berpemisah), `.csv`
- **Keluaran:** `.xlsx` atau `.csv`

Semua kolom dibaca sebagai teks, sehingga nomor rekam medis dan NIK tidak berubah menjadi angka dan nol di depannya tidak hilang.

---

## 📅 Catatan soal tanggal

Tanggal tidak diurai dengan `dayfirst`. Pandas menebak format dari nilai pertama, sehingga `2026-04-01` bisa terbaca sebagai 1 April atau 4 Januari — dan baris yang tanggalnya di atas 12 akan gagal terurai diam-diam. Aplikasi ini mencoba setiap format yang mungkin satu per satu, lalu memakai yang paling banyak berhasil. Baris yang tanggalnya tetap tidak terbaca ditaruh di urutan paling akhir dan jumlahnya dilaporkan, bukan dibuang.

---

## 💻 Menjalankan di komputer sendiri

```bash
pip install -r requirements.txt
streamlit run app.py
```

Lalu buka `http://localhost:8501`.

---

## 🚀 Deploy ke Streamlit Community Cloud

1. Buka [share.streamlit.io](https://share.streamlit.io), masuk dengan akun GitHub
2. Klik **New app**, pilih repositori ini
3. **Main file path:** `app.py`
4. Klik **Deploy!**

Aplikasi aktif dalam satu sampai dua menit.

---

## 🔒 Data pasien

Berkas yang diunggah hanya berada di memori selama sesi berlangsung — tidak ditulis ke penyimpanan dan tidak dikirim ke mana pun.

Jangan pernah memasukkan berkas data pasien ke dalam repositori ini. Berkas keluaran **arsip lengkap** memuat nama dan NIK; hanya berkas **siap SIGAP** yang aman dibawa keluar puskesmas.

---

## 📋 Isi repositori

```
├── app.py            ← Aplikasi Streamlit
├── requirements.txt  ← Daftar pustaka Python
├── .gitignore        ← Penjaga agar berkas data tidak ikut ter-commit
└── README.md         ← Berkas ini
```
