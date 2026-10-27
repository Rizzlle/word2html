# Word ke HTML untuk T&C

Konverter terminal berbasis Python (tanpa paket tambahan) untuk mengubah satu file `.docx` menjadi satu halaman HTML utuh. Urutan teks, judul, daftar dan penomorannya, pemformatan dasar, tautan, serta tabel dipertahankan. Nomor halaman Word diabaikan.

## Pemakaian dengan pemilih file

Di Windows, buka `wizard.pyw` dengan klik dua kali (atau jalankan `pythonw wizard.pyw`). Pilih satu atau beberapa file `.docx` sekaligus (gunakan `Ctrl`/`Shift` pada pemilih file); semuanya diproses ke folder `outputs/`. Jika ada nama hasil yang sudah ada, wizard meminta satu konfirmasi: **Ya** menimpa hasil lama, **Tidak** melewati file tersebut dan tetap memproses file lain. Setelah selesai, jumlah yang berhasil, dilewati, dan gagal ditampilkan.

Jika dua dokumen terpilih mempunyai nama yang sama, hasil kedua diberi akhiran `(2)` agar tidak saling menimpa.

Pemilih file yang sama juga bisa dibuka dari terminal:

```powershell
python word2html.py
```

## Pemakaian lewat terminal

```powershell
python word2html.py "sample.docx"
```

Hasil default tersimpan di folder `outputs/` di dalam proyek, dengan nama file sumber dan ekstensi `.html`. Folder hasil dibuat otomatis. Untuk menentukan lokasi hasil atau menimpa file yang sudah ada:

```powershell
python word2html.py "dokumen.docx" -o "hasil/tnc.html"
python word2html.py "dokumen.docx" --force
```

Jika file hasil sudah ada, program menolak menimpanya kecuali `--force` diberikan. Folder tujuan untuk opsi `-o` juga dibuat otomatis.

## Aturan konversi

- Style Heading 1–9 dipetakan ke `<h1>`–`<h6>`; tingkat di atas enam menjadi `<h6>`.
- Dokumen yang memakai judul manual, seperti contoh, dikenali dari posisi, huruf kapital, dan teks tebal. Paragraf yang tidak cukup jelas tetap menjadi `<p>`.
- Font utama dan ukurannya diambil dari pengaturan dokumen Word; font khusus pada potongan teks tertentu juga dipertahankan. Jika font tersebut tidak tersedia di perangkat pembuka HTML, browser memakai font pengganti.
- Daftar Word menjadi `<ol>` atau `<ul>`. Label nomor dibaca dari definisi penomoran Word dan ditampilkan apa adanya, termasuk nomor yang dimulai ulang.
- URL `http://` dan `https://` yang tertulis sebagai teks juga menjadi tautan.
- Penanda pergantian halaman, header/footer, gambar, dan catatan kaki tidak ikut dikonversi. File contoh tidak mengandung elemen-elemen tersebut.

Karena judul manual tidak memiliki penanda semantik di Word, periksa tingkat judul pada HTML setiap kali memakai dokumen dengan tata letak baru. Program tidak menyisipkan catatan ke HTML.

## Pemeriksaan

```powershell
python -m unittest discover -s tests
```
