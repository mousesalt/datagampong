DATA GAMPONG VERSI WEB - REKAP INTERAKTIF

Database/raw data: DATA-GAMPONG.xlsx, sheet DATAPENDUDUK.

Urutan tampilan:
1. Statistik.
2. Data Penduduk + pencarian/filter.
3. Rekap Kelompok Umur.
4. Rekap Dusun/Lorong dan Pekerjaan.
5. Rekap Pendidikan.
6. Daftar Kepala Keluarga per Dusun.

Fitur interaktif:
- Tombol "Lihat Data" pada setiap kategori umur menampilkan siapa saja yang masuk kategori tersebut di DATA PENDUDUK.
- Tombol "Lihat Data" pada setiap pekerjaan menampilkan seluruh warga dengan pekerjaan tersebut.
- Tombol "Lihat Data" pada setiap pendidikan menampilkan seluruh warga dengan pendidikan tersebut.
- Rekap dusun juga dapat difilter dengan satu klik.
- Filter rekap otomatis mengosongkan pencarian/filter utama agar hasil benar-benar seluruh anggota kategori.
- Tersedia tombol "Hapus Filter Rekap".
- Sel kosong diabaikan; tidak dibuat menjadi kategori.
- Kepala keluarga dikenali dari SHDK yang berisi KEPALA KELUARGA, KEPALA, atau KK (setelah normalisasi spasi/tanda baca).
- Setiap dusun pada Daftar Kepala Keluarga memiliki ekspor Excel tersendiri dengan kolom tanda tangan.

Teknologi:
- HTML + CSS + JavaScript.
- Excel sebagai database/raw data.
- Tidak memakai PHP, SQL, SQLite, MySQL, XAMPP.

Ekspor DATA PENDUDUK utama menggunakan XLSX (bukan CSV) agar NO.KK dan NIK disimpan sebagai TEXT. Dengan demikian Excel tidak otomatis mengubah ID menjadi scientific notation atau membulatkan digit terakhir. Tombol yang sebelumnya bertuliskan CSV sekarang menjadi Excel.
