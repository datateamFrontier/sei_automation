README

program_gabungan_v2

- Perbedaan dengan v1 adalah v2 ini direncanakan untuk data yang masuk dari caller lewat googleform
- ada tambahan kategori online chat
- Status program ini masih BELUM SELESAI


===========================================================
Keterangan file & folder
===========================================================
Folder
bobot			: excel bobot perhitungan per kategori SEI
data kuesioner  : data mentah yang didownload dari google form dalam bentuk csv
data mentah		: data mentah yang siap untuk olah indeks (menggunakan program)
gambar			: gambar-gambar yang dibutuhkan untuk laporan powerpoint
laporan			: output laporan powerpoint yang dihasilkan program
tabel output	: excel tabel-tabel output program


File
buat fungsi ppt.ipynb                       : notebook untuk develop fungsi umum untuk generate file report powerpoint
data_to_gform.gs                            : program untuk generate opsi jawaban brand pada googleform melalui googlesheets
Note koreksi laporan.txt                    : daftar catatan kesalahan output pada report powerpoint hasil generate program
olah_data.py                                : modul program berisi fungsi-fungsi untuk melakukan olah indeks
olah data indeks.ipynb                      : notebook untuk develop fungsi umum untuk generate tabel-tabel output
ppt_ccsei.py                                : modul program berisi fungsi-fungsi untuk generate report powerpoint
program_service_excelence_index_v1.ipynb    : notebook untuk develop gabungan fungsi powerpoint dan olah indeks
program_service_excelence_index_v1.py       : modul program / main program yang bisa dieksekusi untuk keseluruhan proses olah indeks & pembuatan report
transformasi data kuesioner.ipynb           : notebook untuk mentransformasikan data dari folder "data kuesioner" menjadi data yang siap diolah dan disimpan ke folder "data mentah" (status file ini masih belum selesai)



===========================================================
Panduan penggunaan / Cara kerja
===========================================================

1. Buat folder bernama 'data mentah' (jika belum ada) kemudian letakkan data yang akan diolah di dalam folder tersebut. Namai file data mentah dengan format diikuti dengan kategorinya, misalkan 'data mentah call center'.
2. Buat folder bernama 'laporan' (jika belum ada) untuk tempat penyimpanan ppt yang akan dihasilkan program.
3. Buat folder bernama 'tabel output' (jika belum ada) untuk tempat penyimpanan data excel hasil pengolahan program.
4. program dengan ekstensi '.ipynb' merupakan file notebook yang bisa dibuka di browser dan digunakan untuk melihat isi data di tiap tahapan (cek / proses development program).
5. program dengan ekstensi '.py' merupakan file python yang bisa dieksekusi di command prompt / terminal.
