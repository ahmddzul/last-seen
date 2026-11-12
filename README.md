# Last Seen

Game detektif di HP. HP adikmu ditemukan di peron stasiun. Adikmu tidak.

Mainkan: https://ahmddzul.github.io/last-seen/

- Season 1 lengkap: Prolog, Bab 1 sampai 5, tiga ending.
- Kasus Harian: 7 kasus bergilir, streak, dan kartu hasil untuk dibagikan.
- Seluruh kode game ada di `index.html`. Foto ada di `assets/p/`, jadi buka lewat GitHub Pages atau server lokal.

Dibuat oleh [@ahmddzul](https://instagram.com/ahmddzul). Dukung lewat [Saweria](https://saweria.co/ahmddzul).

## Data dan dasbor admin

- `apps-script/Code.gs`: kode Google Apps Script untuk papan skor, statistik, ulasan, dan pengumuman.
- Akun pemain: nama pengguna + PIN 6 angka, progres tersinkron lewat Apps Script (tab Pemain).
- Dasbor admin dibuka dari dalam game (ikon perisai di beranda) oleh akun yang terdaftar di Script Property `ADMIN_USERS`. Isinya ada di `admin.html`.
- Alamat Web app Apps Script dipasang di konstanta `SHEET_URL` pada `index.html` dan `admin.html`.

