# Sistem Rekod e-Hadir SMK Tobiar

Sistem rekod kehadiran berasaskan web untuk Sekolah Menengah Kebangsaan Tobiar.

## Kandungan

- `index.html` - Sistem e-Hadir untuk GitHub Pages.
- `logo-smk-tobiar.png` - Lencana sekolah.
- `apps-script/Code.gs` - Kod backend Google Apps Script.

## Terbitkan melalui GitHub Pages

1. Cipta repositori baharu di GitHub.
2. Muat naik semua kandungan pakej ini ke bahagian akar repositori.
3. Buka **Settings > Pages**.
4. Pada **Build and deployment**, pilih **Deploy from a branch**.
5. Pilih branch `main`, folder `/ (root)`, kemudian tekan **Save**.
6. Tunggu GitHub menyediakan pautan laman anda.

## Google Apps Script

Sistem telah dikonfigurasikan menggunakan URL Web App Google Apps Script yang sedia ada. Jika `Code.gs` dikemas kini:

1. Tampalkan kandungan `apps-script/Code.gs` ke projek Apps Script.
2. Pilih **Deploy > Manage deployments**.
3. Cipta versi baharu dan terbitkan sebagai Web App.
4. Tetapkan akses kepada **Anyone** supaya laman GitHub Pages boleh menghantar dan membaca rekod.

Senarai kelas, murid, subjek dan peserta kelas tambahan dibaca daripada fail CSV Google Sheets yang telah diterbitkan.
