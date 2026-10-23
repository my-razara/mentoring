# Dashboard Mentoring Ashabul Istiqomah

Tampilan dashboard di GitHub Pages, data tetap di Google Spreadsheet lewat Apps Script.

## Isi repo

- `index.html` — tampilan dashboard
- `config.js` — URL web app Apps Script
- `Code.gs` — salinan kode Apps Script (API ke spreadsheet), dipaste manual ke editor Apps Script

## Setup

1. Buka spreadsheet mentoring > Ekstensi > Apps Script.
2. Ganti isi `Code.gs` dengan file `Code.gs` di repo ini, simpan. File HTML `Index` boleh dihapus.
3. Deploy > Manage deployments > edit deployment yang ada (ikon pensil):
   - Version: New version
   - Execute as: Me
   - Who has access: **Anyone**
4. Copy URL yang berakhiran `/exec`, tempel ke `config.js`.
5. Di GitHub: Settings > Pages > Source: Deploy from a branch > `main` / root.

Kalau ada perubahan di `Code.gs`, tetap harus deploy versi baru di Apps Script. Perubahan tampilan cukup push ke GitHub.
