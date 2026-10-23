# Dashboard Mentoring Ashabul Istiqomah

Tampilan dashboard di GitHub Pages, data tetap di Google Spreadsheet lewat Apps Script.

## Isi repo

- `index.html` — tampilan dashboard
- `config.js` — URL web app Apps Script
- `Code.gs` — salinan kode Apps Script (API ke spreadsheet), dipaste manual ke editor Apps Script

## Setup

1. Buka spreadsheet mentoring > Ekstensi > Apps Script.
2. Ganti isi `Code.gs` dengan file `Code.gs` di repo ini, simpan. File HTML `Index` boleh dihapus.
3. Project Settings (ikon gerigi) > Script Properties > Add script property:
   nama `KODE_PENGURUS`, nilai kode rahasia pengurus. Kode ini sengaja tidak ditulis di repo.
4. Deploy > Manage deployments > edit deployment yang ada (ikon pensil):
   - Version: New version
   - Execute as: Me
   - Who has access: **Anyone**
5. Copy URL yang berakhiran `/exec`, tempel ke `config.js`.
6. Di GitHub: Settings > Pages > Source: Deploy from a branch > `main` / root.

Kalau ada perubahan di `Code.gs`, tetap harus deploy versi baru di Apps Script. Perubahan tampilan cukup push ke GitHub.

## Keamanan

- Kode pengurus hanya disimpan di Script Properties, jangan pernah ditulis ke file di repo.
- Salah memasukkan kode 8 kali berturut-turut mengunci mode pengurus selama 10 menit.
