# Audit LDP 2.9.0

Acuan: Code.gs dan Index.html terbaru dengan segmentasi B2B/B2C. File aplikasi dipertahankan; perubahan ini menyiapkan audit, bukan deployment atau perubahan rumus.

## SonarCloud

Proyek yang sudah terhubung: dafgil09-dot_LDP. Repo sebelumnya hanya memuat HTML lama dan TXT duplikat. Workflow SonarScanner gagal karena projectKey dan organization kosong, sementara aplikasi GitHub SonarCloud sudah melakukan automatic analysis.

Audit menggunakan automatic analysis yang sudah aktif. Workflow GitHub memverifikasi source dan mengambil hasilnya, tanpa menjalankan SonarScanner kedua.

- Code.gs adalah backend asli; sonar-src/Code.js merupakan salinan persis agar dikenali sebagai JavaScript.
- Index.html adalah tampilan asli; sonar-src/Index.js menyalin JavaScript inline dengan nomor baris HTML dipertahankan.
- sonar-src/source-manifest.json mencatat SHA-256 dan pemetaan source.
- .sonarcloud.properties membatasi cakupan; TXT duplikat tidak dianalisis.
- Setelah source berubah, jalankan node scripts/prepare-sonar.cjs lalu commit mirror. CI menolak mirror yang tertinggal.

Hasil tersedia pada summary workflow dan artifact ldp-audit-<run-id>. Pengambil laporan memeriksa check SonarCloud untuk SHA source yang tepat serta keberadaan kedua mirror dalam cakupan analisis. Quality Gate merah tetap dilaporkan.

Token yang sudah ada tetap digunakan melalui GitHub Actions secret. Token tidak dimasukkan ke source atau artifact; workflow dari fork tidak diberi secret.

## Temuan performa dari pemeriksaan source

Temuan ini terpisah dari rule SonarCloud dan belum mengukur durasi produksi.

| Prioritas | Bukti | Perbaikan |
|---|---|---|
| Tinggi | login_ membaca Participants lengkap untuk memeriksa NIK. | Gunakan indeks keanggotaan ringkas; pertahankan validasi akun terkini. |
| Tinggi | getParticipantFullDetail_ mengaktifkan liveAssessmentReads tanpa assessmentTargetNik. | Tampilkan ringkasan dahulu, lalu muat rincian kompetensi saat diperlukan. |
| Tinggi | Pembentukan read cache belum dikoordinasikan antar-eksekusi. | Uji lonjakan ketika cache kosong dan cegah pembacaan Sheets yang berulang. |
| Sedang | targetedAssessmentGrid_ membaca semua NIK lalu mengambil sampai 12 baris satu per satu. | Indeks posisi berbasis revisi dengan validasi ulang identitas. |
| Sedang | Cache assessment/dashboard pribadi masih memakai revisi global. | Gunakan revisi tabel dependensi agar penulisan job/audit tidak menggugurkan cache nilai. |

## Validasi dan batasan

Validasi lokal memeriksa sintaks serta kesesuaian mirror. Aplikasi Google tidak dijalankan. Static analysis tidak mengukur kondisi Sheets, formula, waktu login, atau 200 pengguna bersamaan. Target tersebut membutuhkan uji deployment dan log pembacaan/penulisan, cache hit/miss, serta waktu tunggu lock.
