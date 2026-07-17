=== MODUL ARKIB DATA KEHADIRAN ===

FUNGSI UTAMA:
1. arkibBulanan()       - Auto arkib data bulan LEPAS (dipicu trigger 1hb 1:00 AM)
2. pasangTriggerArkib() - Pasang trigger auto. Run SEKALI sahaja.
3. ujiArkibSekarang()   - Test: arkib data bulan lepas serta-merta.
4. ujiArkibSemua()      - TEST PENUH: pindah SEMUA data (semua bulan)
                          dari GURU, AKP, KEBERSIHAN, PELAWAT, TIDAK HADIR
                          ke sheet ARKIB_DATA. Sheet aktif akan dikosongkan
                          (header dikekalkan). ARKIB_DATA auto-create.

CARA JALANKAN ujiArkibSemua:
1. Buka Apps Script editor projek anda.
2. Ganti Code.gs dengan fail baharu ini.
3. Simpan (Ctrl+S).
4. Pilih fungsi 'ujiArkibSemua' dari dropdown atas.
5. Tekan ▶ Run. Beri kebenaran jika diminta.
6. View > Logs untuk lihat laporan.
7. Buka Google Sheets - sheet 'ARKIB_DATA' akan terisi.

ENDPOINT FRONTEND:
- getArkib        : Ambil rekod arkib (filter: bulan, kategori, nama)
- getBulanArkib   : Senarai bulan tersedia dalam arkib

STRUKTUR ARKIB_DATA:
KATEGORI | BULAN_ARKIB | TARIKH | NAMA | JAWATAN/KELAS |
MASA MASUK | MASA KELUAR | STATUS | CATATAN | LOKASI | FOTO
