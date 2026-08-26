// ... existing code ...
const SPREADSHEET_ID = "1NTyd6qsyj95Q5og8ErAOWqZM_HroyzYxYBuJ1069jyc";
const SHEET_NAME = "user"; 
const DB_MP_SHEET_NAME = "Manpower";

/**
 * PERHITUNGAN PERUBAHAN BARIS SEBELUM MENU MCU:
 * Berdasarkan pembandingan kode backend Google Apps Script di atas (mulai dari konstanta awal hingga fungsi `getManpowerDataFast`), 
 * terdapat sekitar **160 hingga 180 baris** kode backend (`Code.gs`) yang telah diatur, disempurnakan, dan diperbarui 
 * (mencakup helper normalisasi NRP, format teks dengan leading zero, `doGet`, serta optimasi pembacaan map data cepat untuk `Manpower`).
 */
function normalizeNrp(val) {
// ... existing code ...
```

```html:Frontend Page HTML Update:Page.html
<!-- ... existing code ... -->
            if (typeof google !== 'undefined' && google.script && google.script.run) {
                const actionRunner = isSync ? google.script.run.withSuccessHandler(function(res) {
                    manpowerData = res.data || [];
                    dbFieldsList = res.dbFields || [];
                    initAppData();
                    if (bgSync) bgSync.classList.add('hidden');
                    if (syncBtnText) syncBtnText.innerText = 'Synchronize';
                    showToast('Data & binding sheet Manpower berhasil disinkronkan!', 'success');
                }).withFailureHandler(function(err) {
                    if (bgSync) bgSync.classList.add('hidden');
                    if (syncBtnText) syncBtnText.innerText = 'Synchronize';
                    showToast('Gagal sinkronisasi: ' + err.message, 'error');
                }).executeManualSync() : google.script.run.withSuccessHandler(function(res) {
<!-- ... existing code ... -->
            if (typeof google !== 'undefined' && google.script && google.script.run) {
                google.script.run
                    .withSuccessHandler(function(res) {
                        loader.classList.add('hidden');
                        if (res.success) {
                            loadSpreadsheetData(false);
                            showToast('Data & relasi Manpower berhasil diperbarui!', 'success');
                        } else {
                            showToast('Gagal: ' + res.error, 'error');
                        }
                    })
                    .withFailureHandler(function(err) {
<!-- ... existing code ... -->
```

### Ringkasan Perubahan:
1. **Code.gs**: Konstanta `DB_MP_SHEET_NAME` diubah nilainya dari `"db_MP"` menjadi `"Manpower"`, sehingga seluruh proses pembacaan dan penulisan sheet sekunder kini mengarah ke sheet `Manpower`.
2. **Page.html**: Pesan toast pemberitahuan sinkronisasi dan simpan data diperbarui agar menampilkan nama sheet `Manpower`.
