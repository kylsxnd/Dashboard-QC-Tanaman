const express = require('express');
const axios = require('axios');
const xlsx = require('xlsx');
const path = require('path'); 

const app = express();
const PORT = process.env.PORT || 3000; 

// ==========================================
// ID GOOGLE SHEETS
// ==========================================
const SHEET_ID_1 = '1JNPHD-Vg1YjN84z0aVHFiNXhLLQ86ARNCb2MMo5s5iY'; // Master Data
const SHEET_ID_2 = '1bbVifsoYTxQFc0t80_tGEU1KyUtFlGsD5kH6opjshQg'; // Evaluasi Kinerja

app.set('views', path.join(__dirname, 'views'));
app.set('view engine', 'ejs');
app.use(express.json());

let cachedData = {}; 
let lastFetchTime = 0;
let changeLogs = [];
let isFetching = false; 

let emailDatabase = [
    { id: 1, sender: "Direktorat Operasi", subject: "Surat Pengawasan Panen PT Sumber Sawindo Kencana", snippet: "Kepada Yth. CRO III Regional Riau 2 & 3...", time: "22:33", label: "BELUM DI CRO", unread: false },
    { id: 2, sender: "Direktorat Operasi", subject: "Surat Pengawasan Panen PT Mutiara Naga Indonesia", snippet: "Kepada Yth. CRO III Regional Riau 2 & 3...", time: "22:35", label: "SUDAH DI CRO", unread: false },
    { id: 3, sender: "Jaden Fergil Simatu...", subject: "Review dan Pengajuan Vendor Regional I Sumut - Aceh", snippet: "Selamat pagi, berikut kami kirim data vendor PT Wira Tiga Putra...", time: "07:59", label: "BELUM DI CRO", unread: false },
    { id: 4, sender: "Jaden Fergil Simatu...", subject: "Review dan Pengajuan Vendor Regional I Sumut - Aceh", snippet: "Selamat pagi, berikut kami kirim data vendor PT Perkebunan Sungai Wang...", time: "08:02", label: "SUDAH DI CRO", unread: false }
];

async function fetchGoogleSheets() {
    if (isFetching) return; // Biar server gak tabrakan narik datanya
    isFetching = true;
    
    try {
        const url1 = `https://docs.google.com/spreadsheets/d/${SHEET_ID_1}/export?format=xlsx`;
        const url2 = `https://docs.google.com/spreadsheets/d/${SHEET_ID_2}/export?format=xlsx`;

        // TARIK 2 FILE SEKALIGUS SECARA PARALEL BIAR CEPAT
        const [response1, response2] = await Promise.all([
            axios.get(url1, { responseType: 'arraybuffer' }),
            axios.get(url2, { responseType: 'arraybuffer' })
        ]);
        
        let allSheets = {};

        // Proses Sheet 1 (Master Data)
        const workbook1 = xlsx.read(response1.data, { type: 'buffer' });
        workbook1.SheetNames.forEach(sheetName => {
            let options = { defval: "-" };
            if (sheetName === 'Pembayaran Vendor' || sheetName === 'Pivot Table') options.range = 1;
            const sheet = workbook1.Sheets[sheetName];
            for (let cellAddress in sheet) {
                if (cellAddress.startsWith('!')) continue;
                const cell = sheet[cellAddress];
                if (cell && cell.l && cell.l.Target) {
                    cell.v = cell.v + "|||" + cell.l.Target;
                    if (cell.w) cell.w = cell.v;
                }
            }
            allSheets[sheetName] = xlsx.utils.sheet_to_json(sheet, options);
        });

        // Proses Sheet 2 (Evaluasi Kinerja)
        const workbook2 = xlsx.read(response2.data, { type: 'buffer' });
        workbook2.SheetNames.forEach(sheetName => {
            let options = { defval: "-" };
            const sheet = workbook2.Sheets[sheetName];
            for (let cellAddress in sheet) {
                if (cellAddress.startsWith('!')) continue;
                const cell = sheet[cellAddress];
                if (cell && cell.l && cell.l.Target) {
                    cell.v = cell.v + "|||" + cell.l.Target;
                    if (cell.w) cell.w = cell.v;
                }
            }
            
            // Simpan format aslinya
            allSheets[sheetName] = xlsx.utils.sheet_to_json(sheet, options);

            // ==========================================
            // LOGIKA KHUSUS: REKAP KEMITRAAN (EXTRACTION TAHAP 1-7)
            // ==========================================
            if (sheetName === 'Rekap Kemitraan') {
                let cleanData = [];
                // Baca sheet sebagai array murni biar header gabungan (merge cells) gak bikin error
                const rawArray = xlsx.utils.sheet_to_json(sheet, { header: 1, defval: "-" });
                
                if (rawArray.length > 2) {
                    // Cari index kolom "Tahap 1-7" dengan melihat baris ke-3 (index 2)
                    let headers = rawArray[2]; 
                    let regionalCount = 0;
                    let targetIdx = -1;
                    
                    for (let c = 0; c < headers.length; c++) {
                        if (String(headers[c]).trim().toLowerCase() === 'regional') {
                            regionalCount++;
                            // Kita incer "Regional" yang KEDUA, karena itu punyanya tabel Tahap 1-7
                            if (regionalCount === 2) { 
                                targetIdx = c;
                                break;
                            }
                        }
                    }

                    if (targetIdx !== -1) {
                        // Looping dari baris ke-4 (index 3) sampai data habis
                        for (let r = 3; r < rawArray.length; r++) {
                            let row = rawArray[r];
                            let cro = String(row[targetIdx] || "").trim();
                            let wilayah = String(row[targetIdx + 1] || "").trim();
                            
                            // Ambil angkanya saja
                            let jmlPT = String(row[targetIdx + 2] || "0").replace(/[^0-9]/g,'');
                            let sudah = String(row[targetIdx + 4] || "0").replace(/[^0-9]/g,'');
                            let belum = String(row[targetIdx + 5] || "0").replace(/[^0-9]/g,'');

                            // Buang baris kosong, header yang nyangkut, dan grand total
                            if (!cro || cro === "-" || cro.toLowerCase() === "regional") continue;
                            if (cro.toLowerCase().includes("total") || cro.toLowerCase().includes("summary") || cro.toLowerCase().includes("grand")) continue;

                            // Rapikan tulisan CRO biar seragam
                            if(cro.toLowerCase().startsWith('cro')) {
                                cro = cro.toUpperCase();
                            } else {
                                cro = "CRO " + cro.toUpperCase(); 
                            }

                            // Masukkan ke array bersih
                            cleanData.push({
                                "Regional": cro,
                                "Wilayah": wilayah,
                                "Jumlah PT": parseInt(jmlPT) || 0,
                                "Sudah Dikelola": parseInt(sudah) || 0,
                                "Belum Dikelola": parseInt(belum) || 0
                            });
                        }
                    }
                }
                // Simpan data bersih khusus ini dengan nama baru "Rekap_Kemitraan_Clean"
                allSheets["Rekap_Kemitraan_Clean"] = cleanData;
            }
        });

        cachedData = allSheets;
        lastFetchTime = Date.now();
        
    } catch (error) {
        console.error("Gagal menarik data:", error.message);
    } finally {
        isFetching = false;
    }
}

// Jalankan tarikan data pertama kali saat server nyala
fetchGoogleSheets();

app.get('/', (req, res) => res.redirect('/login'));
app.get('/login', (req, res) => res.render('login'));

app.get('/dashboard', (req, res) => {
    // INSTANT RENDER: Langsung kasih UI tanpa harus nunggu data (Anti Load Lama)
    res.render('index', { 
        sheetData: JSON.stringify(cachedData), 
        logsData: JSON.stringify(changeLogs),
        emailsData: JSON.stringify(emailDatabase)
    });
    // Kalo cache kosong (baru nyala), suruh server narik di background
    if (Object.keys(cachedData).length === 0) fetchGoogleSheets();
});

// ==========================================
// ANTI-LOADING API (STALE-WHILE-REVALIDATE)
// ==========================================
app.get('/api/check-updates', (req, res) => {
    // 1. Langsung kasih data yang ada (Instant Respon 0.1 detik)
    res.json({ logs: changeLogs, rawData: cachedData });
    
    // 2. Tapi secara diam-diam di background, dia ngecek data terbaru tiap 5 detik
    if (Date.now() - lastFetchTime > 5000) {
        fetchGoogleSheets(); 
    }
});

app.get('/api/check-emails', (req, res) => {
    const unreadCount = emailDatabase.filter(email => email.unread).length;
    res.json({ emails: emailDatabase, unreadCount: unreadCount });
});

app.post('/api/read-emails', (req, res) => {
    emailDatabase.forEach(email => email.unread = false);
    res.sendStatus(200);
});

app.post('/api/update-email-label', (req, res) => {
    const { id, label } = req.body;
    const targetEmail = emailDatabase.find(e => e.id === id);
    if (targetEmail) {
        targetEmail.label = label;
        return res.json({ success: true, emails: emailDatabase });
    }
    res.status(404).json({ success: false, message: "Email tidak ditemukan" });
});

module.exports = app;

if (require.main === module) {
    app.listen(PORT, () => {
        console.log(`🚀 Sistem aktif di http://localhost:${PORT}`);
    });
}
