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

let emailDatabase = [
    { id: 1, sender: "Direktorat Operasi", subject: "Surat Pengawasan Panen PT Sumber Sawindo Kencana", snippet: "Kepada Yth. CRO III Regional Riau 2 & 3...", time: "22:33", label: "BELUM DI CRO", unread: false },
    { id: 2, sender: "Direktorat Operasi", subject: "Surat Pengawasan Panen PT Mutiara Naga Indonesia", snippet: "Kepada Yth. CRO III Regional Riau 2 & 3...", time: "22:35", label: "SUDAH DI CRO", unread: false },
    { id: 3, sender: "Jaden Fergil Simatu...", subject: "Review dan Pengajuan Vendor Regional I Sumut - Aceh", snippet: "Selamat pagi, berikut kami kirim data vendor PT Wira Tiga Putra...", time: "07:59", label: "BELUM DI CRO", unread: false },
    { id: 4, sender: "Jaden Fergil Simatu...", subject: "Review dan Pengajuan Vendor Regional I Sumut - Aceh", snippet: "Selamat pagi, berikut kami kirim data vendor PT Perkebunan Sungai Wang...", time: "08:02", label: "SUDAH DI CRO", unread: false }
];

// KUNCI ANTI-FREEZE VERCEL: Gunakan 1 Promise Global
let fetchPromise = null;

function fetchGoogleSheets() {
    // Kalau lagi proses narik data, tungguin aja biar gak dobel request
    if (fetchPromise) return fetchPromise; 
    
    fetchPromise = (async () => {
        try {
            const url1 = `https://docs.google.com/spreadsheets/d/${SHEET_ID_1}/export?format=xlsx`;
            const url2 = `https://docs.google.com/spreadsheets/d/${SHEET_ID_2}/export?format=xlsx`;

            // TARIK 2 FILE SEKALIGUS SECARA PARALEL BIAR CEPAT
            const [response1, response2] = await Promise.all([
                axios.get(url1, { responseType: 'arraybuffer' }),
                axios.get(url2, { responseType: 'arraybuffer' })
            ]);
            
            let allSheets = {};

            // Proses Sheet 1
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

            // Proses Sheet 2
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
                allSheets[sheetName] = xlsx.utils.sheet_to_json(sheet, options);
            });

            cachedData = allSheets;
            lastFetchTime = Date.now();
            return cachedData;
            
        } catch (error) {
            console.error("Gagal menarik data:", error.message);
            return cachedData; // Kembalikan data lama kalau error
        } finally {
            fetchPromise = null; // Bersihkan status agar bisa ditarik ulang nanti
        }
    })();

    return fetchPromise;
}

// Jalankan tarikan data pertama kali saat server nyala
fetchGoogleSheets().catch(console.error);

app.get('/', (req, res) => res.redirect('/login'));
app.get('/login', (req, res) => res.render('login'));

app.get('/dashboard', (req, res) => {
    // UI Pertama dirender. Kalau kosong gak masalah, karena Javascript di index.ejs bakal langsung nembak /api/check-updates
    res.render('index', { 
        sheetData: JSON.stringify(cachedData), 
        logsData: JSON.stringify(changeLogs),
        emailsData: JSON.stringify(emailDatabase)
    });
});

// ==========================================
// ANTI-LOADING API (STALE-WHILE-REVALIDATE YANG DIBENARKAN)
// ==========================================
app.get('/api/check-updates', async (req, res) => {
    // 1. JIKA CACHE KOSONG (Vercel baru bangun dari tidur):
    // KITA WAJIB TUNGGU (AWAIT) SAMPAI DATA SELESAI DITARIK!
    if (Object.keys(cachedData).length === 0) {
        await fetchGoogleSheets();
        return res.json({ logs: changeLogs, rawData: cachedData });
    }
    
    // 2. JIKA CACHE UDAH ADA:
    // Langsung kirim instan tanpa loading (0.1 detik)
    res.json({ logs: changeLogs, rawData: cachedData });
    
    // 3. AUTO-SYNC BACKGROUND: Tarik data terbaru untuk dikasih ke request selanjutnya
    if (Date.now() - lastFetchTime > 10000) {
        // Tarik diam-diam di background
        fetchGoogleSheets().catch(err => console.error(err)); 
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
