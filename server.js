const express = require('express');
const axios = require('axios');
const xlsx = require('xlsx');
const path = require('path'); // Wajib ada untuk Vercel

const app = express();
const PORT = process.env.PORT || 3000; // Penting untuk Vercel

// ==========================================
// MASUKKAN 2 ID GOOGLE SHEETS DI SINI
// ==========================================
const SHEET_ID_1 = '1JNPHD-Vg1YjN84z0aVHFiNXhLLQ86ARNCb2MMo5s5iY'; // Master Data (Sheet 1)
const SHEET_ID_2 = '1bbVifsoYTxQFc0t80_tGEU1KyUtFlGsD5kH6opjshQg'; // Evaluasi Kinerja (Sheet 2 - Hanya Lihat)

// Wajib biar Vercel nggak nyasar nyari EJS
app.set('views', path.join(__dirname, 'views'));
app.set('view engine', 'ejs');
app.use(express.json());

let cachedData = null;
let lastFetchTime = 0;
let changeLogs = [];

let emailDatabase = [
    { id: 1, sender: "Direktorat Operasi", subject: "Surat Pengawasan Panen PT Sumber Sawindo Kencana", snippet: "Kepada Yth. CRO III Regional Riau 2 & 3...", time: "22:33", label: "BELUM DI CRO", unread: false },
    { id: 2, sender: "Direktorat Operasi", subject: "Surat Pengawasan Panen PT Mutiara Naga Indonesia", snippet: "Kepada Yth. CRO III Regional Riau 2 & 3...", time: "22:35", label: "SUDAH DI CRO", unread: false },
    { id: 3, sender: "Jaden Fergil Simatu...", subject: "Review dan Pengajuan Vendor Regional I Sumut - Aceh", snippet: "Selamat pagi, berikut kami kirim data vendor PT Wira Tiga Putra...", time: "07:59", label: "BELUM DI CRO", unread: false },
    { id: 4, sender: "Jaden Fergil Simatu...", subject: "Review dan Pengajuan Vendor Regional I Sumut - Aceh", snippet: "Selamat pagi, berikut kami kirim data vendor PT Perkebunan Sungai Wang...", time: "08:02", label: "SUDAH DI CRO", unread: false }
];

async function fetchGoogleSheets() {
    if (cachedData && (Date.now() - lastFetchTime < 10000)) return cachedData;
    
    let allSheets = {};
    
    try {
        // ==========================================
        // 1. TARIK DATA DARI SHEET 1 (MASTER DATA)
        // ==========================================
        const url1 = `https://docs.google.com/spreadsheets/d/${SHEET_ID_1}/export?format=xlsx`;
        const response1 = await axios.get(url1, { responseType: 'arraybuffer' });
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

        // ==========================================
        // 2. TARIK DATA DARI SHEET 2 (EVALUASI KINERJA)
        // ==========================================
        const url2 = `https://docs.google.com/spreadsheets/d/${SHEET_ID_2}/export?format=xlsx`;
        const response2 = await axios.get(url2, { responseType: 'arraybuffer' });
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
            // Gabungkan data dari Sheet 2 ke object allSheets
            allSheets[sheetName] = xlsx.utils.sheet_to_json(sheet, options);
        });

        cachedData = allSheets;
        lastFetchTime = Date.now();
        return allSheets;
        
    } catch (error) {
        console.error("Gagal menarik data:", error.message);
        // Tetap kembalikan data lama jika gagal narik, biar web nggak langsung crash
        return cachedData || null; 
    }
}

// ==========================================
// JALUR (ROUTES) HALAMAN WEB
// ==========================================
app.get('/', (req, res) => res.redirect('/login'));
app.get('/login', (req, res) => res.render('login'));

// Halaman 1: Dashboard Utama
app.get('/dashboard', async (req, res) => {
    const data = await fetchGoogleSheets();
    if (!data) return res.send("Gagal memuat data dari Google Sheets.");
    res.render('index', { 
        sheetData: JSON.stringify(data), 
        logsData: JSON.stringify(changeLogs),
        emailsData: JSON.stringify(emailDatabase)
    });
});

// Halaman 2: Evaluasi Kinerja (File evaluasi.ejs harus ada di folder views)
app.get('/evaluasi', async (req, res) => {
    const data = await fetchGoogleSheets();
    if (!data) return res.send("Gagal memuat data dari Google Sheets.");
    res.render('evaluasi', { 
        sheetData: JSON.stringify(data), 
        logsData: JSON.stringify(changeLogs),
        emailsData: JSON.stringify(emailDatabase)
    });
});

// ==========================================
// JALUR API UNTUK AUTO-UPDATE BACKGROUND
// ==========================================
app.get('/api/check-updates', async (req, res) => {
    const data = await fetchGoogleSheets();
    res.json({ logs: changeLogs, rawData: data });
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

// Penting: Export app untuk Vercel Serverless Function
module.exports = app;

// Hanya listen jika dijalankan di local, Vercel nggak butuh app.listen
if (require.main === module) {
    app.listen(PORT, () => {
        console.log(`🚀 Sistem aktif di http://localhost:${PORT}`);
    });
}
