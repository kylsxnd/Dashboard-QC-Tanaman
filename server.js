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

        // TARIK 2 FILE SEKALIGUS SECARA PARALEL BIA CEPAT
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
