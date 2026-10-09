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
app.use(express.static(path.join(__dirname, 'public')));

let cachedData = {}; 
let lastFetchTime = 0;
let changeLogs = [];
let fetchPromise = null; 

let emailDatabase = [
    { id: 1, sender: "Direktorat Operasi", subject: "Surat Pengawasan Panen PT Sumber Sawindo Kencana", snippet: "Kepada Yth. CRO III Regional Riau 2 & 3...", time: "22:33", label: "BELUM DI CRO", unread: false },
    { id: 2, sender: "Direktorat Operasi", subject: "Surat Pengawasan Panen PT Mutiara Naga Indonesia", snippet: "Kepada Yth. CRO III Regional Riau 2 & 3...", time: "22:35", label: "SUDAH DI CRO", unread: false }
];

function fetchGoogleSheets() {
    if (fetchPromise) return fetchPromise; 
    
    fetchPromise = (async () => {
        try {
            console.log("Mulai menarik data dari Google Sheets (Mode Jalur Cepat CSV)...");
            let allSheets = {};

            // 1. TARIK DATA MASTER (SHEET 1) LANGSUNG PAKAI CSV
            // Bypass format XLSX total biar ga kena limit Vercel 10 detik gara-gara ada grafik
            const criticalSheets = ['REKAPITULASI', 'REKAPITULASI PEMBATALAN', 'API_PIVOT'];
            
            await Promise.all(criticalSheets.map(async (sheetName) => {
                try {
                    const csvUrl = `https://docs.google.com/spreadsheets/d/${SHEET_ID_1}/gviz/tq?tqx=out:csv&sheet=${encodeURIComponent(sheetName)}`;
                    const resCsv = await axios.get(csvUrl, { responseType: 'arraybuffer' });
                    const wbCsv = xlsx.read(resCsv.data, { type: 'buffer' });
                    const sheet = wbCsv.Sheets[wbCsv.SheetNames[0]];
                    allSheets[sheetName] = xlsx.utils.sheet_to_json(sheet, { defval: "-" });
                    console.log(`Sukses narik: ${sheetName}`);
                } catch (e) {
                    console.error(`Gagal narik CSV: ${sheetName}`, e.message);
                }
            }));

            // 2. TARIK DATA EVALUASI KINERJA (SHEET 2) NORMAL (Karena ini aman ga ada grafik berat)
            try {
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
                    allSheets[sheetName] = xlsx.utils.sheet_to_json(sheet, options);
                });
                console.log("Sukses narik Sheet 2 (Evaluasi)");
            } catch (err2) {
                console.error("Gagal menarik Sheet 2:", err2.message);
            }

            cachedData = allSheets;
            lastFetchTime = Date.now();
            console.log("Semua data berhasil di-cache! Total Sheet:", Object.keys(cachedData).length);
            return cachedData;
            
        } catch (error) {
            console.error("Sistem gagal total:", error.message);
            return cachedData; 
        } finally {
            fetchPromise = null; 
        }
    })();

    return fetchPromise;
}

fetchGoogleSheets().catch(console.error);

app.get('/', (req, res) => res.redirect('/login'));
app.get('/login', (req, res) => res.render('login'));

app.get('/dashboard', async (req, res) => {
    if (Object.keys(cachedData).length === 0) {
        console.log("Data kosong, menunggu fetch...");
        await fetchGoogleSheets();
    }
    res.render('index', { 
        sheetData: JSON.stringify(cachedData), 
        logsData: JSON.stringify(changeLogs),
        emailsData: JSON.stringify(emailDatabase)
    });
});

app.get('/api/check-updates', async (req, res) => {
    if (Object.keys(cachedData).length === 0) {
        await fetchGoogleSheets();
    }
    res.json({ logs: changeLogs, rawData: cachedData });
    
    if (Date.now() - lastFetchTime > 10000) {
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
