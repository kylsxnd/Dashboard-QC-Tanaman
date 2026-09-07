const express = require('express');
const axios = require('axios');
const xlsx = require('xlsx');

const path = require('path');
app.set('views', path.join(__dirname, 'views'));

const app = express();
const PORT = 3000;
const SHEET_ID = '1JNPHD-Vg1YjN84z0aVHFiNXhLLQ86ARNCb2MMo5s5iY';

app.set('view engine', 'ejs');

let cachedData = null;
let lastFetchTime = 0;
let changeLogs = [];

// ==========================================
// SIMULASI DATABASE EMAIL DARI GMAIL
// ==========================================
// Ini data dummy berdasarkan screenshot Gmail lu
let emailDatabase = [
    { id: 1, sender: "Direktorat Operasi", subject: "Surat Pengawasan Panen PT Sumber Sawindo Kencana", snippet: "Kepada Yth. CRO III Regional Riau 2 & 3 PT Agrinas Palma Nusantara...", time: "22:33", unread: false },
    { id: 2, sender: "Direktorat Operasi", subject: "Surat Pengawasan Panen PT Mutiara Naga Indonesia", snippet: "Kepada Yth. CRO III Regional Riau 2 & 3 PT Agrinas Palma Nusantara...", time: "22:35", unread: false },
    { id: 3, sender: "Jaden Fergil Simatu...", subject: "Review dan Pengajuan Vendor Regional I Sumut - Aceh", snippet: "Selamat pagi, berikut kami kirim data vendor PT Wira Tiga Putra untuk kebun...", time: "07:59", unread: false },
    { id: 4, sender: "Jaden Fergil Simatu...", subject: "Review dan Pengajuan Vendor Regional I Sumut - Aceh", snippet: "Selamat pagi, berikut kami kirim data vendor PT Perkebunan Sungai Wang...", time: "08:02", unread: false }
];

// Buat ngetes notifikasi pop-up, kita bikin server seolah-olah nerima email baru setiap 15 detik
setTimeout(() => {
    emailDatabase.unshift({ 
        id: 5, 
        sender: "Jaden Fergil Simatu...", 
        subject: "SUDAH DI CRO Review dan Pengajuan Vendor", 
        snippet: "Selamat pagi, berikut kami kirim file Vendor PT Rimo Kapas Raya untuk kebun...", 
        time: "08:05", 
        unread: true // Ini bakal bikin notif merah nyala di Sidebar Kiri
    });
    console.log("📥 [SIMULASI] Email baru masuk dari Jaden Fergil!");
}, 15000);


async function fetchGoogleSheets() {
    if (cachedData && (Date.now() - lastFetchTime < 10000)) return cachedData;
    try {
        console.log("Sedot data dari Google Sheets (Auto-Sync)...");
        const url = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/export?format=xlsx`;
        const response = await axios.get(url, { responseType: 'arraybuffer' });
        const workbook = xlsx.read(response.data, { type: 'buffer' });
        
        let allSheets = {};
        workbook.SheetNames.forEach(sheetName => {
            let options = { defval: "-" };
            if (sheetName === 'Pembayaran Vendor' || sheetName === 'Pivot Table') options.range = 1;
            
            const sheet = workbook.Sheets[sheetName];
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

        // ALGORITMA DETEKTIF PERUBAHAN
        if (cachedData) {
            let adaPerubahan = false;
            let waktuSekarang = new Date().toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit', second: '2-digit' });

            for (let sheet of Object.keys(allSheets)) {
                if (!cachedData[sheet]) continue;
                let newData = allSheets[sheet];
                let oldData = cachedData[sheet];

                for (let i = 0; i < newData.length; i++) {
                    if (!oldData[i]) {
                        changeLogs.unshift({ waktu: waktuSekarang, sheet: sheet, baris: i + 2, kolom: 'DATA BARU', lama: 'Kosong', baru: 'Baris Baru Ditambahkan' });
                        adaPerubahan = true;
                        continue;
                    }
                    
                    for (let key of Object.keys(newData[i])) {
                        if (key.includes('__EMPTY')) continue; 
                        if (oldData[i][key] !== newData[i][key]) {
                            changeLogs.unshift({
                                waktu: waktuSekarang,
                                sheet: sheet,
                                baris: i + (sheet === 'Pembayaran Vendor' || sheet === 'Pivot Table' ? 3 : 2),
                                kolom: key,
                                lama: oldData[i][key],
                                baru: newData[i][key]
                            });
                            adaPerubahan = true;
                        }
                    }
                }
            }
            if (changeLogs.length > 30) changeLogs = changeLogs.slice(0, 30);
            if (adaPerubahan) console.log("🔔 Terdeteksi perubahan data! Notifikasi dikirim ke Web.");
        }

        cachedData = allSheets;
        lastFetchTime = Date.now();
        return allSheets;
    } catch (error) {
        console.error("Gagal menarik data:", error.message);
        return null;
    }
}

// ==========================================
// PENGATURAN HALAMAN (ROUTING)
// ==========================================

app.get('/', (req, res) => {
    res.redirect('/login');
});

app.get('/login', (req, res) => {
    res.render('login');
});

app.get('/dashboard', async (req, res) => {
    const data = await fetchGoogleSheets();
    if (!data) return res.send("Gagal memuat data dari Google Sheets.");
    // Kirim data sheets, logs, dan emails ke frontend
    res.render('index', { 
        sheetData: JSON.stringify(data), 
        logsData: JSON.stringify(changeLogs),
        emailsData: JSON.stringify(emailDatabase)
    });
});

// API Auto-Update G-Sheets
app.get('/api/check-updates', async (req, res) => {
    const data = await fetchGoogleSheets();
    res.json({ logs: changeLogs, rawData: data });
});

// API Khusus Ngecek Email Baru
app.get('/api/check-emails', (req, res) => {
    // Cari ada berapa email yang belum dibaca (unread)
    const unreadCount = emailDatabase.filter(email => email.unread).length;
    res.json({ emails: emailDatabase, unreadCount: unreadCount });
});

// API buat nanda-in kalau email udah dibaca pas masuk tab Inbox
app.post('/api/read-emails', (req, res) => {
    emailDatabase.forEach(email => email.unread = false);
    res.sendStatus(200);
});

app.listen(PORT, () => {
    console.log(`🚀 Sistem aktif di http://localhost:${PORT}`);
});