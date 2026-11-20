const multer = require('multer');
const xlsx = require('xlsx');
const upload = multer({ storage: multer.memoryStorage() });
const express = require('express');
const bodyParser = require('body-parser');
const path = require('path');
const fs = require('fs');
const db = require('./database');
const app = express();
const PORT = process.env.PORT || 3000;

app.set('view engine', 'ejs');
app.use(bodyParser.urlencoded({ extended: true }));
app.use(bodyParser.json());
app.use(express.static('public'));

function resolveBpsIconPath() {
    const primaryPath = path.join(__dirname, 'public', 'img', 'bps.png');
    const fallbackPath = path.join(__dirname, 'public', 'img', 'bps1.png');
    if (fs.existsSync(primaryPath)) return primaryPath;
    if (fs.existsSync(fallbackPath)) return fallbackPath;
    return null;
}

// Layani /favicon.ico dan /favicon.png langsung agar browser tidak menampilkan ikon bumi default
app.get(['/favicon.ico', '/favicon.png'], (req, res) => {
    const iconFile = resolveBpsIconPath();
    if (iconFile) {
        res.setHeader('Cache-Control', 'public, max-age=86400');
        return res.sendFile(iconFile);
    }
    res.status(204).end();
});

function injectFavicon(html) {
    if (typeof html !== 'string' || !html.includes('</head>')) return html;

    const iconHref = fs.existsSync(path.join(__dirname, 'public', 'img', 'bps.png'))
        ? '/img/bps.png?v=3'
        : '/img/bps1.png?v=3';

    const faviconTags = [
        `<link rel="icon" type="image/png" sizes="32x32" href="${iconHref}">`,
        `<link rel="shortcut icon" type="image/png" href="${iconHref}">`,
        `<link rel="apple-touch-icon" href="${iconHref}">`
    ].join('\n');

    let updated = html.replace(/<link[^>]*rel=["'](?:shortcut\s+)?icon["'][^>]*>\s*/gi, '');
    updated = updated.replace('</head>', `${faviconTags}\n</head>`);

    // Pastikan judul tab menampilkan PROGRES.NG
    if (/<title>[\s\S]*?<\/title>/i.test(updated)) {
        updated = updated.replace(/<title>([\s\S]*?)<\/title>/i, (_, innerTitle) => {
            const cleanTitle = String(innerTitle || '').trim();
            if (/progress\.?ng/i.test(cleanTitle)) {
                return `<title>${cleanTitle}</title>`;
            }
            return `<title>PROGRES.NG - ${cleanTitle || 'BPS Kab. Bangka Tengah'}</title>`;
        });
    } else {
        updated = updated.replace('</head>', `<title>PROGRES.NG - BPS Kab. Bangka Tengah</title>\n</head>`);
    }

    return updated;
}

// Pasang favicon BPS otomatis untuk semua halaman HTML/EJS (baik res.render maupun res.send)
app.use((req, res, next) => {
    const originalSend = res.send.bind(res);
    res.send = (body) => {
        return originalSend(injectFavicon(body));
    };

    const originalRender = res.render.bind(res);
    res.render = (view, options, callback) => {
        if (typeof options === 'function') {
            callback = options;
            options = {};
        }
        return originalRender(view, options || {}, (err, html) => {
            if (err) {
                if (callback) return callback(err);
                return next(err);
            }
            const updatedHtml = injectFavicon(html);
            if (callback) return callback(null, updatedHtml);
            return res.send(updatedHtml);
        });
    };
    next();
});

let schemaEnsured = false;
async function ensureSchemaColumns() {
    if (schemaEnsured) return;
    try {
        await db.query(`ALTER TABLE activities ADD COLUMN IF NOT EXISTS target INTEGER DEFAULT 0`);
        await db.query(`ALTER TABLE reports ADD COLUMN IF NOT EXISTS target_achieved INTEGER DEFAULT 0`);
        schemaEnsured = true;
    } catch (err) {
        console.error('Gagal memastikan kolom schema:', err.message);
    }
}

function parseNonNegativeInt(val, fallback = 0) {
    if (val === undefined || val === null || String(val).trim() === '') return fallback;
    const cleaned = String(val).replace(/[^0-9]/g, '');
    if (!cleaned) return fallback;
    const num = parseInt(cleaned, 10);
    return Number.isNaN(num) ? fallback : num;
}

function parseMonthInput(rawMonth) {
    if (rawMonth === undefined || rawMonth === null) return null;
    const str = String(rawMonth).trim().toLowerCase();
    if (!str || str === '-') return null;
    const num = parseInt(str, 10);
    if (!Number.isNaN(num) && num >= 1 && num <= 12) return num;
    const monthMap = {
        'jan': 1, 'januari': 1,
        'feb': 2, 'februari': 2,
        'mar': 3, 'maret': 3,
        'apr': 4, 'april': 4,
        'mei': 5, 'may': 5,
        'jun': 6, 'juni': 6,
        'jul': 7, 'juli': 7,
        'agu': 8, 'ags': 8, 'agustus': 8,
        'sep': 9, 'september': 9,
        'okt': 10, 'oct': 10, 'oktober': 10,
        'nov': 11, 'nopember': 11, 'november': 11,
        'des': 12, 'dec': 12, 'desember': 12
    };
    return monthMap[str] || null;
}

function calculateProgressData(activities, reports) {
    const totalActivities = activities.length;
    if (totalActivities === 0) return { totalActivities: 0, completedActivities: 0, overallProgress: 0, monthlyTrend: Array(12).fill(0), teamProgress: [], statusCount: {} };

    const teams = ['Sosial', 'Distribusi', 'Produksi', 'Neraca', 'IPDS', 'Sektoral', 'Umum', 'Reformasi Birokrasi & PEKPPP'];
    let statusCount = { 'Perencanaan': 0, 'Persiapan': 0, 'Pelaksanaan': 0, 'Selesai': 0, 'Belum Lapor': 0 };
    let cumulativeByMonth = Array(12).fill(0);

    for (let m = 1; m <= 12; m++) {
        let monthSum = 0;
        activities.forEach(act => {
            const actReports = reports.filter(r => r.activity_id === act.id && r.month <= m);
            if (actReports.length > 0) {
                actReports.sort((a, b) => b.month - a.month);
                const latestReport = actReports[0];
                monthSum += (Number(latestReport.progress_percentage) || 0);

                if (m === 11) { // Bulan November
                    if (latestReport.stage_status) statusCount[latestReport.stage_status] = (statusCount[latestReport.stage_status] || 0) + 1;
                }
            }
        });
        cumulativeByMonth[m - 1] = Number((monthSum / totalActivities).toFixed(2));
    }

    const completedActivities = activities.filter(a => {
        const actReports = reports.filter(r => r.activity_id === a.id);
        return actReports.some(r => Number(r.progress_percentage) >= 100 || r.is_completed === 1);
    }).length;

    const overallProgress = cumulativeByMonth[10];

    const teamProgress = teams.map(team => {
        const tActs = activities.filter(a => a.team && a.team.toLowerCase() === team.toLowerCase());
        const tTotal = tActs.length;
        if (tTotal === 0) return { team, progress: 0, total: 0 };

        let tSum = 0;
        tActs.forEach(act => {
            const actReports = reports.filter(r => r.activity_id === act.id && r.month <= 11);
            if (actReports.length > 0) {
                actReports.sort((a, b) => b.month - a.month);
                tSum += (Number(actReports[0].progress_percentage) || 0);
            }
        });
        return { team, progress: Number((tSum / tTotal).toFixed(1)), total: tTotal };
    });

    return { totalActivities, completedActivities, overallProgress, monthlyTrend: cumulativeByMonth, teamProgress, statusCount };
}

// 1. Dashboard
app.get('/', async (req, res) => {
    try {
        await ensureSchemaColumns();
        const selectedYear = req.query.year || '2026';
        const actRes = await db.query(`SELECT * FROM activities`);
        const repRes = await db.query(`SELECT * FROM reports`);
        const stats = calculateProgressData(actRes.rows || [], repRes.rows || []);
        res.render('dashboard', { ...stats, selectedYear, activePage: 'dashboard' });
    } catch (err) {
        console.error('Error dashboard:', err);
        res.status(500).send('Database Error: ' + err.message);
    }
});

// 2. Entri Kegiatan (Dengan pesan status import & hapus)
app.get('/entry', async (req, res) => {
    try {
        await ensureSchemaColumns();
        const selectedYear = req.query.year || '2026';
        const search = req.query.search || '';
        const msg = req.query.msg || '';
        const query = `SELECT * FROM activities WHERE name ILIKE $1 OR ro_code ILIKE $2 OR team ILIKE $3 OR sub_team ILIKE $4 ORDER BY team, id DESC`;
        const param = `%${search}%`;
        const result = await db.query(query, [param, param, param, param]);
        res.render('entry', { activities: result.rows || [], search, msg, selectedYear, activePage: 'entry' });
    } catch (err) {
        console.error('Error entry:', err);
        res.status(500).send('Database Error: ' + err.message);
    }
});

// Tambah Kegiatan Manual -> Otomatis terintegrasi ke Pelaporan & Dashboard
app.post('/entry/add', async (req, res) => {
    try {
        await ensureSchemaColumns();
        const { team, sub_team, ro_code, name, jenis, target } = req.body;
        const targetVal = parseNonNegativeInt(target, 0);
        await db.query(
            `INSERT INTO activities (team, sub_team, ro_code, name, jenis, target) VALUES ($1, $2, $3, $4, $5, $6)`,
            [team, sub_team, ro_code, name, jenis, targetVal]
        );
        res.redirect('/entry?msg=Kegiatan+baru+berhasil+ditambahkan');
    } catch (err) {
        res.redirect(`/entry?msg=Gagal+menambah:+${encodeURIComponent(err.message)}`);
    }
});

// Edit Kegiatan
app.post('/entry/edit', async (req, res) => {
    try {
        await ensureSchemaColumns();
        const { id, team, sub_team, ro_code, name, jenis, target } = req.body;
        const targetVal = parseNonNegativeInt(target, 0);
        await db.query(
            `UPDATE activities SET team = $1, sub_team = $2, ro_code = $3, name = $4, jenis = $5, target = $6 WHERE id = $7`,
            [team, sub_team, ro_code, name, jenis, targetVal, id]
        );
        res.redirect('/entry?msg=' + encodeURIComponent('Kegiatan berhasil diperbarui.'));
    } catch (err) {
        console.error(err);
        res.redirect('/entry?msg=' + encodeURIComponent('Gagal memperbarui kegiatan.'));
    }
});

// Fitur Import Excel Kegiatan (Sesuai struktur kolom: No | Kode RO | Tim | Sub Tim | Nama Kegiatan | Jenis | Target)
app.post('/entry/import', upload.single('excel_file'), async (req, res) => {
    try {
        await ensureSchemaColumns();
        if (!req.file) {
            return res.redirect('/entry?msg=File+belum+dipilih');
        }

        const defaultTarget = parseNonNegativeInt(req.body.default_target, 0);
        const workbook = xlsx.read(req.file.buffer, { type: 'buffer' });
        const extractedActivities = [];

        workbook.SheetNames.forEach(sheetName => {
            const sheet = workbook.Sheets[sheetName];
            const rows = xlsx.utils.sheet_to_json(sheet, { header: 1, defval: '' });

            let targetColIndex = -1;
            for (const r of rows) {
                if (!Array.isArray(r)) continue;
                const foundIdx = r.findIndex(cell => String(cell || '').trim().toLowerCase() === 'target');
                if (foundIdx !== -1) {
                    targetColIndex = foundIdx;
                    break;
                }
            }

            rows.forEach((row) => {
                const colB = String(row[1] || '').trim();
                const colC = String(row[2] || '').trim();
                const colD = String(row[3] || '').trim();
                const colE = String(row[4] || '').trim();
                const colF = String(row[5] || '').trim();
                const rawTargetCell = targetColIndex !== -1 ? row[targetColIndex] : row[6];

                if (!colB || !colC || !colE) return;
                if (['kode ro', 'ro', 'tim', 'no'].includes(colB.toLowerCase()) || ['tim', 'tim kerja'].includes(colC.toLowerCase())) {
                    return;
                }

                extractedActivities.push({
                    ro_code: colB,
                    team: colC,
                    sub_team: colD || '-',
                    name: colE,
                    jenis: colF || (colE.toLowerCase().includes('publikasi') ? 'Publikasi' : (colE.toLowerCase().includes('laporan') ? 'Laporan' : 'Survei/Kegiatan')),
                    target: parseNonNegativeInt(rawTargetCell, defaultTarget)
                });
            });
        });

        if (extractedActivities.length === 0) {
            return res.redirect('/entry?msg=Tidak+ada+baris+kegiatan+valid+yang+ditemukan+di+Excel');
        }

        let inserted = 0;
        let lastError = '';

        for (const item of extractedActivities) {
            try {
                const existing = await db.query(
                    `SELECT id FROM activities WHERE ro_code = $1 AND name = $2`,
                    [item.ro_code, item.name]
                );

                if (existing.rows.length === 0) {
                    await db.query(
                        `INSERT INTO activities (team, sub_team, ro_code, name, jenis, target) VALUES ($1, $2, $3, $4, $5, $6)`,
                        [item.team, item.sub_team, item.ro_code, item.name, item.jenis, item.target]
                    );
                    inserted++;
                } else {
                    await db.query(
                        `UPDATE activities SET target = $1 WHERE id = $2`,
                        [item.target, existing.rows[0].id]
                    );
                }
            } catch (insertErr) {
                console.error('PostgreSQL Insert Error:', insertErr.message);
                lastError = insertErr.message;
            }
        }

        if (inserted === 0 && lastError) {
            res.redirect(`/entry?msg=Gagal+menyimpan+ke+database:+${encodeURIComponent(lastError)}`);
        } else if (inserted < extractedActivities.length && inserted > 0) {
            res.redirect(`/entry?msg=Berhasil+mengimpor+${inserted}+kegiatan+baru+(${extractedActivities.length - inserted}+data+sudah+ada+sebelumnya)`);
        } else {
            res.redirect(`/entry?msg=Berhasil+mengimpor+${inserted}+kegiatan+baru+dari+${extractedActivities.length}+baris+Excel`);
        }
    } catch (err) {
        console.error('Error import Excel:', err);
        res.redirect(`/entry?msg=Error:+${encodeURIComponent(err.message)}`);
    }
});

// Hapus Semua Kegiatan & Laporan (Dengan Konfirmasi Password "0000")
app.post('/entry/delete-all', async (req, res) => {
    try {
        const { password } = req.body;
        if (password !== '0000') {
            return res.redirect('/entry?msg=Gagal+menghapus+semua:+Password+konfirmasi+salah!');
        }

        await db.query(`TRUNCATE TABLE reports, activities RESTART IDENTITY CASCADE`);
        res.redirect('/entry?msg=Seluruh+data+kegiatan+dan+laporan+berhasil+dihapus');
    } catch (err) {
        res.redirect('/entry?msg=Gagal+menghapus+data:+' + encodeURIComponent(err.message));
    }
});

app.post('/entry/delete/:id', async (req, res) => {
    try {
        const id = req.params.id;
        await db.query(`DELETE FROM reports WHERE activity_id = $1`, [id]);
        await db.query(`DELETE FROM activities WHERE id = $1`, [id]);
        res.redirect('/entry');
    } catch (err) {
        console.error('Error delete:', err);
        res.redirect('/entry');
    }
});

// 3. Pelaporan Bulanan (Otomatis mengambil data dari tabel activities pusat)
app.get('/pelaporan', async (req, res) => {
    try {
        await ensureSchemaColumns();
        const selectedYear = req.query.year || '2026';
        const actRes = await db.query(`SELECT * FROM activities ORDER BY team`);
        const repRes = await db.query(`SELECT * FROM reports`);
        const teams = ['Sosial', 'Distribusi', 'Produksi', 'Neraca', 'IPDS', 'Sektoral', 'Umum', 'Reformasi Birokrasi & PEKPPP'];
        res.render('pelaporan', { teams, activities: actRes.rows || [], reports: repRes.rows || [], selectedYear, activePage: 'pelaporan' });
    } catch (err) {
        console.error('Error pelaporan:', err);
        res.status(500).send('Database Error: ' + err.message);
    }
});

// Import Pelaporan Kinerja dari Excel / CSV (Sesuai format Template & Export)
app.post('/pelaporan/import', upload.single('pelaporan_file'), async (req, res) => {
    try {
        await ensureSchemaColumns();
        if (!req.file) {
            return res.redirect('/pelaporan?error=' + encodeURIComponent('File import pelaporan belum dipilih.'));
        }

        const workbook = xlsx.read(req.file.buffer, { type: 'buffer' });
        const actRes = await db.query(`SELECT id, name, team FROM activities`);
        const allActivities = actRes.rows || [];

        let importedCount = 0;

        for (const sheetName of workbook.SheetNames) {
            const sheet = workbook.Sheets[sheetName];
            const rows = xlsx.utils.sheet_to_json(sheet, { defval: '' });

            for (const row of rows) {
                const actName = String(row['Nama Kegiatan / Survei'] || row['Nama Kegiatan'] || row['name'] || '').trim();
                const teamName = String(row['Tim'] || row['team'] || '').trim();
                const mNum = parseMonthInput(row['Bulan'] || row['month']);

                if (!actName || !mNum) continue;

                const matchedAct = allActivities.find(a => {
                    const sameName = String(a.name || '').trim().toLowerCase() === actName.toLowerCase();
                    if (!teamName) return sameName;
                    return sameName && String(a.team || '').trim().toLowerCase() === teamName.toLowerCase();
                }) || allActivities.find(a => String(a.name || '').trim().toLowerCase() === actName.toLowerCase());

                if (!matchedAct) continue;

                const rawStageType = String(row['Jenis Tahapan'] || row['stage_type'] || 'Persiapan').trim();
                const stageType = rawStageType || 'Persiapan';
                const rawStageStatus = String(row['Status Pelaksanaan'] || row['stage_status'] || 'Pelaksanaan').trim();
                const progressVal = parseFloat(row['Besaran Progress (%)'] !== undefined && row['Besaran Progress (%)'] !== '' ? row['Besaran Progress (%)'] : (row['progress_percentage'] || 0));
                const safeProgress = Number.isNaN(progressVal) ? 0 : Math.min(100, Math.max(0, progressVal));
                const isCompletedFlag = safeProgress >= 100 ? 1 : 0;
                const finalStatus = safeProgress >= 100 ? 'Selesai' : (rawStageStatus || 'Pelaksanaan');

                const rawTarget = row['Target Tercapai (Publikasi / Laporan / Satuan)'] !== undefined
                    ? row['Target Tercapai (Publikasi / Laporan / Satuan)']
                    : (row['target_achieved'] || 0);
                const targetAchieved = parseNonNegativeInt(rawTarget, 0);

                const description = String(row['Kendala'] || row['Uraian / Catatan Kegiatan'] || row['description'] || '').trim();
                const problem = String(row['Solusi'] || row['Permasalahan'] || row['problem'] || '').trim();
                const followUp = String(row['Tindak Lanjut'] || row['follow_up'] || '').trim();
                const proofLink = String(row['Tautan Bukti Dukung'] || row['proof_link'] || '').trim();

                const existingRes = await db.query(
                    `SELECT id FROM reports WHERE activity_id = $1 AND month = $2 AND stage_type = $3`,
                    [matchedAct.id, mNum, stageType]
                );

                if (existingRes.rows && existingRes.rows.length > 0) {
                    await db.query(
                        `UPDATE reports SET stage_status = $1, progress_percentage = $2, target_achieved = $3, description = $4, problem = $5, follow_up = $6, proof_link = $7, is_completed = $8 WHERE id = $9`,
                        [finalStatus, safeProgress, targetAchieved, description, problem, followUp, proofLink, isCompletedFlag, existingRes.rows[0].id]
                    );
                } else {
                    await db.query(
                        `INSERT INTO reports (activity_id, month, stage_type, stage_status, progress_percentage, target_achieved, description, problem, follow_up, proof_link, is_completed) VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11)`,
                        [matchedAct.id, mNum, stageType, finalStatus, safeProgress, targetAchieved, description, problem, followUp, proofLink, isCompletedFlag]
                    );
                }
                importedCount++;
            }
        }

        if (importedCount === 0) {
            return res.redirect('/pelaporan?error=' + encodeURIComponent('Tidak ada baris laporan valid yang cocok dengan kegiatan di database.'));
        }

        return res.redirect('/pelaporan?success=1&msg=' + encodeURIComponent(`Berhasil mengimpor ${importedCount} baris data pelaporan kinerja.`));
    } catch (err) {
        console.error('Error import pelaporan:', err);
        return res.redirect('/pelaporan?error=' + encodeURIComponent('Gagal mengimpor pelaporan: ' + err.message));
    }
});

// Simpan Laporan dengan Validasi & Status Selesai Otomatis Jika Progress 100%
app.post('/pelaporan/save', async (req, res) => {
    try {
        await ensureSchemaColumns();
        const { activity_id, month, stage_type, stage_status, progress_percentage, target_achieved, description, problem, follow_up, proof_link } = req.body;
        const newProgress = parseFloat(progress_percentage);
        const mNum = parseInt(month);
        const targetAchievedVal = parseNonNegativeInt(target_achieved, 0);
        const isCompletedFlag = newProgress >= 100 ? 1 : 0;
        const finalStatus = newProgress >= 100 ? 'Selesai' : stage_status;

        if (isNaN(newProgress) || newProgress < 0 || newProgress > 100) {
            return res.redirect('/pelaporan?error=' + encodeURIComponent('Besaran progress harus berupa angka antara 0% sampai 100%.'));
        }

        const prevRes = await db.query(
            `SELECT month, progress_percentage FROM reports WHERE activity_id = $1 AND month < $2 ORDER BY month DESC`,
            [activity_id, mNum]
        );

        if (prevRes.rows && prevRes.rows.length > 0) {
            const prevProgress = Number(prevRes.rows[0].progress_percentage) || 0;
            const prevMonth = prevRes.rows[0].month;
            if (newProgress < prevProgress) {
                const reason = `Besaran progress yang diinput (${newProgress}%) tidak boleh lebih kecil dari bulan sebelumnya (Bulan ke-${prevMonth}: ${prevProgress}%).`;
                return res.redirect('/pelaporan?error=' + encodeURIComponent(reason));
            }
        }

        const existingRes = await db.query(
            `SELECT id FROM reports WHERE activity_id = $1 AND month = $2 AND stage_type = $3`,
            [activity_id, mNum, stage_type]
        );

        if (existingRes.rows && existingRes.rows.length > 0) {
            const rowId = existingRes.rows[0].id;
            await db.query(
                `UPDATE reports SET stage_status = $1, progress_percentage = $2, target_achieved = $3, description = $4, problem = $5, follow_up = $6, proof_link = $7, is_completed = $8 WHERE id = $9`,
                [finalStatus, newProgress, targetAchievedVal, description, problem, follow_up, proof_link, isCompletedFlag, rowId]
            );
            res.redirect('/pelaporan?success=1');
        } else {
            await db.query(
                `INSERT INTO reports (activity_id, month, stage_type, stage_status, progress_percentage, target_achieved, description, problem, follow_up, proof_link, is_completed) VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11)`,
                [activity_id, mNum, stage_type, finalStatus, newProgress, targetAchievedVal, description, problem, follow_up, proof_link, isCompletedFlag]
            );
            res.redirect('/pelaporan?success=1');
        }
    } catch (err) {
        return res.redirect('/pelaporan?error=' + encodeURIComponent('Gagal menyimpan laporan ke database: ' + err.message));
    }
});

app.get('/api/report', async (req, res) => {
    try {
        await ensureSchemaColumns();
        const { activity_id, month, stage_type } = req.query;
        const result = await db.query(
            `SELECT * FROM reports WHERE activity_id = $1 AND month = $2 AND stage_type = $3`,
            [activity_id, month, stage_type]
        );
        res.json(result.rows[0] || {});
    } catch (err) {
        res.json({});
    }
});

if (require.main === module) {
    app.listen(PORT, () => {
        console.log(`PROGRES.NG Server running on http://localhost:${PORT}`);
    });
}

module.exports = app;