const multer = require('multer');
const xlsx = require('xlsx');
const upload = multer({ storage: multer.memoryStorage() });
const express = require('express');
const bodyParser = require('body-parser');
const db = require('./database');
const app = express();
const PORT = process.env.PORT || 3000;

app.set('view engine', 'ejs');
app.use(bodyParser.urlencoded({ extended: true }));
app.use(bodyParser.json());
app.use(express.static('public'));

function calculateProgressData(activities, reports) {
    const totalActivities = activities.length;
    if (totalActivities === 0) return { totalActivities: 0, completedActivities: 0, overallProgress: 0, monthlyTrend: Array(12).fill(0), teamProgress: [], statusCount: {} };

    const teams = ['Sosial', 'Distribusi', 'Produksi', 'Neraca', 'IPDS', 'Sektoral', 'Umum'];
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
        const search = req.query.search || '';
        const msg = req.query.msg || '';
        const query = `SELECT * FROM activities WHERE name ILIKE $1 OR ro_code ILIKE $2 OR team ILIKE $3 OR sub_team ILIKE $4 ORDER BY team, id DESC`;
        const param = `%${search}%`;
        const result = await db.query(query, [param, param, param, param]);
        res.render('entry', { activities: result.rows || [], search, msg, activePage: 'entry' });
    } catch (err) {
        console.error('Error entry:', err);
        res.status(500).send('Database Error: ' + err.message);
    }
});

// Tambah Kegiatan Manual -> Otomatis terintegrasi ke Pelaporan & Dashboard
app.post('/entry/add', async (req, res) => {
    try {
        const { team, sub_team, ro_code, name, jenis } = req.body;
        await db.query(
            `INSERT INTO activities (team, sub_team, ro_code, name, jenis) VALUES ($1, $2, $3, $4, $5)`,
            [team, sub_team, ro_code, name, jenis]
        );
        res.redirect('/entry?msg=Kegiatan+baru+berhasil+ditambahkan');
    } catch (err) {
        res.redirect(`/entry?msg=Gagal+menambah:+${encodeURIComponent(err.message)}`);
    }
});

// Edit Kegiatan
app.post('/entry/edit', async (req, res) => {
    try {
        const { id, team, sub_team, ro_code, name, jenis } = req.body;
        await db.query(
            `UPDATE activities SET team = $1, sub_team = $2, ro_code = $3, name = $4, jenis = $5 WHERE id = $6`,
            [team, sub_team, ro_code, name, jenis, id]
        );
        res.redirect('/entry?msg=' + encodeURIComponent('Kegiatan berhasil diperbarui.'));
    } catch (err) {
        console.error(err);
        res.redirect('/entry?msg=' + encodeURIComponent('Gagal memperbarui kegiatan.'));
    }
});

// Fitur Import Excel Kegiatan (Sesuai struktur kolom: No | Kode RO | Tim | Sub Tim | Nama Kegiatan)
app.post('/entry/import', upload.single('excel_file'), async (req, res) => {
    try {
        if (!req.file) {
            return res.redirect('/entry?msg=File+belum+dipilih');
        }

        const workbook = xlsx.read(req.file.buffer, { type: 'buffer' });
        const extractedActivities = [];

        workbook.SheetNames.forEach(sheetName => {
            const sheet = workbook.Sheets[sheetName];
            const rows = xlsx.utils.sheet_to_json(sheet, { header: 1, defval: '' });

            rows.forEach((row) => {
                // Kolom A[0]: No, Kolom B[1]: Kode RO, Kolom C[2]: Tim, Kolom D[3]: Sub Tim, Kolom E[4]: Nama Kegiatan, Kolom F[5]: Jenis (opsional)
                const colB = String(row[1] || '').trim();
                const colC = String(row[2] || '').trim();
                const colD = String(row[3] || '').trim();
                const colE = String(row[4] || '').trim();
                const colF = String(row[5] || '').trim();

                // Lewati baris kosong atau baris header
                if (!colB || !colC || !colE) return;
                if (['kode ro', 'ro', 'tim', 'no'].includes(colB.toLowerCase()) || ['tim', 'tim kerja'].includes(colC.toLowerCase())) {
                    return;
                }

                extractedActivities.push({
                    ro_code: colB,
                    team: colC,
                    sub_team: colD || '-',
                    name: colE,
                    jenis: colF || (colE.toLowerCase().includes('publikasi') ? 'Publikasi' : (colE.toLowerCase().includes('laporan') ? 'Laporan' : 'Survei/Kegiatan'))
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
                        `INSERT INTO activities (team, sub_team, ro_code, name, jenis) VALUES ($1, $2, $3, $4, $5)`,
                        [item.team, item.sub_team, item.ro_code, item.name, item.jenis]
                    );
                    inserted++;
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
        const actRes = await db.query(`SELECT * FROM activities ORDER BY team`);
        const repRes = await db.query(`SELECT * FROM reports`);
        const teams = ['Sosial', 'Distribusi', 'Produksi', 'Neraca', 'IPDS', 'Sektoral', 'Umum'];
        res.render('pelaporan', { teams, activities: actRes.rows || [], reports: repRes.rows || [], activePage: 'pelaporan' });
    } catch (err) {
        console.error('Error pelaporan:', err);
        res.status(500).send('Database Error: ' + err.message);
    }
});

// Simpan Laporan dengan Validasi & Status Selesai Otomatis Jika Progress 100%
app.post('/pelaporan/save', async (req, res) => {
    try {
        const { activity_id, month, stage_type, stage_status, progress_percentage, description, problem, follow_up, proof_link } = req.body;
        const newProgress = parseFloat(progress_percentage);
        const mNum = parseInt(month);
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
                `UPDATE reports SET stage_status = $1, progress_percentage = $2, description = $3, problem = $4, follow_up = $5, proof_link = $6, is_completed = $7 WHERE id = $8`,
                [finalStatus, newProgress, description, problem, follow_up, proof_link, isCompletedFlag, rowId]
            );
            res.redirect('/pelaporan?success=1');
        } else {
            await db.query(
                `INSERT INTO reports (activity_id, month, stage_type, stage_status, progress_percentage, description, problem, follow_up, proof_link, is_completed) VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10)`,
                [activity_id, mNum, stage_type, finalStatus, newProgress, description, problem, follow_up, proof_link, isCompletedFlag]
            );
            res.redirect('/pelaporan?success=1');
        }
    } catch (err) {
        return res.redirect('/pelaporan?error=' + encodeURIComponent('Gagal menyimpan laporan ke database: ' + err.message));
    }
});

app.get('/api/report', async (req, res) => {
    try {
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