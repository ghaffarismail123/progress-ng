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
    if (totalActivities === 0) return { overallProgress: 0, monthlyTrend: Array(12).fill(0), teamProgress: [], statusCount: {} };

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
                monthSum += (latestReport.progress_percentage || 0);

                if (m === 11) { // Bulan November
                    if (latestReport.stage_status) statusCount[latestReport.stage_status] = (statusCount[latestReport.stage_status] || 0) + 1;
                }
            }
        });
        cumulativeByMonth[m - 1] = Number((monthSum / totalActivities).toFixed(2));
    }

    const completedActivities = activities.filter(a => {
        const actReports = reports.filter(r => r.activity_id === a.id);
        return actReports.some(r => r.progress_percentage >= 100 || r.is_completed === 1);
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
                tSum += (actReports[0].progress_percentage || 0);
            }
        });
        return { team, progress: Number((tSum / tTotal).toFixed(1)), total: tTotal };
    });

    return { totalActivities, completedActivities, overallProgress, monthlyTrend: cumulativeByMonth, teamProgress, statusCount };
}

// 1. Dashboard
app.get('/', (req, res) => {
    const selectedYear = req.query.year || '2026';
    db.all(`SELECT * FROM activities`, [], (err, activities) => {
        db.all(`SELECT * FROM reports`, [], (err, reports) => {
            const stats = calculateProgressData(activities || [], reports || []);
            res.render('dashboard', { ...stats, selectedYear, activePage: 'dashboard' });
        });
    });
});

// 2. Entri Kegiatan (Dengan pesan status import & hapus)
app.get('/entry', (req, res) => {
    const search = req.query.search || '';
    const msg = req.query.msg || '';
    const query = `SELECT * FROM activities WHERE name LIKE ? OR ro_code LIKE ? OR team LIKE ? OR sub_team LIKE ? ORDER BY team, id DESC`;
    db.all(query, [`%${search}%`, `%${search}%`, `%${search}%`, `%${search}%`], (err, activities) => {
        res.render('entry', { activities: activities || [], search, msg, activePage: 'entry' });
    });
});

// Tambah Kegiatan Manual -> Otomatis terintegrasi ke Pelaporan & Dashboard
app.post('/entry/add', (req, res) => {
    const { team, sub_team, ro_code, name, jenis } = req.body;
    db.run(
        `INSERT INTO activities (team, sub_team, ro_code, name, jenis) VALUES (?, ?, ?, ?, ?)`,
        [team, sub_team, ro_code, name, jenis],
        (err) => {
            if (err) return res.redirect(`/entry?msg=Gagal+menambah:+${encodeURIComponent(err.message)}`);
            res.redirect('/entry?msg=Kegiatan+baru+berhasil+ditambahkan');
        }
    );
});

// Tambahkan rute ini di file server Node.js utama Anda
app.post('/entry/edit', async (req, res) => {
    const { id, team, sub_team, ro_code, name, jenis } = req.body;
    
    // Contoh eksekusi query ke database SQLite
    db.run(
        `UPDATE activities SET team = ?, sub_team = ?, ro_code = ?, name = ?, jenis = ? WHERE id = ?`,
        [team, sub_team, ro_code, name, jenis, id],
        (err) => {
            if (err) {
                console.error(err);
                return res.redirect('/entry?msg=' + encodeURIComponent('Gagal memperbarui kegiatan.'));
            }
            res.redirect('/entry?msg=' + encodeURIComponent('Kegiatan berhasil diperbarui.'));
        }
    );
});

// Fitur Import Excel Kegiatan (Sesuai struktur kolom: No | Kode RO | Tim | Sub Tim | Nama Kegiatan)
app.post('/entry/import', upload.single('excel_file'), (req, res) => {
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

        db.serialize(() => {
            db.run(`CREATE TABLE IF NOT EXISTS activities (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                team TEXT NOT NULL,
                sub_team TEXT,
                ro_code TEXT NOT NULL,
                name TEXT NOT NULL,
                jenis TEXT,
                created_at DATETIME DEFAULT CURRENT_TIMESTAMP
            )`);

            let completed = 0;
            let inserted = 0;
            let lastError = '';

            extractedActivities.forEach(item => {
                db.get(
                    `SELECT id FROM activities WHERE ro_code = ? AND name = ?`,
                    [item.ro_code, item.name],
                    (err, existing) => {
                        if (existing) {
                            completed++;
                            if (completed === extractedActivities.length) {
                                res.redirect(`/entry?msg=Berhasil+mengimpor+${inserted}+kegiatan+baru+(${extractedActivities.length - inserted}+data+sudah+ada+sebelumnya)`);
                            }
                        } else {
                            db.run(
                                `INSERT INTO activities (team, sub_team, ro_code, name, jenis) VALUES (?, ?, ?, ?, ?)`,
                                [item.team, item.sub_team, item.ro_code, item.name, item.jenis],
                                function (insertErr) {
                                    if (insertErr) {
                                        console.error('SQLite Insert Error:', insertErr.message);
                                        lastError = insertErr.message;
                                    } else {
                                        inserted++;
                                    }
                                    completed++;
                                    if (completed === extractedActivities.length) {
                                        if (inserted === 0 && lastError) {
                                            res.redirect(`/entry?msg=Gagal+menyimpan+ke+database:+${encodeURIComponent(lastError)}`);
                                        } else {
                                            res.redirect(`/entry?msg=Berhasil+mengimpor+${inserted}+kegiatan+baru+dari+${extractedActivities.length}+baris+Excel`);
                                        }
                                    }
                                }
                            );
                        }
                    }
                );
            });
        });
    } catch (err) {
        console.error('Error import Excel:', err);
        res.redirect(`/entry?msg=Error:+${encodeURIComponent(err.message)}`);
    }
});

// Hapus Semua Kegiatan & Laporan (Dengan Konfirmasi Password "0000")
app.post('/entry/delete-all', (req, res) => {
    const { password } = req.body;
    if (password !== '0000') {
        return res.redirect('/entry?msg=Gagal+menghapus+semua:+Password+konfirmasi+salah!');
    }

    db.serialize(() => {
        db.run(`DELETE FROM reports`, [], (err) => {
            if (err) {
                return res.redirect('/entry?msg=Gagal+menghapus+laporan:+' + encodeURIComponent(err.message));
            }
            db.run(`DELETE FROM activities`, [], (err2) => {
                if (err2) {
                    return res.redirect('/entry?msg=Gagal+menghapus+kegiatan:+' + encodeURIComponent(err2.message));
                }
                db.run(`DELETE FROM sqlite_sequence WHERE name IN ('activities', 'reports')`, [], () => {
                    res.redirect('/entry?msg=Seluruh+data+kegiatan+dan+laporan+berhasil+dihapus');
                });
            });
        });
    });
});

app.post('/entry/delete/:id', (req, res) => {
    const id = req.params.id;
    db.run(`DELETE FROM activities WHERE id = ?`, [id], (err) => {
        db.run(`DELETE FROM reports WHERE activity_id = ?`, [id], () => {
            res.redirect('/entry');
        });
    });
});

// 3. Pelaporan Bulanan (Otomatis mengambil data dari tabel activities pusat)
app.get('/pelaporan', (req, res) => {
    db.all(`SELECT * FROM activities ORDER BY team`, [], (err, activities) => {
        db.all(`SELECT * FROM reports`, [], (err, reports) => {
            const teams = ['Sosial', 'Distribusi', 'Produksi', 'Neraca', 'IPDS', 'Sektoral', 'Umum'];
            res.render('pelaporan', { teams, activities: activities || [], reports: reports || [], activePage: 'pelaporan' });
        });
    });
});

// Simpan Laporan dengan Validasi & Status Selesai Otomatis Jika Progress 100%
app.post('/pelaporan/save', (req, res) => {
    const { activity_id, month, stage_type, stage_status, progress_percentage, description, problem, follow_up, proof_link } = req.body;
    const newProgress = parseFloat(progress_percentage);
    const mNum = parseInt(month);
    const isCompletedFlag = newProgress >= 100 ? 1 : 0;
    const finalStatus = newProgress >= 100 ? 'Selesai' : stage_status;

    if (isNaN(newProgress) || newProgress < 0 || newProgress > 100) {
        return res.redirect('/pelaporan?error=' + encodeURIComponent('Besaran progress harus berupa angka antara 0% sampai 100%.'));
    }

    db.all(`SELECT month, progress_percentage FROM reports WHERE activity_id = ? AND month < ? ORDER BY month DESC`, [activity_id, mNum], (err, rows) => {
        if (err) {
            return res.redirect('/pelaporan?error=' + encodeURIComponent('Gagal membaca riwayat laporan: ' + err.message));
        }

        if (rows && rows.length > 0) {
            const prevProgress = rows[0].progress_percentage || 0;
            const prevMonth = rows[0].month;
            if (newProgress < prevProgress) {
                const reason = `Besaran progress yang diinput (${newProgress}%) tidak boleh lebih kecil dari bulan sebelumnya (Bulan ke-${prevMonth}: ${prevProgress}%).`;
                return res.redirect('/pelaporan?error=' + encodeURIComponent(reason));
            }
        }

        db.get(`SELECT id FROM reports WHERE activity_id = ? AND month = ? AND stage_type = ?`, [activity_id, mNum, stage_type], (err, row) => {
            if (err) {
                return res.redirect('/pelaporan?error=' + encodeURIComponent('Gagal memeriksa data laporan: ' + err.message));
            }

            if (row) {
                db.run(`UPDATE reports SET stage_status = ?, progress_percentage = ?, description = ?, problem = ?, follow_up = ?, proof_link = ?, is_completed = ? WHERE id = ?`,
                    [finalStatus, newProgress, description, problem, follow_up, proof_link, isCompletedFlag, row.id], (updateErr) => {
                        if (updateErr) {
                            return res.redirect('/pelaporan?error=' + encodeURIComponent('Gagal memperbarui laporan ke database: ' + updateErr.message));
                        }
                        res.redirect('/pelaporan?success=1');
                    });
            } else {
                db.run(`INSERT INTO reports (activity_id, month, stage_type, stage_status, progress_percentage, description, problem, follow_up, proof_link, is_completed) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
                    [activity_id, mNum, stage_type, finalStatus, newProgress, description, problem, follow_up, proof_link, isCompletedFlag], (insertErr) => {
                        if (insertErr) {
                            return res.redirect('/pelaporan?error=' + encodeURIComponent('Gagal menyimpan laporan baru ke database: ' + insertErr.message));
                        }
                        res.redirect('/pelaporan?success=1');
                    });
            }
        });
    });
});

app.get('/api/report', (req, res) => {
    const { activity_id, month, stage_type } = req.query;
    db.get(`SELECT * FROM reports WHERE activity_id = ? AND month = ? AND stage_type = ?`, [activity_id, month, stage_type], (err, row) => {
        res.json(row || {});
    });
});

app.listen(PORT, () => {
    console.log(`PROGRES.NG Server running on http://localhost:${PORT}`);
});