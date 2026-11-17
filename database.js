const { Pool } = require('pg');

const pool = new Pool({
    connectionString: process.env.DATABASE_URL,
    ssl: {
        rejectUnauthorized: false
    }
});

pool.on('connect', () => {
    console.log('Connected to PostgreSQL database (Supabase) for PROGRES.NG.');
});

pool.on('error', (err) => {
    console.error('Unexpected error on idle client', err);
    process.exit(-1);
});

// Fungsi inisialisasi tabel otomatis di Supabase (PostgreSQL)
async function initDB() {
    try {
        // Buat tabel activities jika belum ada
        await pool.query(`
            CREATE TABLE IF NOT EXISTS activities (
                id SERIAL PRIMARY KEY,
                year INTEGER NOT NULL DEFAULT 2026,
                team VARCHAR(50) NOT NULL,
                sub_team VARCHAR(100),
                ro_code VARCHAR(50) NOT NULL,
                name TEXT NOT NULL,
                jenis VARCHAR(50),
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            );
        `);

        // Cek dan tambahkan kolom jika tabel lama belum memilikinya (migrasi otomatis)
        await pool.query(`
            DO $$ 
            BEGIN 
                IF NOT EXISTS (SELECT 1 FROM information_schema.columns WHERE table_name='activities' AND column_name='year') THEN
                    ALTER TABLE activities ADD COLUMN year INTEGER NOT NULL DEFAULT 2026;
                END IF;
                IF NOT EXISTS (SELECT 1 FROM information_schema.columns WHERE table_name='activities' AND column_name='sub_team') THEN
                    ALTER TABLE activities ADD COLUMN sub_team VARCHAR(100);
                END IF;
                IF NOT EXISTS (SELECT 1 FROM information_schema.columns WHERE table_name='activities' AND column_name='jenis') THEN
                    ALTER TABLE activities ADD COLUMN jenis VARCHAR(50);
                END IF;
            END $$;
        `);

        // Buat tabel reports jika belum ada
        await pool.query(`
            CREATE TABLE IF NOT EXISTS reports (
                id SERIAL PRIMARY KEY,
                activity_id INTEGER REFERENCES activities(id) ON DELETE CASCADE,
                year INTEGER NOT NULL DEFAULT 2026,
                month INTEGER NOT NULL,
                stage_type VARCHAR(50),
                stage_status VARCHAR(50),
                progress_percentage NUMERIC(5,2) DEFAULT 0,
                description TEXT,
                problem TEXT,
                follow_up TEXT,
                proof_link TEXT,
                is_completed INTEGER DEFAULT 0
            );
        `);

        // Migrasi otomatis kolom year pada tabel reports jika belum ada
        await pool.query(`
            DO $$ 
            BEGIN 
                IF NOT EXISTS (SELECT 1 FROM information_schema.columns WHERE table_name='reports' AND column_name='year') THEN
                    ALTER TABLE reports ADD COLUMN year INTEGER NOT NULL DEFAULT 2026;
                END IF;
            END $$;
        `);

        console.log('Database tables & migrations checked/initialized successfully.');
    } catch (err) {
        console.error('Error initializing database tables:', err.message);
    }
}

// Jalankan inisialisasi
initDB();

// Helper ekspor query dan pool
module.exports = {
    query: (text, params) => pool.query(text, params),
    pool
};