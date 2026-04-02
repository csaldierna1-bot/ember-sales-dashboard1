const express = require('express');
const cors = require('cors');
const compression = require('compression');
const { Pool } = require('pg');

const app = express();
const PORT = process.env.PORT || 3000;

// ── DATABASE ────────────────────────────────────────────────────────────────
const pool = new Pool({
  connectionString: process.env.DATABASE_URL,
  ssl: process.env.NODE_ENV === 'production' ? { rejectUnauthorized: false } : false,
});

// ── MIDDLEWARE ───────────────────────────────────────────────────────────────
app.use(cors());
app.use(compression());
app.use(express.json({ limit: '50mb' }));

// ── INIT DATABASE ───────────────────────────────────────────────────────────
async function initDB() {
  const client = await pool.connect();
  try {
    await client.query(`
      CREATE TABLE IF NOT EXISTS imports (
        id SERIAL PRIMARY KEY,
        gross_data JSONB NOT NULL,
        cancel_data JSONB NOT NULL,
        data_date VARCHAR(20) NOT NULL,
        imported_by VARCHAR(100),
        created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW()
      );
    `);
    console.log('Database initialized');
  } finally {
    client.release();
  }
}

// ── ROUTES ──────────────────────────────────────────────────────────────────

// Health check
app.get('/api/health', (req, res) => {
  res.json({ status: 'ok', timestamp: new Date().toISOString() });
});

// Save imported data
app.post('/api/import', async (req, res) => {
  try {
    const { grossData, cancelData, dataDate, importedBy } = req.body;

    if (!grossData || !cancelData || !dataDate) {
      return res.status(400).json({ error: 'Missing required fields: grossData, cancelData, dataDate' });
    }

    const result = await pool.query(
      `INSERT INTO imports (gross_data, cancel_data, data_date, imported_by)
       VALUES ($1, $2, $3, $4)
       RETURNING id, data_date, imported_by, created_at`,
      [JSON.stringify(grossData), JSON.stringify(cancelData), dataDate, importedBy || 'anonymous']
    );

    res.json({
      success: true,
      import: result.rows[0],
      message: 'Data saved successfully'
    });
  } catch (err) {
    console.error('Import error:', err);
    res.status(500).json({ error: 'Failed to save import data' });
  }
});

// Get latest imported data
app.get('/api/data', async (req, res) => {
  try {
    const result = await pool.query(
      `SELECT id, gross_data, cancel_data, data_date, imported_by, created_at
       FROM imports ORDER BY created_at DESC LIMIT 1`
    );

    if (result.rows.length === 0) {
      return res.json({ hasData: false });
    }

    const row = result.rows[0];
    res.json({
      hasData: true,
      grossData: row.gross_data,
      cancelData: row.cancel_data,
      dataDate: row.data_date,
      importedBy: row.imported_by,
      createdAt: row.created_at
    });
  } catch (err) {
    console.error('Fetch error:', err);
    res.status(500).json({ error: 'Failed to fetch data' });
  }
});

// Get import history
app.get('/api/history', async (req, res) => {
  try {
    const result = await pool.query(
      `SELECT id, data_date, imported_by, created_at
       FROM imports ORDER BY created_at DESC LIMIT 10`
    );
    res.json({ imports: result.rows });
  } catch (err) {
    console.error('History error:', err);
    res.status(500).json({ error: 'Failed to fetch history' });
  }
});

// Load a specific import by ID
app.get('/api/data/:id', async (req, res) => {
  try {
    const result = await pool.query(
      `SELECT id, gross_data, cancel_data, data_date, imported_by, created_at
       FROM imports WHERE id = $1`,
      [req.params.id]
    );

    if (result.rows.length === 0) {
      return res.status(404).json({ error: 'Import not found' });
    }

    const row = result.rows[0];
    res.json({
      hasData: true,
      grossData: row.gross_data,
      cancelData: row.cancel_data,
      dataDate: row.data_date,
      importedBy: row.imported_by,
      createdAt: row.created_at
    });
  } catch (err) {
    console.error('Fetch error:', err);
    res.status(500).json({ error: 'Failed to fetch data' });
  }
});

// Delete all imports (reset)
app.delete('/api/data', async (req, res) => {
  try {
    await pool.query('DELETE FROM imports');
    res.json({ success: true, message: 'All import data cleared' });
  } catch (err) {
    console.error('Delete error:', err);
    res.status(500).json({ error: 'Failed to clear data' });
  }
});

// ── START ───────────────────────────────────────────────────────────────────
initDB().then(() => {
  app.listen(PORT, () => {
    console.log(`Ember Sales API running on port ${PORT}`);
  });
}).catch(err => {
  console.error('Failed to initialize database:', err);
  process.exit(1);
});
