const express = require('express');
const cors = require('cors');
const path = require('path');
const XLSX = require('xlsx');

const app = express();
const PORT = process.env.PORT || 3000;
const FILE = process.env.RESUME_XLSX || path.join(__dirname, 'data', 'resume.xlsx');

// "Team Size" -> "teamSize"
const camel = (s) =>
  String(s).trim().replace(/[^a-zA-Z0-9]+(.)?/g, (_, c) => (c ? c.toUpperCase() : ''))
    .replace(/^./, (c) => c.toLowerCase());

// Re-reads the file on every request, so saved Excel edits show up on refresh.
function readWorkbook() {
  const wb = XLSX.readFile(FILE);
  const out = {};
  for (const name of wb.SheetNames) {
    if (name.startsWith('_')) continue; // private tabs
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { defval: '', raw: false });
    out[camel(name)] = rows.map((row) =>
      Object.fromEntries(Object.entries(row).map(([k, v]) => [camel(k), typeof v === 'string' ? v.trim() : v]))
    );
  }
  return out;
}

app.use(cors());
app.use((req, res, next) => { res.set('Cache-Control', 'no-store'); next(); });

app.get('/api/resume', (req, res) => {
  try { res.json(readWorkbook()); }
  catch (e) { res.status(500).json({ error: `Could not read ${FILE}: ${e.message}` }); }
});

app.get('/api/resume/:sheet', (req, res) => {
  try {
    const data = readWorkbook()[camel(req.params.sheet)];
    if (!data) return res.status(404).json({ error: `No sheet named "${req.params.sheet}"` });
    res.json(data);
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.listen(PORT, () => console.log(`Resume API on http://localhost:${PORT}  (reading ${FILE})`));
