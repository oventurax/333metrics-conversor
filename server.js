const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const ExcelJS = require('exceljs');
const path = require('path');

const app = express();
const upload = multer({ storage: multer.memoryStorage() });

app.use(express.static(path.join(__dirname, 'public')));

function formatDate(d) {
  const dd = String(d.getDate()).padStart(2, '0');
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  const yyyy = d.getFullYear();
  return `${dd}/${mm}/${yyyy}`;
}

function getYesterday() {
  const d = new Date();
  d.setDate(d.getDate() - 1);
  return d;
}

function parseDate(str) {
  const [dd, mm, yyyy] = str.split('/');
  return new Date(parseInt(yyyy), parseInt(mm) - 1, parseInt(dd));
}

function dateSuffix(d) {
  const dd = String(d.getDate()).padStart(2, '0');
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  return `${dd}-${mm}-${d.getFullYear()}`;
}

function parseCampaigns(rows, platform) {
  const header = rows[0].map(h => String(h).trim().toLowerCase());

  let colCampaign, colConvValue, colCost, colDateStart;

  if (platform === 'facebook') {
    // Facebook raw export: Portuguese column names, conversion value already in $
    // Columns: Nome da campanha | Moeda | Valor de conversão da compra | Valor usado | Início dos relatórios | Encerramento dos relatórios
    colCampaign  = header.findIndex(h => h.includes('nome da campanha'));
    colConvValue = header.findIndex(h => h.includes('valor de convers'));
    colCost      = header.findIndex(h => h.includes('valor usado'));
    colDateStart = header.findIndex(h => h.includes('início dos relatórios') || h.includes('inicio dos relatorios'));
  } else {
    // TikTok raw export: English column names, conversions = quantity × 230
    colCampaign  = header.findIndex(h => h.includes('campaign name') || h.includes('campaign'));
    colConvValue = header.findIndex(h => h.includes('conversions'));
    colCost      = header.findIndex(h => h.includes('cost'));
    colDateStart = -1;
  }

  if (colCampaign === -1 || colConvValue === -1 || colCost === -1) {
    const plat = platform === 'facebook' ? 'Facebook Ads' : 'TikTok Ads';
    throw new Error(`Colunas não encontradas. Verifique se é uma planilha do ${plat}.`);
  }

  // For Facebook: extract date from first valid data row (YYYY-MM-DD → DD/MM/YYYY)
  let fileDate = null;
  if (platform === 'facebook' && colDateStart !== -1) {
    for (let i = 1; i < rows.length; i++) {
      const raw = String(rows[i][colDateStart] || '').trim();
      if (raw.match(/^\d{4}-\d{2}-\d{2}$/)) {
        const [yyyy, mm, dd] = raw.split('-');
        fileDate = `${dd}/${mm}/${yyyy}`;
        break;
      }
    }
  }

  const campaigns = [];
  for (let i = 1; i < rows.length; i++) {
    const row  = rows[i];
    const name = String(row[colCampaign] || '').trim();

    // Skip: empty campaign name (Facebook total rows) or TikTok "Total of X" footer
    if (!name || name.toLowerCase().startsWith('total of')) continue;

    const rawConv   = parseFloat(row[colConvValue]) || 0;
    const cost      = parseFloat(row[colCost]) || 0;
    const convValue = platform === 'facebook' ? rawConv : rawConv * 230;

    campaigns.push({ name, convValue, cost });
  }

  return { campaigns, fileDate };
}

async function processFile(fileBuffer, originalName, reportDate, platform, userPickedDate) {
  let dateStr = formatDate(reportDate);

  const wb   = XLSX.read(fileBuffer, { type: 'buffer' });
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: '' });

  if (rows.length < 2) throw new Error(`"${originalName}": planilha sem dados.`);

  const { campaigns, fileDate } = parseCampaigns(rows, platform);
  if (campaigns.length === 0) throw new Error(`"${originalName}": nenhuma campanha encontrada.`);

  // Facebook: if user didn't pick a date, use the date from the file itself
  if (platform === 'facebook' && !userPickedDate && fileDate) {
    reportDate = parseDate(fileDate);
    dateStr = fileDate;
  }

  // ── Build formatted workbook ──────────────────────────────────────────────
  const workbook = new ExcelJS.Workbook();
  const ws = workbook.addWorksheet('Worksheet');

  ws.columns = [
    { width: 18 }, { width: 22 }, { width: 45 }, { width: 28 }, { width: 22 },
  ];

  const DARK_BLUE = '1F4E79';
  const LIGHT_BLUE = 'DCE6F1';
  const CYAN = '17A8B8';
  const WHITE = 'FFFFFFFF';

  const fills = {
    header:    { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + DARK_BLUE } },
    lightBlue: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + LIGHT_BLUE } },
    white:     { type: 'pattern', pattern: 'solid', fgColor: { argb: WHITE } },
    darkBlue:  { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + DARK_BLUE } },
    cyan:      { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + CYAN } },
  };

  const border = {
    top: { style: 'thin' }, left: { style: 'thin' },
    bottom: { style: 'thin' }, right: { style: 'thin' },
  };

  // Header row
  const headerRow = ws.addRow([
    'Início dos relatórios', 'Encerramento dos relatórios',
    'Nome da campanha', 'Valor de conversão da compra', 'Valor usado (USD)',
  ]);
  headerRow.height = 26.25;
  headerRow.eachCell(cell => {
    cell.fill = fills.header;
    cell.font = { name: 'Arial', size: 11, bold: true, color: { argb: WHITE } };
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    cell.border = border;
  });
  ws.views = [{ state: 'frozen', ySplit: 1 }];

  // Data rows
  campaigns.forEach((c, idx) => {
    const fill = idx % 2 === 0 ? fills.white : fills.lightBlue;
    const convCell = c.convValue > 0 ? c.convValue : null;
    const costCell = c.cost > 0 ? c.cost : null;

    const row = ws.addRow([dateStr, dateStr, c.name, convCell, costCell]);
    for (let col = 1; col <= 5; col++) {
      const cell = row.getCell(col);
      cell.fill = fill;
      cell.font = { name: 'Arial', size: 11 };
      cell.border = border;
      cell.alignment = { vertical: 'middle' };
      if (col === 4 || col === 5) {
        cell.numFmt = '#,##0.00';
        cell.alignment = { vertical: 'middle', horizontal: 'right' };
      }
    }
  });

  // Separator
  ws.addRow([]);

  // Totals
  const totalSold  = campaigns.reduce((s, c) => s + (c.convValue > 0 ? c.convValue : 0), 0);
  const totalSpent = campaigns.reduce((s, c) => s + (c.cost > 0 ? c.cost : 0), 0);

  const fmt = v => v.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });

  const totalsRow = ws.addRow([null, null, null,
    `Total vendido: $${fmt(totalSold)}`,
    `Total gasto: $${fmt(totalSpent)}`,
  ]);
  for (let col = 1; col <= 5; col++) {
    const cell = totalsRow.getCell(col);
    cell.fill = fills.darkBlue;
    cell.font = { name: 'Arial', size: 11, bold: true, color: { argb: WHITE } };
    cell.border = border;
    cell.alignment = { vertical: 'middle', horizontal: col >= 4 ? 'center' : 'left' };
  }

  // Profit
  const profit = totalSold - totalSpent;
  const profitStr = profit >= 0 ? `$${fmt(profit)}` : `-$${fmt(Math.abs(profit))}`;

  const profitRow = ws.addRow([null, null, null, 'LUCRO FINAL:', profitStr]);
  for (let col = 1; col <= 5; col++) {
    const cell = profitRow.getCell(col);
    cell.fill = fills.cyan;
    cell.font = { name: 'Arial', size: 11, bold: true, color: { argb: WHITE } };
    cell.border = border;
    cell.alignment = { vertical: 'middle', horizontal: col >= 4 ? 'center' : 'left' };
  }

  const buffer = await workbook.xlsx.writeBuffer();
  const baseName = path.basename(originalName, path.extname(originalName));
  const outName = `${baseName}_${dateSuffix(reportDate)}.xlsx`;

  return { buffer, outName, usedDate: reportDate };
}

app.post('/convert', upload.array('files', 5), async (req, res) => {
  try {
    if (!req.files || req.files.length === 0) {
      return res.status(400).json({ error: 'Nenhum arquivo enviado.' });
    }

    const platform = req.body.platform === 'facebook' ? 'facebook' : 'tiktok';

    const userPickedDate = !!(req.body.date && req.body.date.trim());
    let reportDate = userPickedDate ? parseDate(req.body.date.trim()) : getYesterday();

    const results = [];
    for (const file of req.files) {
      const { buffer, outName, usedDate } = await processFile(
        file.buffer, file.originalname, reportDate, platform, userPickedDate
      );
      results.push({ name: outName, data: buffer.toString('base64'), dateSuffix: dateSuffix(usedDate) });
    }

    // dateSuffix for display: use first file's date (all files share same date)
    res.json({ files: results, dateSuffix: results[0].dateSuffix });

  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message || 'Erro interno.' });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`\n✅ Servidor rodando em http://localhost:${PORT}\n`);
});
