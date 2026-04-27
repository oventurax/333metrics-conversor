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

// Normalize any date format to "DD/MM/YYYY":
//   "YYYY-MM-DD"  → "DD/MM/YYYY"
//   "MM/DD/YYYY"  → "DD/MM/YYYY"  (American — detected when 2nd part > 12)
//   "DD/MM/YYYY"  → unchanged
function normalizeDate(raw) {
  if (!raw) return null;
  const s = String(raw).trim();

  // ISO format: YYYY-MM-DD
  if (s.match(/^\d{4}-\d{2}-\d{2}$/)) {
    const [yyyy, mm, dd] = s.split('-');
    return `${dd}/${mm}/${yyyy}`;
  }

  // Slash-separated: DD/MM/YYYY or MM/DD/YYYY
  if (s.match(/^\d{1,2}\/\d{1,2}\/\d{4}$/)) {
    const [a, b, yyyy] = s.split('/');
    const numA = parseInt(a, 10);
    const numB = parseInt(b, 10);
    // If b > 12, the second part is the day → American MM/DD/YYYY → swap
    if (numB > 12) {
      return `${b.padStart(2, '0')}/${a.padStart(2, '0')}/${yyyy}`;
    }
    // Otherwise treat as DD/MM/YYYY (Brazilian)
    return `${a.padStart(2, '0')}/${b.padStart(2, '0')}/${yyyy}`;
  }

  return null;
}

// Strip currency symbols/codes and parse as float.
// Handles: $, €, R$, USD, EUR — and Brazilian decimal comma (1.200,50 → 1200.50)
function parseNumber(val) {
  if (val == null || val === '') return 0;
  let s = String(val).trim();
  s = s.replace(/R\$|USD|EUR|\$|€/g, '').trim();
  // "1.200,50" → "1200.50"
  if (s.includes(',') && s.includes('.')) {
    s = s.replace(/\./g, '').replace(',', '.');
  } else if (s.includes(',')) {
    // "1200,50" → "1200.50"
    s = s.replace(',', '.');
  }
  return parseFloat(s) || 0;
}

function parseCampaigns(rows, platform) {
  const header = rows[0].map(h => String(h).trim().toLowerCase());

  let colCampaign, colConvValue, colCost, colDateStart;
  let colAdSet = -1, colAd = -1, colCurrency = -1;

  if (platform === 'facebook') {
    colCampaign  = header.findIndex(h => h.includes('nome da campanha'));
    colAdSet     = header.findIndex(h => h.includes('nome do conjunto'));
    colAd        = header.findIndex(h => h.includes('nome do anúncio') || h.includes('nome do anuncio'));
    colCurrency  = header.findIndex(h => h === 'moeda' || (h.includes('moeda') && !h.includes('valor')));
    colConvValue = header.findIndex(h => h.includes('valor de convers'));
    colCost      = header.findIndex(h => h.includes('valor usado'));
    colDateStart = header.findIndex(h => h.includes('início dos relatórios') || h.includes('inicio dos relatorios'));
  } else {
    // TikTok: English columns, conversions = quantity × 230
    colCampaign  = header.findIndex(h => h.includes('campaign name') || h.includes('campaign'));
    colConvValue = header.findIndex(h => h.includes('conversions'));
    colCost      = header.findIndex(h => h.includes('cost'));
    colDateStart = -1;
  }

  if (colCampaign === -1 || colConvValue === -1 || colCost === -1) {
    const plat = platform === 'facebook' ? 'Facebook Ads' : 'TikTok Ads';
    throw new Error(`Colunas não encontradas. Verifique se é uma planilha do ${plat}.`);
  }

  // Extract date from first valid Facebook row
  let fileDate = null;
  if (platform === 'facebook' && colDateStart !== -1) {
    for (let i = 1; i < rows.length; i++) {
      const d = normalizeDate(rows[i][colDateStart]);
      if (d) { fileDate = d; break; }
    }
  }

  const campaigns = [];
  for (let i = 1; i < rows.length; i++) {
    const row  = rows[i];
    const name = String(row[colCampaign] || '').trim();

    // Skip empty campaign names (Meta totals row) or TikTok footer
    if (!name || name.toLowerCase().startsWith('total of')) continue;

    const rawConv   = parseNumber(row[colConvValue]);
    const cost      = parseNumber(row[colCost]);
    const convValue = platform === 'facebook' ? rawConv : rawConv * 230;

    if (platform === 'facebook') {
      const adSet    = String(row[colAdSet]    || '').trim();
      const ad       = String(row[colAd]       || '').trim();
      const currency = colCurrency !== -1 ? String(row[colCurrency] || '').trim() : '';
      const rowDate  = colDateStart !== -1 ? normalizeDate(row[colDateStart]) : null;
      campaigns.push({ name, adSet, ad, currency, convValue, cost, rowDate });
    } else {
      campaigns.push({ name, convValue, cost });
    }
  }

  return { campaigns, fileDate };
}

async function generateWorkbook(allCampaigns, reportDate, platform) {
  const dateStr  = formatDate(reportDate);
  const workbook = new ExcelJS.Workbook();
  const ws       = workbook.addWorksheet('Worksheet');

  const isFB    = platform === 'facebook';
  const numCols = isFB ? 7 : 5;

  ws.columns = isFB
    ? [{ width: 18 }, { width: 10 }, { width: 42 }, { width: 38 }, { width: 38 }, { width: 22 }, { width: 22 }]
    : [{ width: 18 }, { width: 22 }, { width: 45 }, { width: 28 }, { width: 22 }];

  const DARK_BLUE  = '1F4E79';
  const LIGHT_BLUE = 'DCE6F1';
  const CYAN       = '17A8B8';
  const WHITE      = 'FFFFFFFF';

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
  const headerValues = isFB
    ? ['Data', 'Moeda', 'Campanha', 'Conjunto', 'Anúncio', 'Faturado', 'Gasto']
    : ['Início dos relatórios', 'Encerramento dos relatórios', 'Nome da campanha', 'Valor de conversão da compra', 'Valor usado (USD)'];

  const headerRow = ws.addRow(headerValues);
  headerRow.height = 26.25;
  headerRow.eachCell(cell => {
    cell.fill = fills.header;
    cell.font = { name: 'Arial', size: 11, bold: true, color: { argb: WHITE } };
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    cell.border = border;
  });
  ws.views = [{ state: 'frozen', ySplit: 1 }];

  const fmt = v => v.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });

  // Data rows
  allCampaigns.forEach((c, idx) => {
    const fill     = idx % 2 === 0 ? fills.white : fills.lightBlue;
    const convCell = c.convValue > 0 ? fmt(c.convValue) : null;
    const costCell = c.cost > 0 ? fmt(c.cost) : null;

    const rowValues = isFB
      ? [c.rowDate || dateStr, c.currency, c.name, c.adSet, c.ad, convCell, costCell]
      : [dateStr, dateStr, c.name, convCell, costCell];

    const row = ws.addRow(rowValues);
    for (let col = 1; col <= numCols; col++) {
      const cell = row.getCell(col);
      cell.fill = fill;
      cell.font = { name: 'Arial', size: 11 };
      cell.border = border;
      cell.alignment = { vertical: 'middle' };
      const isValueCol = isFB ? (col === 6 || col === 7) : (col === 4 || col === 5);
      if (isValueCol) cell.alignment = { vertical: 'middle', horizontal: 'right' };
    }
  });

  // Separator
  ws.addRow([]);

  // Totals
  const totalSold  = allCampaigns.reduce((s, c) => s + (c.convValue > 0 ? c.convValue : 0), 0);
  const totalSpent = allCampaigns.reduce((s, c) => s + (c.cost > 0 ? c.cost : 0), 0);

  const totalsValues = isFB
    ? [null, null, null, null, null, `Total vendido: $${fmt(totalSold)}`, `Total gasto: $${fmt(totalSpent)}`]
    : [null, null, null, `Total vendido: $${fmt(totalSold)}`, `Total gasto: $${fmt(totalSpent)}`];

  const totalsRow = ws.addRow(totalsValues);
  for (let col = 1; col <= numCols; col++) {
    const cell = totalsRow.getCell(col);
    cell.fill = fills.darkBlue;
    cell.font = { name: 'Arial', size: 11, bold: true, color: { argb: WHITE } };
    cell.border = border;
    const isLabelCol = isFB ? col >= 6 : col >= 4;
    cell.alignment = { vertical: 'middle', horizontal: isLabelCol ? 'center' : 'left' };
  }

  // Profit
  const profit    = totalSold - totalSpent;
  const profitStr = profit >= 0 ? `$${fmt(profit)}` : `-$${fmt(Math.abs(profit))}`;

  const profitValues = isFB
    ? [null, null, null, null, null, 'LUCRO FINAL:', profitStr]
    : [null, null, null, 'LUCRO FINAL:', profitStr];

  const profitRow = ws.addRow(profitValues);
  for (let col = 1; col <= numCols; col++) {
    const cell = profitRow.getCell(col);
    cell.fill = fills.cyan;
    cell.font = { name: 'Arial', size: 11, bold: true, color: { argb: WHITE } };
    cell.border = border;
    const isLabelCol = isFB ? col >= 6 : col >= 4;
    cell.alignment = { vertical: 'middle', horizontal: isLabelCol ? 'center' : 'left' };
  }

  return workbook.xlsx.writeBuffer();
}

app.post('/convert', upload.array('files', 5), async (req, res) => {
  try {
    if (!req.files || req.files.length === 0) {
      return res.status(400).json({ error: 'Nenhum arquivo enviado.' });
    }

    const platform     = req.body.platform === 'facebook' ? 'facebook' : 'tiktok';
    const userPickedDate = !!(req.body.date && req.body.date.trim());
    let reportDate     = userPickedDate ? parseDate(req.body.date.trim()) : getYesterday();

    const allCampaigns = [];

    for (const file of req.files) {
      const wb   = XLSX.read(file.buffer, { type: 'buffer' });
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: '' });

      if (rows.length < 2) throw new Error(`"${file.originalname}": planilha sem dados.`);

      const { campaigns, fileDate } = parseCampaigns(rows, platform);
      if (campaigns.length === 0) throw new Error(`"${file.originalname}": nenhuma campanha encontrada.`);

      // Facebook: use the file's own date if the user didn't pick one
      if (platform === 'facebook' && !userPickedDate && fileDate) {
        reportDate = parseDate(fileDate);
      }

      allCampaigns.push(...campaigns);
    }

    if (allCampaigns.length === 0) {
      return res.status(400).json({ error: 'Nenhuma campanha encontrada nos arquivos enviados.' });
    }

    const buffer = await generateWorkbook(allCampaigns, reportDate, platform);
    const suffix = dateSuffix(reportDate);
    const outName = `${suffix}.xlsx`;

    res.json({
      name:       outName,
      data:       buffer.toString('base64'),
      dateSuffix: suffix,
    });

  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message || 'Erro interno.' });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`\n✅ Servidor rodando em http://localhost:${PORT}\n`);
});
