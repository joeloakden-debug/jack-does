const ExcelJS = require('exceljs');

// ========================================
// STYLES
// ========================================
const headerFont = { name: 'Calibri', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
const headerFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1a1a2e' } };
const headerBorder = { bottom: { style: 'thin', color: { argb: 'FF4a4a6e' } } };
const dataFont = { name: 'Calibri', size: 10 };
const altRowFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8F9FA' } };
const currencyFormat = '$#,##0.00';
const percentFormat = '0.00%';
const dateFormat = 'yyyy-mm-dd';

function styleHeaders(sheet) {
  const headerRow = sheet.getRow(1);
  headerRow.eachCell(cell => {
    cell.font = headerFont;
    cell.fill = headerFill;
    cell.border = headerBorder;
    cell.alignment = { vertical: 'middle', horizontal: 'left' };
  });
  headerRow.height = 24;
  sheet.views = [{ state: 'frozen', ySplit: 1 }];
}

function styleDataRows(sheet) {
  for (let i = 2; i <= sheet.rowCount; i++) {
    const row = sheet.getRow(i);
    row.eachCell(cell => {
      cell.font = dataFont;
      if (i % 2 === 0) cell.fill = altRowFill;
    });
  }
}

// ========================================
// GENERATE WORKBOOK
// ========================================
async function generateWorkbook(clientId, clientName, clientData) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Jack Does Accounting';
  workbook.created = new Date();

  const assets = clientData.assets || [];
  const classes = clientData.assetClasses || [];
  const runs = clientData.amortizationRuns || [];

  // ---- Sheet 1: Asset Register ----
  const assetSheet = workbook.addWorksheet('Asset Register', {
    properties: { tabColor: { argb: 'FF3b82f6' } },
  });

  assetSheet.columns = [
    { header: 'Asset ID', key: 'id', width: 22 },
    { header: 'Name', key: 'name', width: 32 },
    { header: 'Description', key: 'description', width: 36 },
    { header: 'GL Account', key: 'glAccountName', width: 24 },
    { header: 'Vendor', key: 'vendorName', width: 20 },
    { header: 'Original Cost', key: 'originalCost', width: 16, style: { numFmt: currencyFormat } },
    { header: 'Salvage Value', key: 'salvageValue', width: 16, style: { numFmt: currencyFormat } },
    { header: 'Acquisition Date', key: 'acquisitionDate', width: 18, style: { numFmt: dateFormat } },
    { header: 'Asset Account ID', key: 'assetAccountId', width: 18 },
    { header: 'Asset Account Name', key: 'assetAccountName', width: 24 },
    { header: 'Expense Account ID', key: 'expenseAccountId', width: 18 },
    { header: 'Expense Account Name', key: 'expenseAccountName', width: 24 },
    { header: 'Accum Account ID', key: 'accumAccountId', width: 18 },
    { header: 'Accum Account Name', key: 'accumAccountName', width: 24 },
    { header: 'QBO Account ID', key: 'qboAccountId', width: 16 },
    { header: 'Txn Key', key: 'txnKey', width: 20 },
    { header: 'Active', key: 'active', width: 8 },
    { header: 'Created At', key: 'createdAt', width: 20 },
  ];

  for (const a of assets) {
    assetSheet.addRow({
      id: a.id || '',
      name: a.name || '',
      description: a.description || '',
      glAccountName: a.glAccountName || '',
      vendorName: a.vendorName || '',
      originalCost: parseFloat(a.originalCost) || 0,
      salvageValue: parseFloat(a.salvageValue) || 0,
      acquisitionDate: a.acquisitionDate || '',
      assetAccountId: a.assetAccountId || '',
      assetAccountName: a.assetAccountName || '',
      expenseAccountId: a.expenseAccountId || '',
      expenseAccountName: a.expenseAccountName || '',
      accumAccountId: a.accumAccountId || '',
      accumAccountName: a.accumAccountName || '',
      qboAccountId: a.qboAccountId || '',
      txnKey: a.txnKey || '',
      active: a.active ? 'Yes' : 'No',
      createdAt: a.createdAt || '',
    });
  }

  styleHeaders(assetSheet);
  styleDataRows(assetSheet);

  // ---- Sheet 2: Amortization Policies ----
  const policySheet = workbook.addWorksheet('Amortization Policies', {
    properties: { tabColor: { argb: 'FF22c55e' } },
  });

  policySheet.columns = [
    { header: 'Class ID', key: 'id', width: 22 },
    { header: 'GL Account ID', key: 'glAccountId', width: 16 },
    { header: 'GL Account Name', key: 'glAccountName', width: 28 },
    { header: 'Method', key: 'method', width: 18 },
    { header: 'Useful Life (months)', key: 'usefulLifeMonths', width: 20 },
    { header: 'Declining Rate', key: 'decliningRate', width: 16, style: { numFmt: percentFormat } },
    { header: 'Salvage Value', key: 'salvageValue', width: 16, style: { numFmt: currencyFormat } },
    { header: 'Expense Account ID', key: 'expenseAccountId', width: 18 },
    { header: 'Expense Account Name', key: 'expenseAccountName', width: 24 },
    { header: 'Accum Account ID', key: 'accumAccountId', width: 18 },
    { header: 'Accum Account Name', key: 'accumAccountName', width: 24 },
    { header: 'CCA Class', key: 'ccaClass', width: 12 },
    { header: 'CCA Rate', key: 'ccaRate', width: 12 },
    { header: 'AI Reasoning', key: 'aiReasoning', width: 40 },
    { header: 'Created At', key: 'createdAt', width: 20 },
    { header: 'Updated At', key: 'updatedAt', width: 20 },
  ];

  for (const c of classes) {
    policySheet.addRow({
      id: c.id || '',
      glAccountId: c.glAccountId || '',
      glAccountName: c.glAccountName || '',
      method: c.method || 'straight-line',
      usefulLifeMonths: c.usefulLifeMonths || '',
      decliningRate: c.decliningRate || '',
      salvageValue: parseFloat(c.salvageValue) || 0,
      expenseAccountId: c.expenseAccountId || '',
      expenseAccountName: c.expenseAccountName || '',
      accumAccountId: c.accumAccountId || '',
      accumAccountName: c.accumAccountName || '',
      ccaClass: c.aiSuggestion?.ccaClass || '',
      ccaRate: c.aiSuggestion?.ccaRate || '',
      aiReasoning: c.aiSuggestion?.reasoning || '',
      createdAt: c.createdAt || '',
      updatedAt: c.updatedAt || '',
    });
  }

  styleHeaders(policySheet);
  styleDataRows(policySheet);

  // ---- Sheet 3: Amortization Schedule ----
  const schedSheet = workbook.addWorksheet('Amortization Schedule', {
    properties: { tabColor: { argb: 'FFf97316' } },
  });

  schedSheet.columns = [
    { header: 'Run Month', key: 'month', width: 14 },
    { header: 'Run Date', key: 'ranAt', width: 20 },
    { header: 'Asset ID', key: 'assetId', width: 22 },
    { header: 'Asset Name', key: 'assetName', width: 30 },
    { header: 'GL Account', key: 'glAccountName', width: 24 },
    { header: 'Expense Account', key: 'expenseAccountName', width: 24 },
    { header: 'Accum Account', key: 'accumAccountName', width: 24 },
    { header: 'Amount', key: 'amount', width: 16, style: { numFmt: currencyFormat } },
    { header: 'Run Total', key: 'runTotal', width: 16, style: { numFmt: currencyFormat } },
    { header: 'Assets in Run', key: 'assetCount', width: 14 },
    { header: 'Journal Entry ID', key: 'journalEntryId', width: 18 },
  ];

  for (const run of runs) {
    const runAssets = run.assets || [];
    for (const ra of runAssets) {
      schedSheet.addRow({
        month: run.month || '',
        ranAt: run.ranAt || '',
        assetId: ra.assetId || '',
        assetName: ra.assetName || '',
        glAccountName: ra.glAccountName || '',
        expenseAccountName: ra.expenseAccountName || '',
        accumAccountName: ra.accumAccountName || '',
        amount: parseFloat(ra.amount) || 0,
        runTotal: parseFloat(run.totalAmount) || 0,
        assetCount: run.assetCount || 0,
        journalEntryId: run.journalEntryId || '',
      });
    }
  }

  styleHeaders(schedSheet);
  styleDataRows(schedSheet);

  return workbook;
}

// ========================================
// PARSE WORKBOOK
// ========================================
function parseWorkbook(workbook) {
  const warnings = [];
  const assets = [];
  const assetClasses = [];
  const amortizationRuns = [];

  // Helper: build header-to-column-index map from row 1
  function getHeaderMap(sheet) {
    const map = {};
    const row = sheet.getRow(1);
    row.eachCell((cell, colNumber) => {
      const val = String(cell.value || '').trim();
      if (val) map[val] = colNumber;
    });
    return map;
  }

  // Helper: get cell value by header name
  function getVal(row, headerMap, headerName) {
    const col = headerMap[headerName];
    if (!col) return '';
    const cell = row.getCell(col);
    if (cell.value === null || cell.value === undefined) return '';
    // Handle formula results
    if (typeof cell.value === 'object' && cell.value.result !== undefined) return cell.value.result;
    // Handle rich text
    if (typeof cell.value === 'object' && cell.value.richText) {
      return cell.value.richText.map(r => r.text).join('');
    }
    return cell.value;
  }

  function toStr(v) { return String(v ?? '').trim(); }
  function toNum(v) {
    if (typeof v === 'number') return v;
    const s = String(v ?? '').replace(/[$,]/g, '').trim();
    return parseFloat(s) || 0;
  }

  // ---- Parse Sheet 1: Asset Register ----
  const assetSheet = workbook.getWorksheet('Asset Register');
  if (assetSheet) {
    const hm = getHeaderMap(assetSheet);
    for (let i = 2; i <= assetSheet.rowCount; i++) {
      const row = assetSheet.getRow(i);
      const name = toStr(getVal(row, hm, 'Name'));
      if (!name) continue; // skip empty rows

      assets.push({
        id: toStr(getVal(row, hm, 'Asset ID')) || 'asset_' + Date.now() + '_' + Math.random().toString(36).substring(2, 7),
        clientId: '', // will be set by the caller
        name,
        description: toStr(getVal(row, hm, 'Description')),
        glAccountName: toStr(getVal(row, hm, 'GL Account')),
        vendorName: toStr(getVal(row, hm, 'Vendor')),
        originalCost: toNum(getVal(row, hm, 'Original Cost')),
        salvageValue: toNum(getVal(row, hm, 'Salvage Value')),
        acquisitionDate: toStr(getVal(row, hm, 'Acquisition Date')),
        assetAccountId: toStr(getVal(row, hm, 'Asset Account ID')),
        assetAccountName: toStr(getVal(row, hm, 'Asset Account Name')),
        expenseAccountId: toStr(getVal(row, hm, 'Expense Account ID')),
        expenseAccountName: toStr(getVal(row, hm, 'Expense Account Name')),
        accumAccountId: toStr(getVal(row, hm, 'Accum Account ID')),
        accumAccountName: toStr(getVal(row, hm, 'Accum Account Name')),
        qboAccountId: toStr(getVal(row, hm, 'QBO Account ID')),
        txnKey: toStr(getVal(row, hm, 'Txn Key')) || null,
        active: toStr(getVal(row, hm, 'Active')).toLowerCase() !== 'no',
        createdAt: toStr(getVal(row, hm, 'Created At')) || new Date().toISOString(),
      });
    }
  } else {
    warnings.push('Sheet "Asset Register" not found — no assets imported.');
  }

  // ---- Parse Sheet 2: Amortization Policies ----
  const policySheet = workbook.getWorksheet('Amortization Policies');
  if (policySheet) {
    const hm = getHeaderMap(policySheet);
    for (let i = 2; i <= policySheet.rowCount; i++) {
      const row = policySheet.getRow(i);
      const glName = toStr(getVal(row, hm, 'GL Account Name'));
      if (!glName) continue;

      const decliningRate = toNum(getVal(row, hm, 'Declining Rate'));
      const ccaClass = toStr(getVal(row, hm, 'CCA Class'));
      const ccaRate = toStr(getVal(row, hm, 'CCA Rate'));
      const aiReasoning = toStr(getVal(row, hm, 'AI Reasoning'));
      const method = toStr(getVal(row, hm, 'Method')) || 'straight-line';

      assetClasses.push({
        id: toStr(getVal(row, hm, 'Class ID')) || 'class_' + Date.now() + '_' + Math.random().toString(36).substring(2, 7),
        glAccountId: toStr(getVal(row, hm, 'GL Account ID')),
        glAccountName: glName,
        method,
        usefulLifeMonths: parseInt(getVal(row, hm, 'Useful Life (months)')) || 36,
        decliningRate: decliningRate || null,
        salvageValue: toNum(getVal(row, hm, 'Salvage Value')),
        expenseAccountId: toStr(getVal(row, hm, 'Expense Account ID')),
        expenseAccountName: toStr(getVal(row, hm, 'Expense Account Name')),
        accumAccountId: toStr(getVal(row, hm, 'Accum Account ID')),
        accumAccountName: toStr(getVal(row, hm, 'Accum Account Name')),
        aiSuggestion: (ccaClass || aiReasoning) ? {
          ccaClass: ccaClass || null,
          ccaRate: ccaRate || null,
          reasoning: aiReasoning || '',
          method,
          suggestedAt: new Date().toISOString(),
        } : null,
        createdAt: toStr(getVal(row, hm, 'Created At')) || new Date().toISOString(),
        updatedAt: toStr(getVal(row, hm, 'Updated At')) || new Date().toISOString(),
      });
    }
  } else {
    warnings.push('Sheet "Amortization Policies" not found — no policies imported.');
  }

  // ---- Parse Sheet 3: Amortization Schedule ----
  const schedSheet = workbook.getWorksheet('Amortization Schedule');
  if (schedSheet) {
    const hm = getHeaderMap(schedSheet);
    // Group rows by Run Month to reconstruct run records
    const runMap = new Map();
    for (let i = 2; i <= schedSheet.rowCount; i++) {
      const row = schedSheet.getRow(i);
      const month = toStr(getVal(row, hm, 'Run Month'));
      if (!month) continue;

      if (!runMap.has(month)) {
        runMap.set(month, {
          month,
          ranAt: toStr(getVal(row, hm, 'Run Date')),
          totalAmount: toNum(getVal(row, hm, 'Run Total')),
          assetCount: parseInt(getVal(row, hm, 'Assets in Run')) || 0,
          journalEntryId: toStr(getVal(row, hm, 'Journal Entry ID')),
          assets: [],
        });
      }

      runMap.get(month).assets.push({
        assetId: toStr(getVal(row, hm, 'Asset ID')),
        assetName: toStr(getVal(row, hm, 'Asset Name')),
        glAccountName: toStr(getVal(row, hm, 'GL Account')),
        expenseAccountName: toStr(getVal(row, hm, 'Expense Account')),
        accumAccountName: toStr(getVal(row, hm, 'Accum Account')),
        amount: toNum(getVal(row, hm, 'Amount')),
      });
    }

    for (const run of runMap.values()) {
      amortizationRuns.push(run);
    }
  } else {
    warnings.push('Sheet "Amortization Schedule" not found — no amortization history imported.');
  }

  return { assets, assetClasses, amortizationRuns, warnings };
}

module.exports = { generateWorkbook, parseWorkbook };
