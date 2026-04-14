/**
 * Prepaid Expenses Review Service — "Claude as a reviewer"
 *
 * After we generate a prepaid-expenses Excel workbook, this module gathers the
 * saved client data, the computed schedule (per-item monthly amortization),
 * and the live QBO trial balance, then asks Claude (Sonnet 4.6) to review the
 * most recent close period for reconciliation, calculation, and reasonability
 * issues.
 *
 * Returns a structured findings object that the export endpoint caches and
 * gets rendered both as a "Review Notes" sheet inside the workbook and as a
 * traffic-light panel in the admin UI.
 */

const Anthropic = require('@anthropic-ai/sdk').default;
const qbo = require('./qbo-service');

const anthropic = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

const REVIEW_MODEL = 'claude-sonnet-4-6';

const round2 = (n) => Math.round((Number(n) || 0) * 100) / 100;

function parseMonthStr(yyyymm) {
  const [y, m] = String(yyyymm).split('-').map(Number);
  return { year: y, monthIndex: m - 1 };
}
function lastDayOfMonthDate(year, monthIndex) {
  return new Date(year, monthIndex + 1, 0);
}
function monthsBetweenInclusive(startYYYYMMDD, endYYYYMMDD) {
  if (!startYYYYMMDD || !endYYYYMMDD) return 0;
  const s = new Date(startYYYYMMDD);
  const e = new Date(endYYYYMMDD);
  if (isNaN(s) || isNaN(e) || e < s) return 0;
  return (e.getFullYear() - s.getFullYear()) * 12 + (e.getMonth() - s.getMonth()) + 1;
}
function monthsRecognizedThroughEnd(item, asOfMonth) {
  // How many months have been recognized from startDate through end of asOfMonth (inclusive)
  const { year: ty, monthIndex: tm } = parseMonthStr(asOfMonth);
  const s = new Date(item.startDate + 'T00:00:00');
  if (isNaN(s)) return 0;
  const elapsed = (ty - s.getFullYear()) * 12 + (tm - s.getMonth()) + 1;
  const total = monthsBetweenInclusive(item.startDate, item.endDate);
  return Math.max(0, Math.min(elapsed, total));
}

function flattenTrialBalance(tb) {
  const map = new Map();
  function walk(rows) {
    if (!rows) return;
    for (const row of rows) {
      if (row.ColData) {
        const acctName = row.ColData[0]?.value || '';
        const debit = parseFloat(String(row.ColData[1]?.value || '').replace(/[$,\s]/g, '')) || 0;
        const credit = parseFloat(String(row.ColData[2]?.value || '').replace(/[$,\s]/g, '')) || 0;
        if (acctName) map.set(acctName, debit - credit);
      }
      if (row.Rows?.Row) walk(row.Rows.Row);
      if (row.Rows && Array.isArray(row.Rows)) walk(row.Rows);
    }
  }
  walk(tb?.Rows?.Row);
  return map;
}

/**
 * Build the per-item schedule snapshot as of asOfMonth, using the item list and
 * amortizationRuns history. Closing balance = openingBalance - (recognized through asOfMonth).
 */
function buildItemSnapshots(clientData, asOfMonth) {
  const runs = clientData.amortizationRuns || [];
  const items = clientData.items || [];

  // Build per-item recognized-through-asOfMonth total from the posted runs
  // (sum of line amounts for this item across all runs whose month <= asOfMonth)
  const recognizedByItem = new Map();
  for (const run of runs) {
    if ((run.month || '') > asOfMonth) continue;
    for (const line of run.lines || []) {
      const key = line.itemId;
      if (!key) continue;
      recognizedByItem.set(key, round2((recognizedByItem.get(key) || 0) + (Number(line.amount) || 0)));
    }
  }

  const snapshots = items.map(item => {
    const totalMonths = monthsBetweenInclusive(item.startDate, item.endDate);
    const amortizableAmount = round2(Number(item.openingBalance ?? item.totalAmount) || 0);
    const monthlyAmount = totalMonths > 0 ? round2(amortizableAmount / totalMonths) : 0;
    const monthsThrough = monthsRecognizedThroughEnd(item, asOfMonth);
    const expectedRecognized = totalMonths > 0
      ? round2((amortizableAmount / totalMonths) * monthsThrough)
      : 0;
    const actualRecognized = round2(recognizedByItem.get(item.id) || 0);
    const closingBalance = round2(amortizableAmount - actualRecognized);
    const variance = round2(actualRecognized - expectedRecognized);
    return {
      id: item.id,
      vendor: item.vendor || '',
      description: item.description || '',
      expenseAccountName: item.expenseAccountName || '',
      startDate: item.startDate || '',
      endDate: item.endDate || '',
      totalAmount: round2(Number(item.totalAmount) || 0),
      openingBalance: amortizableAmount,
      totalMonths,
      monthlyAmount,
      monthsThrough,
      expectedRecognized,
      actualRecognized,
      recognizedVariance: variance,
      closingBalance,
    };
  });

  const totals = snapshots.reduce((acc, s) => {
    acc.opening = round2(acc.opening + s.openingBalance);
    acc.recognized = round2(acc.recognized + s.actualRecognized);
    acc.closing = round2(acc.closing + s.closingBalance);
    return acc;
  }, { opening: 0, recognized: 0, closing: 0 });

  return { snapshots, totals };
}

async function buildReviewContext(clientId, clientData, asOfMonth) {
  const { year, monthIndex } = parseMonthStr(asOfMonth);
  const asOfDate = lastDayOfMonthDate(year, monthIndex).toISOString().split('T')[0];

  const { snapshots, totals } = buildItemSnapshots(clientData, asOfMonth);

  // Fetch QBO trial balance for Prepaid account comparison
  let qboTrialBalance = {};
  let qboError = null;
  if (qbo.isConnected(clientId)) {
    try {
      const tb = await qbo.getTrialBalance(asOfDate, asOfDate, clientId);
      const tbMap = flattenTrialBalance(tb);
      for (const [name, balance] of tbMap) {
        qboTrialBalance[name] = round2(balance);
      }
      if (tbMap.size === 0) {
        qboError = 'QBO trial balance returned no rows for the as-of date';
      }
    } catch (e) {
      qboError = `Failed to fetch QBO trial balance: ${e.message}`;
    }
  } else {
    qboError = 'QBO not connected for this client';
  }

  return {
    asOfMonth,
    asOfDate,
    qboError,
    prepaidAccount: clientData.prepaidAccount || null,
    items: snapshots,
    totals,
    qboTrialBalance,
    amortizationRunCount: (clientData.amortizationRuns || []).length,
  };
}

const REVIEW_SYSTEM_PROMPT = `You are a senior accountant reviewing a prepaid-expenses schedule prepared by a junior bookkeeper. Your job is to spot reconciliation differences, calculation errors, and unreasonable inputs that a partner would catch in a file review.

You will be given (a) per-item prepaid records with expected vs actual recognized-to-date amounts, (b) totals (opening, recognized, closing), (c) the Prepaid GL account configured for this client, and (d) a QBO trial balance slice for the same as-of date.

Run these checks ONLY for the most recent period. Do not flag historical period-by-period differences.

RECONCILIATION CHECKS:
- Does the sum of item closingBalance equal the QBO trial balance for the Prepaid account? (within $0.01) QBO may return fully-qualified names like "Current Assets:Prepaid Expenses" — match by leaf name (the part after the last colon).
- Flag any item where closingBalance < 0.

CALCULATION CHECKS:
- For each item, compare actualRecognized to expectedRecognized. |variance| > $1 is a finding.
- Flag items where monthsThrough > totalMonths but closingBalance > 0 (should be fully amortized).
- Flag items where totalMonths <= 0 (end date before start date) or monthlyAmount is 0 with a non-zero opening balance.

REASONABILITY CHECKS:
- Flag items missing an expenseAccountName.
- Flag items where the amortization period is unreasonably long (> 60 months) or short (< 1 month) given the vendor/description.
- Flag items where startDate is in the future relative to asOfMonth.

OUTPUT FORMAT — RESPOND WITH JSON ONLY, NO PROSE OR MARKDOWN:
{
  "status": "clean" | "warnings" | "errors",
  "summary": "1-2 sentence overall assessment",
  "findings": [
    {
      "severity": "error" | "warning" | "info",
      "category": "reconciliation" | "calculation" | "reasonability",
      "assetId": "<item id or null>",
      "assetName": "<vendor + description or null>",
      "message": "<specific, actionable description of the issue>",
      "expected": <number or null>,
      "actual": <number or null>
    }
  ]
}

Status rules:
- "errors" if any finding has severity "error"
- "warnings" if any finding has severity "warning" but no errors
- "clean" if no findings

Be specific. "Prepaid balance mismatch: schedule closing $2,400.00 vs QBO $1,800.00 (diff $600.00)" — not "balance may be wrong".`;

async function reviewPrepaid(clientId, clientData, asOfMonth) {
  if (!asOfMonth) {
    return {
      status: 'skipped',
      summary: 'No close period specified.',
      findings: [],
      generatedAt: new Date().toISOString(),
      asOfMonth: null,
    };
  }
  if (!(clientData.items || []).length) {
    return {
      status: 'skipped',
      summary: 'No prepaid items to review.',
      findings: [],
      generatedAt: new Date().toISOString(),
      asOfMonth,
    };
  }

  const ctx = await buildReviewContext(clientId, clientData, asOfMonth);

  const userMessage = `Review the prepaid expenses schedule as of ${ctx.asOfMonth}.

${ctx.qboError ? `NOTE: ${ctx.qboError}. Reconciliation against QBO is not possible — focus on calculation and reasonability checks.\n` : ''}
=== PREPAID GL ACCOUNT ===
${JSON.stringify(ctx.prepaidAccount, null, 2)}

=== TOTALS ===
${JSON.stringify(ctx.totals, null, 2)}

=== ITEMS ===
${JSON.stringify(ctx.items, null, 2)}

=== QBO TRIAL BALANCE (as of ${ctx.asOfDate}) ===
${JSON.stringify(ctx.qboTrialBalance, null, 2)}

Respond with JSON only.`;

  let parsed;
  try {
    const response = await anthropic.messages.create({
      model: REVIEW_MODEL,
      max_tokens: 4000,
      system: REVIEW_SYSTEM_PROMPT,
      messages: [{ role: 'user', content: userMessage }],
    });
    const text = response.content?.[0]?.text || '';
    const cleaned = text.replace(/^```json\s*/i, '').replace(/^```\s*/i, '').replace(/```\s*$/i, '').trim();
    const jsonMatch = cleaned.match(/\{[\s\S]*\}/);
    if (!jsonMatch) throw new Error('Claude response did not contain JSON');
    parsed = JSON.parse(jsonMatch[0]);
  } catch (e) {
    console.error('[prepaid-review] Claude call failed:', e.message);
    return {
      status: 'error',
      summary: `Review could not be completed: ${e.message}`,
      findings: [],
      generatedAt: new Date().toISOString(),
      asOfMonth: ctx.asOfMonth,
      reviewError: e.message,
    };
  }

  const findings = Array.isArray(parsed.findings) ? parsed.findings : [];
  const normalized = findings.map(f => ({
    severity: ['error', 'warning', 'info'].includes(f.severity) ? f.severity : 'info',
    category: f.category || 'general',
    assetId: f.assetId || null,
    assetName: f.assetName || null,
    message: f.message || '',
    expected: f.expected ?? null,
    actual: f.actual ?? null,
  }));

  let status = 'clean';
  if (normalized.some(f => f.severity === 'error')) status = 'errors';
  else if (normalized.some(f => f.severity === 'warning')) status = 'warnings';

  return {
    status,
    summary: parsed.summary || (status === 'clean' ? 'All checks passed.' : `${normalized.length} finding(s).`),
    findings: normalized,
    generatedAt: new Date().toISOString(),
    asOfMonth: ctx.asOfMonth,
  };
}

module.exports = { reviewPrepaid, buildItemSnapshots };
