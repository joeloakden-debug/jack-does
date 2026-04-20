/**
 * Excel Review Service — "Claude as a reviewer"
 *
 * After we generate a fixed-asset Excel workbook, this module gathers the saved
 * client data, the continuity schedule we just produced, and live QBO trial-balance
 * + source-transaction data, then asks Claude (Sonnet 4.6) to review the most
 * recent period for reconciliation, calculation, and reasonability issues.
 *
 * Returns a structured findings object that the export endpoint caches and that
 * gets rendered both as a "Review Notes" sheet inside the workbook and as a
 * traffic-light panel in the admin UI.
 *
 * Scope (per product decision): only review the most recent period. Reasonability
 * tests use life-to-date amortization, not period-by-period replays of history.
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
function monthsBetween(startYYYYMMDD, endYYYYMM) {
  if (!startYYYYMMDD || !endYYYYMM) return 0;
  const start = new Date(startYYYYMMDD);
  if (isNaN(start)) return 0;
  const { year, monthIndex } = parseMonthStr(endYYYYMM);
  const end = lastDayOfMonthDate(year, monthIndex);
  if (end < start) return 0;
  // Inclusive count of months from acquisition month through asOfMonth
  return (end.getFullYear() - start.getFullYear()) * 12 + (end.getMonth() - start.getMonth()) + 1;
}

/**
 * Compute life-to-date amortization for an asset by summing all per-asset amounts
 * across the entire amortizationRuns history. We match by id, then by name.
 */
function ltdAmortFor(asset, runs) {
  let total = 0;
  for (const run of runs || []) {
    for (const a of run.assets || []) {
      if ((a.assetId && a.assetId === asset.id) || (a.name && a.name === asset.name) || (a.assetName && a.assetName === asset.name)) {
        total += parseFloat(a.amount) || 0;
      }
    }
  }
  return round2(total);
}

/**
 * Compute the EXPECTED life-to-date amortization for an asset using its policy.
 * Straight-line: (cost - salvage) / usefulLife * monthsElapsed (capped).
 * Declining-balance: iterate month by month at decliningRate/12 (closed form is
 * messy with salvage, so iterate — fine for review purposes).
 */
function expectedLtdAmort(asset, policy, asOfMonth) {
  const cost = parseFloat(asset.originalCost) || 0;
  const salvage = parseFloat(asset.salvageValue || policy.salvageValue || 0) || 0;
  const monthsElapsed = monthsBetween(asset.acquisitionDate, asOfMonth);
  if (monthsElapsed <= 0) return 0;

  const method = (policy.method || 'straight-line').toLowerCase();
  if (method === 'declining-balance' && policy.decliningRate) {
    let bookValue = cost;
    const monthlyRate = policy.decliningRate / 12;
    let total = 0;
    for (let i = 0; i < monthsElapsed; i++) {
      if (bookValue <= salvage) break;
      const monthly = bookValue * monthlyRate;
      const cap = Math.max(0, bookValue - salvage);
      const m = Math.min(monthly, cap);
      total += m;
      bookValue -= m;
    }
    return round2(total);
  }
  // Straight-line
  const life = policy.usefulLifeMonths || asset.usefulLifeMonths || 60;
  const usable = Math.min(monthsElapsed, life);
  return round2(((cost - salvage) / life) * usable);
}

/**
 * Walk a QBO trial-balance report into a name -> signed balance map.
 * Same logic as reconcileScheduleToQBO in server.js, kept local so this module is standalone.
 */
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
 * Build the structured input that gets handed to Claude. Keep this small —
 * one row per asset, totals only for the most recent period, plus a slim
 * QBO TB slice and a list of source-transaction acquisitions for completeness checks.
 */
async function buildReviewContext(clientId, clientData, continuitySchedule, getAssetPolicy) {
  const asOfMonth = continuitySchedule.asOfMonth;
  const { year, monthIndex } = parseMonthStr(asOfMonth);
  const asOfDate = lastDayOfMonthDate(year, monthIndex).toISOString().split('T')[0];

  // ---- Per-asset rows ----
  const assets = [];
  for (const a of (clientData.assets || []).filter(x => x.active)) {
    const policy = getAssetPolicy(a, clientData.assetClasses) || {};
    const ltdActual = ltdAmortFor(a, clientData.amortizationRuns);
    const ltdExpected = expectedLtdAmort(a, policy, asOfMonth);
    const variance = round2(ltdActual - ltdExpected);
    assets.push({
      id: a.id,
      name: a.name,
      description: a.description || '',
      glAccountName: a.glAccountName || '',
      vendorName: a.vendorName || '',
      originalCost: round2(a.originalCost),
      salvageValue: round2(a.salvageValue || 0),
      acquisitionDate: a.acquisitionDate || '',
      method: policy.method || '',
      usefulLifeMonths: policy.usefulLifeMonths || a.usefulLifeMonths || null,
      decliningRate: policy.decliningRate || null,
      expenseAccountName: policy.expenseAccountName || a.expenseAccountName || '',
      accumAccountName: policy.accumAccountName || a.accumAccountName || '',
      monthsElapsed: monthsBetween(a.acquisitionDate, asOfMonth),
      ltdAmortActual: ltdActual,
      ltdAmortExpected: ltdExpected,
      ltdAmortVariance: variance,
      netBookValue: round2((parseFloat(a.originalCost) || 0) - ltdActual),
      txnKey: a.txnKey || null,
    });
  }

  // ---- Continuity schedule subtotals (most recent period only) ----
  const scheduleSummary = {
    asOfMonth,
    fiscalYearStart: continuitySchedule.fiscalYearStart,
    fiscalYearEnd: continuitySchedule.fiscalYearEnd,
    total: continuitySchedule.total,
    glAccounts: (continuitySchedule.glAccounts || []).map(g => ({
      glAccountName: g.glAccountName,
      assetCount: g.assets.length,
      subtotal: g.subtotal,
    })),
  };

  // ---- QBO trial balance ----
  // Return the WHOLE TB (it's small) rather than pre-filtering against the schedule's
  // account names — QBO often returns fully-qualified names like "Fixed Assets:Computer
  // hardware" while the schedule stores just the leaf name, so an exact-match filter
  // would silently drop everything. Let Claude do the matching with its own judgment.
  let qboTrialBalance = {};
  let qboError = null;
  if (qbo.isConnected(clientId)) {
    try {
      const tb = await qbo.getTrialBalance(asOfDate, asOfDate, clientId);
      const tbMap = flattenTrialBalance(tb);
      console.log(`[review] QBO TB rows: ${tbMap.size}`);
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

  // ---- QBO source acquisitions (for completeness checks) ----
  let qboAcquisitions = [];
  if (qbo.isConnected(clientId)) {
    try {
      // Build cost-account id set from active assets so we ask QBO for the right scope
      const costAcctIds = new Set(
        (clientData.assets || [])
          .filter(a => a.active && a.qboAccountId)
          .map(a => String(a.qboAccountId))
      );
      if (costAcctIds.size > 0) {
        const acqs = await qbo.getFixedAssetAcquisitions(clientId, costAcctIds);
        qboAcquisitions = acqs.map(a => ({
          txnType: a.txnType,
          txnId: a.txnId,
          txnDate: a.txnDate,
          accountName: a.accountName || '',
          amount: round2(a.amount),
          description: a.description || '',
          vendorName: a.vendorName || '',
          txnKey: `${a.txnType}-${a.txnId}-${a.accountId}-${a.amount}`,
        }));
      }
    } catch (e) {
      // Non-fatal — review can still happen without source-txn list
      console.error('[review] failed to load QBO acquisitions:', e.message);
    }
  }

  return {
    asOfMonth,
    asOfDate,
    qboError,
    assets,
    scheduleSummary,
    qboTrialBalance,
    qboAcquisitions,
  };
}

const REVIEW_SYSTEM_PROMPT = `You are a senior accountant reviewing a fixed-asset schedule prepared by a junior bookkeeper. Your job is to spot reconciliation differences, calculation errors, and unreasonable inputs that a partner would catch in a file review.

You will be given (a) the schedule's most-recent-period totals, (b) per-asset records with computed life-to-date amortization (actual vs expected based on policy), (c) a slice of the QBO trial balance for the same as-of date, and (d) the source acquisition transactions pulled directly from QBO.

Run these checks ONLY for the most recent period. Do not flag historical period-by-period differences.

RECONCILIATION CHECKS:
- For each GL account in the schedule, does the schedule cost subtotal equal the QBO trial balance for that cost account? (within $0.01)
- For each accum account named on assets, does the schedule's closing accum equal the absolute value of the QBO trial balance for that accum account?
- Are there any cost-account acquisitions in QBO (qboAcquisitions) that are NOT represented as an active asset (match by txnKey)? Flag any orphans.
- Are there any active assets whose txnKey is not present in QBO acquisitions? Flag as "asset has no QBO source txn".

CALCULATION CHECKS:
- For each asset, compare ltdAmortActual to ltdAmortExpected. Variance > $1 OR > 5% of cost is a finding.
- Flag any asset where closing net book value < 0.
- Flag any asset where (originalCost - salvageValue) <= 0.
- Flag any asset with method "straight-line" but missing usefulLifeMonths, or "declining-balance" but missing decliningRate.

REASONABILITY CHECKS:
- Flag useful lives that look implausible (< 12 months or > 480 months) given the asset type implied by the name/description.
- Flag salvage values >= original cost.
- Flag assets with $0 LTD amortization despite monthsElapsed > 0 and a defined policy.
- Flag missing accum or expense account names on active assets.

OUTPUT FORMAT — RESPOND WITH JSON ONLY, NO PROSE OR MARKDOWN:
{
  "status": "clean" | "warnings" | "errors",
  "summary": "1-2 sentence overall assessment",
  "findings": [
    {
      "severity": "error" | "warning" | "info",
      "category": "reconciliation" | "calculation" | "reasonability" | "completeness",
      "assetId": "<asset id or null>",
      "assetName": "<asset name or null>",
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

Be specific. "Cost mismatch on Computer Hardware: schedule $2,197.60 vs QBO $2,053.83 (diff $143.77)" — not "cost may be wrong".`;

async function reviewWorkbook(clientId, clientData, continuitySchedule, getAssetPolicy) {
  if (!continuitySchedule || !continuitySchedule.glAccounts) {
    return {
      status: 'skipped',
      summary: 'No continuity schedule available to review.',
      findings: [],
      generatedAt: new Date().toISOString(),
      asOfMonth: null,
    };
  }

  const ctx = await buildReviewContext(clientId, clientData, continuitySchedule, getAssetPolicy);

  // If we couldn't load QBO data at all, do an offline-only review (limited but useful)
  const userMessage = `Review the most recent period (${ctx.asOfMonth}) of this fixed-asset schedule.

${ctx.qboError ? `NOTE: ${ctx.qboError}. Reconciliation checks against QBO are not possible — focus on calculation and reasonability checks.\n` : ''}
=== SCHEDULE SUMMARY (most recent period) ===
${JSON.stringify(ctx.scheduleSummary, null, 2)}

=== ASSETS (life-to-date) ===
${JSON.stringify(ctx.assets, null, 2)}

=== QBO TRIAL BALANCE (as of ${ctx.asOfDate}) ===
The keys here are the account names exactly as QBO returned them. They may use
fully-qualified names like "Fixed Assets:Computer hardware" while the schedule
uses the leaf name "Computer hardware" — match by leaf name (the part after the
last colon) when comparing to the schedule's glAccountName / accumAccountName /
expenseAccountName. Values are signed (debit positive, credit negative); accum
balances will appear as negative numbers and should be compared by absolute value.
${JSON.stringify(ctx.qboTrialBalance, null, 2)}

=== QBO SOURCE ACQUISITIONS ===
${JSON.stringify(ctx.qboAcquisitions, null, 2)}

Respond with JSON only.`;

  let parsed;
  try {
    // Cache the static review system prompt (Anthropic prompt caching). Cuts
    // per-call cost/latency on repeated reviews with the same instructions.
    const response = await anthropic.messages.create({
      model: REVIEW_MODEL,
      max_tokens: 4000,
      system: [{
        type: 'text',
        text: REVIEW_SYSTEM_PROMPT,
        cache_control: { type: 'ephemeral' },
      }],
      messages: [{ role: 'user', content: userMessage }],
    });
    const text = response.content?.[0]?.text || '';
    // Strip any code fences if Claude added them despite instructions
    const cleaned = text.replace(/^```json\s*/i, '').replace(/^```\s*/i, '').replace(/```\s*$/i, '').trim();
    const jsonMatch = cleaned.match(/\{[\s\S]*\}/);
    if (!jsonMatch) throw new Error('Claude response did not contain JSON');
    parsed = JSON.parse(jsonMatch[0]);
  } catch (e) {
    console.error('[review] Claude call failed:', e.message);
    return {
      status: 'error',
      summary: `Review could not be completed: ${e.message}`,
      findings: [],
      generatedAt: new Date().toISOString(),
      asOfMonth: ctx.asOfMonth,
      reviewError: e.message,
    };
  }

  // Normalize shape
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

  // Recompute status from severities so we can trust it downstream
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

module.exports = { reviewWorkbook };
