#!/usr/bin/env node
/**
 * Mock Financial Data Generator — 3 Test Sets
 *
 * SET A — CLEAN: Pinnacle Distribution Co., LLC
 *   Well-structured, standardized labels, complete data, easy for RAG/Gemini.
 *
 * SET B — MESSY: Cascade Building Materials Inc.
 *   Inconsistent column headers, mixed period labels, some rounding, narrative
 *   text mixed with numbers, some cells missing, alternate terminology.
 *
 * SET C — VERY MESSY: Heartland Food Services LLC
 *   Raw QB export dump, inconsistent account names, periods in different tabs,
 *   numbers in text paragraphs, non-standard formatting, partial data,
 *   requires Gemini to infer heavily. Maximum stress test.
 *
 * Usage: node scripts/generate-mock-data.js
 */
const XLSX = require('xlsx')
const path = require('path')
const fs = require('fs')

const OUT_DIR = path.join(__dirname, '..', 'mock-data')
if (!fs.existsSync(OUT_DIR)) fs.mkdirSync(OUT_DIR, { recursive: true })

// ─────────────────────────────────────────────────────────────────────────────
// SHARED HELPERS
// ─────────────────────────────────────────────────────────────────────────────
const $ = (n) => Math.round(n)
const pct = (n, d = 1) => (n * 100).toFixed(d) + '%'
const fmt = (n) => n.toLocaleString('en-US', { style: 'currency', currency: 'USD', maximumFractionDigits: 0 })
const wb_write = (wb, name) => {
  const fp = path.join(OUT_DIR, name)
  XLSX.writeFile(wb, fp)
  console.log('✓', fp)
}
const new_wb = () => XLSX.utils.book_new()
const add_sheet = (wb, rows, name, colWidths) => {
  const ws = XLSX.utils.aoa_to_sheet(rows)
  if (colWidths) ws['!cols'] = colWidths.map((w) => ({ wch: w }))
  XLSX.utils.book_append_sheet(wb, ws, name)
}

// ─────────────────────────────────────────────────────────────────────────────
// ══════════════════════════════════════════════════════════════════════════════
// SET A: CLEAN — Pinnacle Distribution Co., LLC
// ══════════════════════════════════════════════════════════════════════════════
// ─────────────────────────────────────────────────────────────────────────────
function buildCleanSet() {
  console.log('\n── SET A: CLEAN (Pinnacle Distribution Co., LLC) ──')

  const PERIODS = ['FY2020', 'FY2021', 'FY2022', 'TTM Oct22-Sep23']

  // Core financials
  const rev   = [21_847_203, 27_041_106, 32_607_418, 35_724_507]
  const cogs  = [13_501_905, 16_709_123, 20_124_249, 22_057_633]
  const gp    = rev.map((r, i) => r - cogs[i])

  const opex_detail = {
    sales_marketing:       [1_215_000, 1_489_000, 1_834_000, 2_045_000],
    salaries_payroll:      [2_187_000, 2_734_000, 3_298_000, 3_612_000],
    benefits_401k:         [  234_000,   287_000,   346_000,   379_000],
    rent_occupancy:        [  312_000,   360_000,   440_000,   480_000],
    professional_services: [   87_654,   213_456,   197_234,   187_654],
    technology_software:   [  123_456,   145_678,   167_890,   178_943],
    insurance:             [   98_765,   112_345,   128_901,   134_567],
    payment_processing:    [  218_472,   270_411,   326_074,   357_245],
    travel_entertainment:  [   43_210,    52_345,    67_890,    72_345],
    office_admin:          [   98_765,   112_345,   128_901,   134_567],
  }
  const opex = PERIODS.map((_, i) => Object.values(opex_detail).reduce((s, v) => s + v[i], 0))
  const ebitda = gp.map((g, i) => g - opex[i])
  const da     = [312_000, 298_000, 287_500, 274_000]
  const int_ex = [187_234, 156_882, 134_221, 121_447]
  const tax    = [ 50_568,  67_777,  85_198,  95_310]
  const ni     = ebitda.map((e, i) => e - da[i] - int_ex[i] - tax[i])

  const mgmt_adj = {
    charitable:     [45_000, 52_000, 61_000, 65_000],
    owner_tax_legal:[18_500, 22_000, 19_500, 18_000],
    personal_exp:   [23_400, 27_800, 31_200, 28_500],
  }
  const total_mgmt = PERIODS.map((_, i) => Object.values(mgmt_adj).reduce((s, v) => s + v[i], 0))
  const mgmt_ebitda = ebitda.map((e, i) => e + total_mgmt[i])

  const dilig_adj = {
    professional_fees_onetime: [      0, 125_000,  95_000,       0],
    executive_search:          [      0,       0,  45_000,  22_000],
    facility_relocation:       [      0,       0,  38_500,       0],
    severance:                 [      0,  32_000,       0,       0],
    normalize_owner_comp:      [-175_000,-175_000,-175_000,-175_000],
  }
  const total_dilig = PERIODS.map((_, i) => Object.values(dilig_adj).reduce((s, v) => s + v[i], 0))
  const adj_ebitda = mgmt_ebitda.map((e, i) => e + total_dilig[i])

  // Balance sheet
  const bs = {
    cash:      [1_234_567, 1_876_543, 2_345_678, 2_789_432],
    ar_gross:  [2_187_394, 2_834_671, 3_456_789, 3_891_245],
    ar_allow:  [   43_748,    56_693,    69_136,    77_825],
    inventory: [3_456_821, 4_123_456, 5_234_567, 5_678_234],
    prepaid:   [  287_643,   334_782,   423_891,   489_234],
    other_ca:  [   45_123,    54_234,    67_890,    74_321],
    ppe_gross: [1_867_234, 1_987_234, 2_087_234, 2_137_234],
    accum_dep: [  632_343,   930_343, 1_217_843, 1_491_843],
    ap:        [1_876_432, 2_234_891, 2_678_234, 2_987_654],
    cc_pay:    [   43_219,    56_789,    72_345,    81_234],
    accrued_payroll: [234_567, 289_342, 345_678, 378_943],
    accrued_other:   [145_789, 187_654, 234_567, 267_890],
    sales_tax: [  87_654,   108_234,   132_456,   145_678],
    deferred:  [  23_456,    28_901,    34_567,    38_901],
    curr_debt: [ 400_000,   400_000,   400_000,   400_000],
    lt_debt:   [1_800_000, 1_400_000, 1_000_000, 600_000],
    common_stk:[   1_000,     1_000,     1_000,     1_000],
    apic:      [1_249_000, 1_249_000, 1_249_000, 1_249_000],
    re:        [  716_976, 1_602_194, 2_743_645, 4_748_368],
  }
  const net_ar = bs.ar_gross.map((a, i) => a - bs.ar_allow[i])
  const tot_inv = bs.inventory
  const tot_ca = bs.cash.map((c, i) => c + net_ar[i] + tot_inv[i] + bs.prepaid[i] + bs.other_ca[i])
  const net_fa = bs.ppe_gross.map((g, i) => g - bs.accum_dep[i])
  const tot_assets = tot_ca.map((c, i) => c + net_fa[i])
  const tot_cl = PERIODS.map((_, i) => bs.ap[i]+bs.cc_pay[i]+bs.accrued_payroll[i]+bs.accrued_other[i]+bs.sales_tax[i]+bs.deferred[i]+bs.curr_debt[i])
  const tot_liab = tot_cl.map((c, i) => c + bs.lt_debt[i])
  const tot_eq = tot_assets.map((a, i) => a - tot_liab[i])

  // Vendors
  const vendors = {
    'Global Manufacturing Inc.':    [3_450_000, 4_312_000, 5_234_000, 5_789_000],
    'Pacific Logistics Group':       [2_756_000, 3_456_000, 4_187_000, 4_623_000],
    'Midwest Distribution Co.':      [1_987_000, 2_456_000, 2_987_000, 3_312_000],
    'TechComponents LLC':            [1_543_000, 1_923_000, 2_345_000, 2_598_000],
    'Atlantic Packaging Solutions':  [  876_000, 1_087_000, 1_312_000, 1_456_000],
    'Sterling Materials Group':      [  654_000,   812_000,   987_000, 1_087_000],
    'Express Freight Services':      [  312_456,   389_234,   452_187,   498_234],
    'Quality Control Labs':          [   98_765,   112_345,   128_901,   134_567],
  }
  const vendor_sub = PERIODS.map((_, i) => Object.values(vendors).reduce((s, v) => s + v[i], 0))
  const vendor_other = vendor_sub.map((v, i) => cogs[i] - v)

  // Monthly TTM
  const months = ['Oct-22','Nov-22','Dec-22','Jan-23','Feb-23','Mar-23','Apr-23','May-23','Jun-23','Jul-23','Aug-23','Sep-23']
  const m_rev  = [2_934_000,3_012_000,3_445_000,2_876_000,2_987_000,3_123_000,2_934_000,3_045_000,3_098_000,2_876_000,2_934_000,3_460_507]
  const m_cogs = m_rev.map(r => Math.round(r * 0.617))
  const m_opex = m_rev.map(r => Math.round(r * 0.245))

  // Bank deposits
  const bank_d = {
    FY2020: [1_612_000,1_589_000,1_734_000,1_823_000,1_789_000,1_867_000,1_912_000,1_843_000,1_978_000,1_934_000,1_856_000,2_102_000],
    FY2021: [2_012_000,1_987_000,2_134_000,2_287_000,2_212_000,2_345_000,2_412_000,2_287_000,2_456_000,2_389_000,2_312_000,2_567_000],
    FY2022: [2_456_000,2_389_000,2_634_000,2_756_000,2_678_000,2_823_000,2_912_000,2_756_000,2_934_000,2_867_000,2_756_000,3_234_000],
  }

  // Product lines
  const prod = {
    'Premium Series A - Commercial Safety PPE': [4_215_000,5_187_000,6_234_000,6_812_000],
    'Standard Line B - Standard MRO Supplies':  [3_456_000,4_287_000,5_187_000,5_678_000],
    'Pro Edition C - Industrial Tooling':       [5_234_000,6_478_000,7_812_000,8_543_000],
    'Value Pack D - Bulk/Value Bundle':         [2_187_000,2_712_000,3_287_000,3_598_000],
    'Elite Model E - Enterprise PPE Kits':      [1_876_000,2_312_000,2_812_000,3_078_000],
    'Classic Series F - Classic Safety':        [1_654_000,2_045_000,2_478_000,2_712_000],
    'Compact Unit G - Compact Tools':           [1_234_000,1_534_000,1_867_000,2_045_000],
    'Advanced Kit H - Advanced Industrial':     [  987_000,1_234_000,1_512_000,1_656_000],
    'Starter Set I - Entry Level':              [  789_000,  978_000,1_187_000,1_298_000],
    'Professional J - Specialty':               [  654_000,  812_000,  987_000,1_082_000],
    'All Other Products':                       [  456_430,  543_820,  644_650,  706_440],
  }

  // ── File 1: Financial Statements ─────────────────────────────────────────
  const wb1 = new_wb()
  add_sheet(wb1, [
    ['PINNACLE DISTRIBUTION CO., LLC'],
    ['Income Statement — Quality of Earnings Analysis'],
    ['Periods: FY2020, FY2021, FY2022, TTM October 2022 – September 2023'],
    ['Basis: Accrual / US GAAP / Unaudited'],
    [],
    ['($ in USD)', ...PERIODS],
    [],
    ['REVENUE'],
    ['Gross Product Sales',         ...rev.map((r,i) => r + Math.round(r*0.021))],
    ['Less: Returns & Allowances',  ...rev.map(r => -Math.round(r*0.0203))],
    ['Less: Promotional Discounts', ...rev.map(r => -Math.round(r*0.0102))],
    ['Net Product Sales',           ...rev.map(r => Math.round(r*0.9695))],
    ['Shipping Revenue',            ...rev.map(r => Math.round(r*0.0143))],
    ['Other Revenue',               ...rev.map(r => Math.round(r*0.00059))],
    ['Total Net Revenue',           ...rev],
    [],
    ['COST OF GOODS SOLD'],
    ['Product Cost of Goods Sold',    ...cogs.map(c => Math.round(c*0.906))],
    ['Warehousing & Fulfillment',     ...cogs.map(c => Math.round(c*0.0643))],
    ['Shipping & Freight Out',        ...rev.map(r => Math.round(r*0.0143))],
    ['Inventory Adjustments',         ...cogs.map(c => Math.round(c*0.0065))],
    ['Total Cost of Goods Sold',      ...cogs],
    [],
    ['GROSS PROFIT',                  ...gp],
    ['Gross Margin %',                ...gp.map((g,i) => pct(g/rev[i]))],
    [],
    ['OPERATING EXPENSES'],
    ['Sales & Marketing',             ...opex_detail.sales_marketing],
    ['Salaries, Wages & Payroll Tax', ...opex_detail.salaries_payroll],
    ['Employee Benefits & 401k',      ...opex_detail.benefits_401k],
    ['Rent & Occupancy',              ...opex_detail.rent_occupancy],
    ['Professional Services',         ...opex_detail.professional_services],
    ['Technology & Software',         ...opex_detail.technology_software],
    ['Insurance',                     ...opex_detail.insurance],
    ['Payment Processing Fees',       ...opex_detail.payment_processing],
    ['Travel & Entertainment',        ...opex_detail.travel_entertainment],
    ['Office & General Admin',        ...opex_detail.office_admin],
    ['Total Operating Expenses',      ...opex],
    [],
    ['Operating Income (EBITDA)',     ...ebitda],
    ['EBITDA Margin %',               ...ebitda.map((e,i) => pct(e/rev[i]))],
    [],
    ['Interest Expense',              ...int_ex],
    ['Depreciation & Amortization',   ...da],
    ['Income Before Tax',             ...ebitda.map((e,i) => e - da[i] - int_ex[i])],
    ['Tax Provision',                 ...tax],
    ['Net Income',                    ...ni],
    ['Net Margin %',                  ...ni.map((n,i) => pct(n/rev[i]))],
    [],
    ['NOTES:'],
    ['Revenue recognition: ASC 606, recognized upon shipment. 30-day return policy.'],
    ['Owner Compensation: CEO/Owner Michael Torres earned $475,000 in FY2022 (salary + distributions).'],
    ['Professional Services FY2021: Includes $125,000 non-recurring patent dispute legal fees (Acme Industrial, settled Oct 2021).'],
    ['Professional Services FY2022: Includes $95,000 non-recurring legal fees (patent compliance follow-up).'],
    ['Facility Relocation FY2022: $38,500 one-time warehouse move from 18,000 to 32,000 sqft facility, Commerce City, CO.'],
    ['Charitable Contributions: Included in Office & General Admin. Non-business; will cease post-transaction.'],
    ['Personal Expenses in T&E/Payroll: Owner personal cell phones $8,400/yr, personal auto $12,000/yr, personal vacation travel.'],
    ['Executive Search: VP Operations hired Q3 FY2022 ($45,000), Director of Sales hired Q1 TTM ($22,000). Non-recurring.'],
  ], 'Income Statement', [42,16,16,16,22])

  add_sheet(wb1, [
    ['PINNACLE DISTRIBUTION CO., LLC'],
    ['Balance Sheet — Adjusted'],
    ['As of December 31 (FY20/21/22) and September 30, 2023 (TTM)'],
    [],
    ['($ in USD)', 'Dec 31, 2020','Dec 31, 2021','Dec 31, 2022','Sep 30, 2023'],
    [],
    ['ASSETS'],
    ['Cash & Cash Equivalents',               ...bs.cash],
    ['Accounts Receivable',                   ...bs.ar_gross],
    ['Less: Allowance for Doubtful Accounts', ...bs.ar_allow.map(v=>-v)],
    ['Net Accounts Receivable',               ...net_ar],
    ['Inventory - Finished Goods',            ...bs.inventory],
    ['Inventory - Raw Materials',             0,0,0,0],
    ['Total Inventory',                       ...bs.inventory],
    ['Prepaid Expenses',                      ...bs.prepaid],
    ['Other Current Assets',                  ...bs.other_ca],
    ['Total Current Assets',                  ...tot_ca],
    [],
    ['Property, Plant & Equipment (gross)',   ...bs.ppe_gross],
    ['Less: Accumulated Depreciation',        ...bs.accum_dep.map(v=>-v)],
    ['Net Fixed Assets',                      ...net_fa],
    ['Intangible Assets - Software',          0,0,0,0],
    ['Total Assets',                          ...tot_assets],
    [],
    ['LIABILITIES'],
    ['Accounts Payable - Trade',              ...bs.ap],
    ['Credit Cards Payable',                  ...bs.cc_pay],
    ['Accrued Payroll & Benefits',            ...bs.accrued_payroll],
    ['Accrued Expenses - Other',              ...bs.accrued_other],
    ['Sales Tax Payable',                     ...bs.sales_tax],
    ['Deferred Revenue',                      ...bs.deferred],
    ['Current Portion - LT Debt',             ...bs.curr_debt],
    ['Total Current Liabilities',             ...tot_cl],
    ['Long-Term Debt, Net of Current',        ...bs.lt_debt],
    ['Total Liabilities',                     ...tot_liab],
    [],
    ["STOCKHOLDERS' EQUITY"],
    ['Common Stock',                          ...bs.common_stk],
    ['Additional Paid-in Capital',            ...bs.apic],
    ['Retained Earnings',                     ...bs.re],
    ['Current Year Net Income',               ...ni],
    ["Total Stockholders' Equity",            ...tot_eq],
    ['Total Liabilities & Equity',            ...tot_assets],
    [],
    ['LT Debt: SBA 7(a) loan originated June 2017, original balance $2,400,000, matures June 2027, monthly P&I $33,333.'],
    ['Inventory: Finished goods at Commerce City DC and Phoenix 3PL. FIFO cost. Physical counts monthly.'],
    ['AR: B2B trade receivables. DSO ≈ 38 days. No significant bad debt history.'],
  ], 'Balance Sheet', [42,16,16,16,16])

  add_sheet(wb1, [
    ['PINNACLE DISTRIBUTION CO., LLC — Quality of Earnings: EBITDA Bridge'],
    [],
    ['($ in USD)', ...PERIODS],
    [],
    ['Net Income',                          ...ni],
    ['+ Interest Expense',                  ...int_ex],
    ['+ Tax Provision',                     ...tax],
    ['+ Depreciation',                      ...da],
    ['+ Amortization',                      0,0,0,0],
    ['EBITDA, as defined',                  ...ebitda],
    ['EBITDA Margin %',                     ...ebitda.map((e,i)=>pct(e/rev[i]))],
    [],
    ['MANAGEMENT ADJUSTMENTS'],
    ['a) Charitable Contributions',         ...mgmt_adj.charitable],
    ['b) Owner Tax & Legal (personal)',      ...mgmt_adj.owner_tax_legal],
    ['c) Excess Owner Compensation',        0,0,0,0],
    ['d) Personal Expenses',               ...mgmt_adj.personal_exp],
    ['Total Management Adjustments',        ...total_mgmt],
    ['Management-Adjusted EBITDA',          ...mgmt_ebitda],
    [],
    ['DILIGENCE ADJUSTMENTS'],
    ['e) Professional Fees - One-time',     ...dilig_adj.professional_fees_onetime],
    ['f) Executive Search Fees',            ...dilig_adj.executive_search],
    ['g) Facility Relocation',             ...dilig_adj.facility_relocation],
    ['h) Severance',                        ...dilig_adj.severance],
    ['i) Above-Market Rent',                0,0,0,0],
    ['j) Normalize Owner Comp',             ...dilig_adj.normalize_owner_comp],
    ['Total Diligence Adjustments',         ...total_dilig],
    ['Diligence-Adjusted EBITDA',           ...adj_ebitda],
    ['Adjusted EBITDA Margin %',            ...adj_ebitda.map((e,i)=>pct(e/rev[i]))],
    [],
    ['ADJUSTMENT NOTES:'],
    ['a) Charitable: Denver Food Bank + Mile High United Way. Management confirms will cease post-close.'],
    ['b) Owner tax/legal: Personal tax prep ($11,500) and personal estate legal fees billed to company.'],
    ['d) Personal expenses: Owner personal cell ($8,400), personal auto ($12,000), personal vacation travel.'],
    ['e) Non-recurring patent litigation with Acme Industrial Supply. Filed Q2 2021, settled October 2021.'],
    ['f) Retained executive search (Korn Ferry). Non-recurring hiring costs for two key leadership roles.'],
    ['g) One-time costs to relocate from 18,000 to 32,000 sqft warehouse. Includes movers, lease overlap, buildout.'],
    ['h) One-time severance to terminated Warehouse Manager (FY2021).'],
    ['j) CEO earns $475,000 (salary + distributions). Market replacement cost = $300,000. Normalizing to market rate.'],
  ], 'Quality of Earnings', [46,16,16,16,22])

  add_sheet(wb1, [
    ['PINNACLE DISTRIBUTION CO., LLC — Working Capital Analysis'],
    [],
    ['($ in USD)', 'Dec 31, 2020','Dec 31, 2021','Dec 31, 2022','Sep 30, 2023'],
    [],
    ['Accounts Receivable',           ...net_ar],
    ['Inventory',                     ...bs.inventory],
    ['Prepaid Expenses',              ...bs.prepaid],
    ['Other Current Assets',          ...bs.other_ca],
    ['(Less) Accounts Payable',       ...bs.ap.map(v=>-v)],
    ['(Less) Accrued Payroll',        ...bs.accrued_payroll.map(v=>-v)],
    ['(Less) Accrued Expenses',       ...bs.accrued_other.map(v=>-v)],
    ['(Less) Credit Cards Payable',   ...bs.cc_pay.map(v=>-v)],
    ['(Less) Sales Tax Payable',      ...bs.sales_tax.map(v=>-v)],
    ['(Less) Deferred Revenue',       ...bs.deferred.map(v=>-v)],
    ['Adjusted Net Working Capital',  ...PERIODS.map((_,i)=>net_ar[i]+bs.inventory[i]+bs.prepaid[i]+bs.other_ca[i]-bs.ap[i]-bs.accrued_payroll[i]-bs.accrued_other[i]-bs.cc_pay[i]-bs.sales_tax[i]-bs.deferred[i])],
    [],
    ['METRICS'],
    ['DSO (days)',         '36.5','38.2','38.7','39.8'],
    ['DOH (days)',         '93.4','90.7','95.8','94.2'],
    ['DPO (days)',         '42.1','44.8','46.2','48.3'],
    ['Inventory Turns',   '3.9x','4.0x','3.8x','3.9x'],
    ['Cash Conv. Cycle',  '87.8','84.1','88.3','85.7'],
    [],
    ['NWC Note: Excludes cash, current debt, and financing items. Adjusted NWC is the basis for working capital peg negotiation.'],
    ['DSO is consistent with net 30-day B2B terms. DPO reflects vendor payment terms (30-60 days). DOH reflects seasonal build of safety inventory.'],
  ], 'Working Capital', [42,16,16,16,16])

  add_sheet(wb1, [
    ['PINNACLE DISTRIBUTION CO., LLC — Proof of Cash'],
    [],
    ['($ in USD)', ...PERIODS],
    [],
    ['Total Bank Deposits per Bank Statements',  ...Object.values(bank_d).map(d=>d.reduce((a,b)=>a+b,0)), 61_895_507],
    ['Less: Beginning Accounts Receivable',      -1_987_345, ...bs.ar_gross.slice(0,3).map(v=>-v)],
    ['Add: Ending Accounts Receivable',          ...bs.ar_gross],
    ['Less: Non-Revenue Deposits',               0,-400_000,-650_000,-750_000],
    ['Estimated Revenues (Cash Basis)', ...[0,1,2,3].map(i=>{
      const bd=[bank_d.FY2020,bank_d.FY2021,bank_d.FY2022,[0,0,0,0,0,0,0,0,0,0,0,35_724_507]][i]
      const sum=bd.reduce((a,b)=>a+b,0)
      const beg=i===0?1_987_345:bs.ar_gross[i-1]
      const nrd=[0,-400_000,-650_000,-750_000][i]
      return sum-beg+bs.ar_gross[i]+nrd
    })],
    ['Revenue per General Ledger',               ...rev],
    ['Variance',                                 ...rev.map((_,i)=>{
      const bd=[bank_d.FY2020,bank_d.FY2021,bank_d.FY2022,[0]][i]
      if(!bd||!bd.length)return 'N/A'
      const sum=bd.reduce((a,b)=>a+b,0)
      const beg=i===0?1_987_345:bs.ar_gross[i-1]
      const nrd=[0,-400_000,-650_000,-750_000][i]
      return sum-beg+bs.ar_gross[i]+nrd-rev[i]
    })],
    [],
    ['Note: Company banks primarily with FirstBank Colorado x4821 (operating) and Bank of Colorado x2934 (savings).'],
    ['Non-revenue deposits FY2021: SBA loan draw $400,000. FY2022: Line of credit draws totaling $650,000.'],
    ['Immaterial variances attributable to timing/settlement lag of merchant processor deposits (1-2 business days).'],
  ], 'Proof of Cash', [46,16,16,16,22])

  add_sheet(wb1, [
    ['PINNACLE DISTRIBUTION CO., LLC — COGS Vendor Concentration'],
    ['Source: QuickBooks Vendor Detail Report, agreed to General Ledger'],
    [],
    ['($ in USD)', ...PERIODS],
    [],
    ...Object.entries(vendors).map(([name, vals]) => [name, ...vals]),
    ['All Other Vendors',       ...vendor_other],
    ['Total COGS',              ...cogs],
    [],
    ['% of Total COGS',         ...PERIODS],
    ...Object.entries(vendors).map(([name, vals]) => [name, ...vals.map((v,i)=>pct(v/cogs[i]))]),
    ['All Other Vendors',       ...vendor_other.map((v,i)=>pct(v/cogs[i]))],
    [],
    ['Global Manufacturing Inc.: 5-year supply agreement (renews Jan 2026). FOB destination. Net 60 terms. Primary PPE supplier.'],
    ['Pacific Logistics Group: Industrial tools. Annual pricing negotiation. Net 45 terms. Secondary supplier identified.'],
    ['Midwest Distribution Co.: MRO supplies. Net 30 terms. No long-term contract.'],
  ], 'COGS Vendors', [42,16,16,16,22])

  add_sheet(wb1, [
    ['PINNACLE DISTRIBUTION CO., LLC — Revenue by Product Line'],
    [],
    ['($ in USD)', ...PERIODS],
    [],
    ...Object.entries(prod).map(([name, vals]) => [name, ...vals]),
    ['Total Product Sales', ...PERIODS.map((_,i)=>Object.values(prod).reduce((s,v)=>s+v[i],0))],
    [],
    ['CHANNEL BREAKDOWN (% of total revenue):'],
    ['Direct B2B Sales Team', '58%','58%','58%','58%'],
    ['Amazon Business',       '22%','22%','22%','22%'],
    ['Company Website',       '12%','12%','12%','12%'],
    ['Wholesale/Distributor', ' 8%',' 8%',' 8%',' 8%'],
    [],
    ['CUSTOMER CONCENTRATION:'],
    ['Apex Construction Group',         pct(0.11),pct(0.11),pct(0.11),pct(0.11),'of TTM revenue'],
    ['Rocky Mountain Facilities Mgmt',  pct(0.08),pct(0.08),pct(0.08),pct(0.08),'of TTM revenue'],
    ['Colorado DOT',                     pct(0.06),pct(0.06),pct(0.06),pct(0.06),'of TTM revenue'],
    ['No single customer exceeds 15% of total revenue.'],
  ], 'Sales by Product', [45,16,16,16,22])

  add_sheet(wb1, [
    ['PINNACLE DISTRIBUTION CO., LLC — Monthly Margins (TTM Oct 2022 – Sep 2023)'],
    [],
    ['($ in USD)', ...months],
    [],
    ['Revenue',          ...m_rev],
    ['Cost of Goods Sold',...m_cogs],
    ['Operating Expenses',...m_opex],
    ['Gross Profit',     ...m_rev.map((r,i)=>r-m_cogs[i])],
    ['Gross Margin %',   ...m_rev.map((r,i)=>pct((r-m_cogs[i])/r))],
    ['Net Profit',       ...m_rev.map((r,i)=>r-m_cogs[i]-m_opex[i])],
    ['Net Margin %',     ...m_rev.map((r,i)=>pct((r-m_cogs[i]-m_opex[i])/r))],
    [],
    ['Revenue peaks Oct-Dec (Q4) due to government contract renewals and year-end safety budget spending.'],
  ], 'Margins by Month', [28,...months.map(()=>13)])

  add_sheet(wb1, [
    ['PINNACLE DISTRIBUTION CO., LLC — Company Overview & Key Findings'],
    [],
    ['Company:', 'Pinnacle Distribution Co., LLC'],
    ['Industry:', 'Industrial Supply Distribution (NAICS 423840)'],
    ['Founded:', 'January 2017'],
    ['Location:', 'Commerce City, Colorado'],
    ['Employees:', '78 FTE (as of December 31, 2022)'],
    ['Owner:', 'Michael Torres (100% owner, CEO)'],
    ['Accounting System:', 'QuickBooks Enterprise 22.0'],
    ['Inventory System:', 'Fishbowl Inventory'],
    ['Fiscal Year End:', 'December 31'],
    ['Accounting Basis:', 'Accrual / US GAAP'],
    ['CPA Firm:', 'Anderson & Mueller CPAs, Denver CO (compilation)'],
    [],
    ['KEY FINANCIAL SUMMARY'],
    ['($ in USD)', ...PERIODS],
    ['Reported Revenues',         ...rev],
    ['EBITDA, as defined',        ...ebitda],
    ['Diligence-Adjusted EBITDA', ...adj_ebitda],
    ['Adjusted EBITDA Margin %',  ...adj_ebitda.map((e,i)=>pct(e/rev[i]))],
    [],
    ['KEY FINDINGS:'],
    ['1. Revenue CAGR FY2020-FY2022: 22.1%. Growth driven by B2B sales team expansion and Amazon Business channel.'],
    ['2. Gross margin stable at ~38% across all periods. Improved vendor contracts partially offset freight inflation.'],
    ['3. Largest non-recurring items: patent litigation (FY21/22), facility relocation (FY22), executive search (FY22/TTM).'],
    ['4. Owner compensation normalization ($175,000/yr reduction to market) is the largest diligence adjustment.'],
    ['5. Working capital is healthy. DSO 38.7 days, DPO 46.2 days, Inventory Turns 3.8x — all within industry norms.'],
    ['6. Proof of cash agrees within immaterial variance for all periods. No revenue concerns identified.'],
    ['7. No unrecorded liabilities found in AP cutoff testing.'],
    ['8. Top 3 vendors represent 60% of COGS. Mitigated by 5-year contract with Global Mfg and identified backups.'],
  ], 'Overview', [30,22,22,22,24])

  wb_write(wb1, 'A_CLEAN_Pinnacle_Financial_Statements.xlsx')

  // File 2: GL + Trial Balance
  const wb2 = new_wb()
  const coa_rows = [
    ['PINNACLE DISTRIBUTION CO., LLC — Chart of Accounts'],
    [],
    ['Acct #', 'Account Name', 'Type', 'Balance FY2022'],
    ['1000','Cash - FirstBank Checking x4821','Bank', bs.cash[2]],
    ['1010','Cash - Bank of Colorado Savings x2934','Bank', bs.cash[2]*0.17],
    ['1100','Accounts Receivable','Accounts Receivable', bs.ar_gross[2]],
    ['1105','Allowance for Doubtful Accounts','Other Current Asset', -bs.ar_allow[2]],
    ['1200','Inventory - Finished Goods','Other Current Asset', bs.inventory[2]],
    ['1300','Prepaid Expenses','Other Current Asset', bs.prepaid[2]],
    ['1400','Other Current Assets','Other Current Asset', bs.other_ca[2]],
    ['1500','Property & Equipment','Fixed Asset', bs.ppe_gross[2]],
    ['1510','Accumulated Depreciation','Fixed Asset', -bs.accum_dep[2]],
    ['2000','Accounts Payable','Accounts Payable', -bs.ap[2]],
    ['2010','Credit Cards Payable','Credit Card', -bs.cc_pay[2]],
    ['2100','Accrued Payroll','Other Current Liability', -bs.accrued_payroll[2]],
    ['2110','Accrued Expenses','Other Current Liability', -bs.accrued_other[2]],
    ['2120','Sales Tax Payable','Other Current Liability', -bs.sales_tax[2]],
    ['2130','Deferred Revenue','Other Current Liability', -bs.deferred[2]],
    ['2200','Current Portion - LT Debt','Other Current Liability', -bs.curr_debt[2]],
    ['2300','SBA Loan - Long Term','Long Term Liability', -bs.lt_debt[2]],
    ['3000','Common Stock','Equity', -bs.common_stk[2]],
    ['3100','Additional Paid-in Capital','Equity', -bs.apic[2]],
    ['3900','Retained Earnings','Equity', -bs.re[2]],
    ['4000','Gross Product Sales','Income', -rev[2]*1.021],
    ['4010','Returns & Allowances','Income', rev[2]*0.0203],
    ['4020','Promotional Discounts','Income', rev[2]*0.0102],
    ['4100','Shipping Revenue','Income', -rev[2]*0.0143],
    ['4200','Other Revenue','Income', -24_321],
    ['5000','Product COGS','Cost of Goods Sold', cogs[2]*0.906],
    ['5100','Warehousing & Fulfillment','Cost of Goods Sold', cogs[2]*0.0643],
    ['5200','Shipping & Freight Out','Cost of Goods Sold', rev[2]*0.0143],
    ['5300','Inventory Adjustments','Cost of Goods Sold', 113_456],
    ['6000','Sales & Marketing','Expense', opex_detail.sales_marketing[2]],
    ['6100','Salaries, Wages & Payroll Tax','Expense', opex_detail.salaries_payroll[2]],
    ['6110','Employee Benefits & 401k','Expense', opex_detail.benefits_401k[2]],
    ['6200','Rent & Occupancy','Expense', opex_detail.rent_occupancy[2]],
    ['6300','Professional Services','Expense', opex_detail.professional_services[2]],
    ['6400','Technology & Software','Expense', opex_detail.technology_software[2]],
    ['6500','Insurance','Expense', opex_detail.insurance[2]],
    ['6600','Payment Processing Fees','Expense', opex_detail.payment_processing[2]],
    ['6700','Travel & Entertainment','Expense', opex_detail.travel_entertainment[2]],
    ['6800','Office & General Admin','Expense', opex_detail.office_admin[2]],
    ['7000','Interest Expense','Other Expense', int_ex[2]],
    ['7100','Depreciation & Amortization','Other Expense', da[2]],
    ['8000','Tax Provision','Other Expense', tax[2]],
  ]
  add_sheet(wb2, coa_rows, 'Chart of Accounts', [8,42,26,18])

  const tb_rows = [
    ['PINNACLE DISTRIBUTION CO., LLC — Trial Balance (All Periods)'],
    [],
    ['Account #','Account Name','Dec 31 2020','Dec 31 2021','Dec 31 2022','Sep 30 2023'],
    ['1100','Accounts Receivable',...bs.ar_gross],
    ['1200','Inventory - Finished Goods',...bs.inventory],
    ['1300','Prepaid Expenses',...bs.prepaid],
    ['1500','PP&E (net)',...net_fa],
    ['2000','Accounts Payable',...bs.ap.map(v=>-v)],
    ['2100','Accrued Payroll',...bs.accrued_payroll.map(v=>-v)],
    ['2110','Accrued Expenses',...bs.accrued_other.map(v=>-v)],
    ['2300','SBA Loan LT',...bs.lt_debt.map(v=>-v)],
    ['4000','Gross Product Sales',...rev.map(v=>-v)],
    ['5000','Product COGS',...cogs.map(c=>Math.round(c*0.906))],
    ['5100','Warehousing & Fulfillment',...cogs.map(c=>Math.round(c*0.0643))],
    ['6000','Sales & Marketing',...opex_detail.sales_marketing],
    ['6100','Salaries, Wages & Payroll Tax',...opex_detail.salaries_payroll],
    ['6200','Rent & Occupancy',...opex_detail.rent_occupancy],
    ['6300','Professional Services',...opex_detail.professional_services],
    ['7000','Interest Expense',...int_ex],
    ['7100','Depreciation & Amortization',...da],
    ['8000','Tax Provision',...tax],
  ]
  add_sheet(wb2, tb_rows, 'Trial Balance', [8,40,16,16,16,18])

  const gl_rows = [
    ['PINNACLE DISTRIBUTION CO., LLC — GL Detail FY2022 (Representative Sample)'],
    ['QuickBooks Enterprise Export | Accrual Basis | Jan 1 - Dec 31 2022'],
    [],
    ['Date','Acct #','Account Name','Vendor/Customer','Memo','Debit','Credit'],
    ['2022-01-12','1100','Accounts Receivable','Apex Construction Group','INV #22-0045 PPE Order',187_450,''],
    ['2022-01-12','4000','Gross Product Sales','Apex Construction Group','INV #22-0045 PPE Order','',187_450],
    ['2022-01-18','1100','Accounts Receivable','Rocky Mountain Facilities','INV #22-0067 MRO Supplies',143_230,''],
    ['2022-01-18','4000','Gross Product Sales','Rocky Mountain Facilities','INV #22-0067 MRO Supplies','',143_230],
    ['2022-01-14','5000','Product COGS','Global Manufacturing Inc.','PO #22-GM-001 PPE Inventory',534_000,''],
    ['2022-01-14','2000','Accounts Payable','Global Manufacturing Inc.','PO #22-GM-001 PPE Inventory','',534_000],
    ['2022-01-31','6100','Salaries, Wages & Payroll Tax','ADP Payroll','Bi-weekly payroll Jan 31',132_615,''],
    ['2022-01-31','1000','Cash','ADP Payroll','Bi-weekly payroll Jan 31','',132_615],
    ['2022-03-15','6300','Professional Services','Brownstein Hyatt Farber','Legal - Acme patent dispute - NON-RECURRING',45_000,''],
    ['2022-06-10','6300','Professional Services','Brownstein Hyatt Farber','Legal - patent compliance - NON-RECURRING',50_000,''],
    ['2022-04-15','6300','Professional Services','Allied Van Lines','Warehouse relocation Commerce City - NON-RECURRING',38_500,''],
    ['2022-07-01','6000','Sales & Marketing','Korn Ferry','Executive search VP Operations - NON-RECURRING',45_000,''],
    ['2022-12-31','6100','Salaries, Wages & Payroll Tax','Michael Torres (Owner/CEO)','CEO annual salary + bonus FY2022 - OWNER COMP $475,000',475_000,''],
    ['2022-12-15','6800','Office & General Admin','Denver Food Bank','Charitable donation - PERSONAL/NON-BUSINESS',42_000,''],
    ['2022-12-20','6800','Office & General Admin','Mile High United Way','Charitable donation - PERSONAL/NON-BUSINESS',19_000,''],
    ['2022-12-31','6700','Travel & Entertainment','Michael Torres','Personal vacation Maui - PERSONAL/NON-BUSINESS',11_200,''],
    ['2022-12-31','6100','Salaries, Wages & Payroll Tax','Michael Torres','Personal cell phone plan 4 lines - PERSONAL',8_400,''],
    ['2022-12-31','5300','Inventory Adjustments','','Year-end physical count variance shrinkage',113_456,''],
    ['2022-12-31','7100','Depreciation & Amortization','','Annual D&A forklifts shelving equipment',287_500,''],
    ['2022-12-31','7000','Interest Expense','FirstBank SBA Loan','SBA 7(a) loan interest FY2022 total',134_221,''],
    [],
    ['FY2022 ANNUAL TOTALS:'],
    ['','4000','Gross Product Sales','','Total Revenue','',33_124_650],
    ['','5000+5100+5200+5300','Total COGS','','Total COGS',20_124_249,''],
    ['','6000','Sales & Marketing','','',1_834_000,''],
    ['','6100','Salaries, Wages & Payroll Tax','','Incl $475k owner comp',3_298_000,''],
    ['','6300','Professional Services','','Incl $183k non-recurring items',197_234,''],
  ]
  add_sheet(wb2, gl_rows, 'GL Detail FY2022', [13,8,32,28,50,14,14])
  wb_write(wb2, 'A_CLEAN_Pinnacle_General_Ledger.xlsx')

  console.log('  → SET A complete. Revenue: $21.8M–$35.7M | Adj EBITDA margins ~10-13%')
}


// ─────────────────────────────────────────────────────────────────────────────
// ══════════════════════════════════════════════════════════════════════════════
// SET B: MESSY — Cascade Building Materials Inc.
// Manufacturing/distribution, inconsistent formatting, mixed terminology
// ══════════════════════════════════════════════════════════════════════════════
// ─────────────────────────────────────────────────────────────────────────────
function buildMessySet() {
  console.log('\n── SET B: MESSY (Cascade Building Materials Inc.) ──')

  // Financial data — different numbers, same structural reality
  const rev   = [18_432_109, 22_187_654, 28_934_210, 31_456_789]
  const cogs  = [11_982_871, 14_477_936, 18_807_237, 20_306_877] // ~65% GM 35%
  const gp    = rev.map((r,i) => r - cogs[i])
  const opex  = gp.map(g => Math.round(g * 0.75))  // ~26% of rev as opex leaves ~9% EBITDA
  const ebitda= gp.map((g,i)=>g-opex[i])
  const da    = [287_000, 312_000, 298_000, 267_000]
  const int_ex= [213_456, 187_654, 156_789, 134_567]
  const tax   = [ 34_210,  45_678,  56_789,  67_890]
  const ni    = ebitda.map((e,i)=>e-da[i]-int_ex[i]-tax[i])

  const wb = new_wb()

  // Sheet 1: P&L with intentionally inconsistent headers
  add_sheet(wb, [
    ['CASCADE BUILDING MATERIALS INC.'],
    ['Profit & Loss — Management Prepared'],
    ['For fiscal years ending 12/31 and trailing 12 months as of 9/30/23'],
    ['UNAUDITED — Internal Use Only'],
    [],
    // Inconsistent: years instead of "FY20" style, and "YTD" instead of "TTM"
    ['Line Item', 'Year 2020', 'Year 2021', 'Year 2022', 'YTD Thru Sep-23'],
    [],
    ['SALES'],
    // Revenue labeled differently than standard
    ['Gross Sales',                    ...rev.map(r=>r+Math.round(r*0.023))],
    ['  less: Returns/Credits',        ...rev.map(r=>-Math.round(r*0.018))],
    ['  less: Discounts Given',        ...rev.map(r=>-Math.round(r*0.0052))],
    // Note: missing "shipping revenue" line — intentionally partial
    ['NET SALES',                      ...rev],
    [],
    ['COST OF SALES'],
    // Mixed terminology
    ['Materials & Purchased Parts',    ...cogs.map(c=>Math.round(c*0.712))],
    ['Direct Labor (Warehouse)',       ...cogs.map(c=>Math.round(c*0.148))],
    ['Freight-In',                     ...cogs.map(c=>Math.round(c*0.089))],
    // Note: "Freight-Out" buried separately below, not in COGS
    ['Inventory Write-Offs',           ...cogs.map(c=>Math.round(c*0.015))],
    ['Other Direct Costs',             ...cogs.map(c=>Math.round(c*0.036))],
    ['TOTAL COST OF SALES',            ...cogs],
    [],
    ['GROSS PROFIT',                   ...gp],
    // % missing — no gross margin % listed here
    [],
    ['OPERATING EXPENSES'],
    // Some lines consolidated/relabeled vs standard
    ['Personnel (wages + benefits)',   ...opex.map(o=>Math.round(o*0.512))],
    ['Commissions & Marketing',        ...opex.map(o=>Math.round(o*0.181))],
    ['Facility & Rent',                ...opex.map(o=>Math.round(o*0.121))],
    ['Admin & Office',                 ...opex.map(o=>Math.round(o*0.063))],
    ['Legal & Accounting',             ...opex.map(o=>Math.round(o*0.047))],
    ['Freight-Out (customer delivery)',...rev.map(r=>Math.round(r*0.018))],  // misplaced
    ['Software / IT',                  ...opex.map(o=>Math.round(o*0.034))],
    ['Insurance Expense',              ...opex.map(o=>Math.round(o*0.042))],
    ['Total Operating Expenses',       ...opex],
    [],
    // EBITDA presented differently: "Operating Income" not EBITDA
    ['OPERATING INCOME',               ...ebitda],
    // Margin % presented differently
    ['Operating Margin',               ...ebitda.map((e,i)=>(e/rev[i]*100).toFixed(2)+'%')],
    [],
    ['OTHER INCOME/EXPENSE'],
    ['Interest Expense',               ...int_ex.map(v=>-v)],
    ['Depreciation',                   ...da.map(v=>-v)],   // Note: D&A inside "other" not separately broken out
    ['INCOME BEFORE TAXES',            ...ebitda.map((e,i)=>e-da[i]-int_ex[i])],
    ['Income Tax Expense',             ...tax.map(v=>-v)],
    ['NET INCOME',                     ...ni],
    [],
    // Notes buried at bottom, not clearly labeled
    ['Notes & Comments (FY2022):'],
    ['Owner/President Ryan Caldwell compensation included in Personnel line: $412,000 total (salary $300k + distributions $112k).'],
    ['Market rate for CEO replacement in this market: estimated $275,000-$295,000.'],
    ['Legal & Accounting FY21 includes one-time $87,500 for employment dispute settlement (settled March 2021).'],
    ['Legal & Accounting FY22 includes $62,000 for OSHA compliance consulting (non-recurring, completed Q3 2022).'],
    ['Admin & Office includes owner personal expenses: estimated $19,200/yr (cell phones x6 lines, personal subscriptions).'],
    ['Facility relocation Q2 FY22: moved to new 45,000 sqft facility. Total one-time costs $52,300 included in Facility & Rent.'],
    ['Freight-Out: Management is reconsidering whether this belongs in COGS or OpEx. Currently in OpEx for presentation purposes.'],
  ], 'P&L', [35,16,16,16,18])

  // Sheet 2: Balance Sheet — inconsistent date formats, different account groupings
  const bs_ap   = [1_456_789, 1_789_234, 2_134_567, 2_345_678]
  const bs_ar   = [1_876_543, 2_345_678, 2_987_654, 3_234_567]
  const bs_inv  = [2_987_654, 3_567_890, 4_234_567, 4_678_901]
  const bs_cash = [  987_654, 1_234_567, 1_567_890, 1_789_456]
  const bs_prep = [  198_765,   234_567,   298_765,   334_567]
  const bs_ppe  = [1_987_654, 1_876_543, 1_798_765, 1_712_345] // net
  const bs_ltdb = [1_500_000, 1_200_000,   900_000,   600_000]
  const bs_ap_curr = [400_000, 400_000, 400_000, 400_000]
  const bs_accr = [  345_678,   412_345,   498_765,   543_210]
  add_sheet(wb, [
    ['CASCADE BUILDING MATERIALS INC. — Balance Sheet'],
    // Inconsistent date formats
    ['','Dec-20','Dec-21','12/31/2022','9/30/23 (prelim)'],
    [],
    ['ASSETS'],
    ['Cash & equivalents',          ...bs_cash],
    ['Trade Receivables (net)',      ...bs_ar],   // Note: "net" here but no allowance breakdown
    ['Inventories',                  ...bs_inv],
    ['Prepaid & Other',              ...bs_prep],
    // "Total Current" missing — not calculated
    [],
    ['Fixed Assets (net of deprec)', ...bs_ppe],
    ['Other Long-term Assets',        12_345, 9_876, 8_765, 7_654],
    ['TOTAL ASSETS',                 ...bs_cash.map((_,i)=>bs_cash[i]+bs_ar[i]+bs_inv[i]+bs_prep[i]+bs_ppe[i]+[12_345,9_876,8_765,7_654][i])],
    [],
    ['LIABILITIES & EQUITY'],
    ['Accounts Payable',             ...bs_ap],
    ['Accrued Liabilities',          ...bs_accr],
    ['Current Debt (ST portion)',     ...bs_ap_curr],
    // Total current liabilities not shown
    [],
    ['Long-Term Debt',               ...bs_ltdb],
    ['Other Long-term Liab',          23_456, 18_234, 15_678, 12_345],
    // Equity is just a plug
    ['Total Equity',                 ...bs_cash.map((_,i)=>bs_cash[i]+bs_ar[i]+bs_inv[i]+bs_prep[i]+bs_ppe[i]+[12_345,9_876,8_765,7_654][i]-bs_ap[i]-bs_accr[i]-bs_ap_curr[i]-bs_ltdb[i]-[23_456,18_234,15_678,12_345][i])],
    ['TOTAL LIAB + EQUITY',          ...bs_cash.map((_,i)=>bs_cash[i]+bs_ar[i]+bs_inv[i]+bs_prep[i]+bs_ppe[i]+[12_345,9_876,8_765,7_654][i])],
    [],
    ['Note: Balance sheet is management-prepared. Fixed assets net of accumulated depreciation of $734k (2020), $1,046k (2021), $1,344k (2022), $1,611k (TTM).'],
    ['AR: Net of allowance for doubtful accounts of approximately $38k-$65k per period.'],
  ], 'Balance Sheet', [35,16,16,16,18])

  // Sheet 3: EBITDA bridge — presented differently, some items missing or mislabeled
  const mgmt_tot = [62_900, 78_450, 93_700, 98_200]
  const dilig_tot= [-88_700, -52_300, -28_500, -105_000]
  add_sheet(wb, [
    ['CASCADE BUILDING MATERIALS INC. — Normalized EBITDA'],
    ['Management prepared reconciliation of Net Income to Normalized EBITDA'],
    [],
    ['', 'CY 2020', 'CY 2021', 'CY 2022', 'TTM Sep-23'],
    [],
    ['Net Income per Financials',     ...ni],
    ['Add: Income Taxes',             ...tax],
    ['Add: Interest',                 ...int_ex],
    ['Add: Depreciation',             ...da],
    // Note: Amortization missing line
    ['EBITDA',                        ...ebitda],
    [],
    ['ADJUSTMENTS - MANAGEMENT PROPOSED:'],
    // Inconsistent labeling vs standard template
    ['+ Owner personal expenses',     19_200, 21_600, 19_200, 19_200],
    ['+ Owner charitable donations',  28_750, 35_000, 42_000, 44_500],
    ['+ Owner legal/tax (personal)',  14_950, 21_850, 32_500, 34_500],
    ['Total Mgmt. Addbacks',          ...mgmt_tot],
    ['Mgmt-Adjusted EBITDA',          ...ebitda.map((e,i)=>e+mgmt_tot[i])],
    [],
    ['DILIGENCE ADJUSTMENTS:'],
    // Some negative adjustments below market
    ['Employment dispute settlement (FY21)',  0,-87_500,0,0],
    ['OSHA consulting (non-recur FY22)',      0,0,-62_000,0],
    ['Facility relocation one-time (FY22)',   0,0,-52_300,0],
    // Owner comp normalization listed as a lump
    ['Normalize CEO comp to market ($285k)',  -75_000,-75_000,-75_000,-75_000],  // $412k-$285k=$127k adj??? Note: labeled as $75k — inconsistency for stress test
    ['Executive recruiting (non-recur)',       0,0,0,-30_000],
    ['Total Diligence Adjustments',           ...dilig_tot],
    ['NORMALIZED EBITDA',             ...ebitda.map((e,i)=>e+mgmt_tot[i]+dilig_tot[i])],
    ['Normalized EBITDA Margin %',    ...ebitda.map((e,i)=>((e+mgmt_tot[i]+dilig_tot[i])/rev[i]*100).toFixed(1)+'%')],
    [],
    ['Key Notes:'],
    ['Ryan Caldwell (Owner/President) draws $300k salary + approximately $112k in distributions annually.'],
    ['Market replacement cost for equivalent President/COO role: $275,000-$295,000 (management estimate).'],
    ['FY2021 employment settlement: wrongful termination claim, one-time. No pending litigation.'],
    ['OSHA consulting: completed Q3 2022. Non-recurring compliance work. Company now in full compliance.'],
    ['Facility relocation: moved operations from 28,000 sqft to 45,000 sqft in Renton, WA. Q2 2022.'],
  ], 'Normalized EBITDA', [42,15,15,15,16])

  // Sheet 4: Vendor spend — partial, inconsistent
  add_sheet(wb, [
    ['CASCADE BUILDING MATERIALS INC. — Key Vendor Spend (COGS)'],
    ['Source: Management prepared from QuickBooks AP detail'],
    [],
    // Missing one period column
    ['Vendor', 'FY2021', 'FY2022', 'TTM'],
    [],
    ['Summit Pacific Lumber Co.',     4_234_000, 5_456_000, 5_987_000],
    ['Renton Concrete Supply',        2_876_000, 3_678_000, 4_012_000],
    ['Pacific Rim Fasteners',         1_567_000, 2_012_000, 2_198_000],
    ['Cascade Steel Works',           1_234_000, 1_567_000, 1_712_000],
    ['Mountain West Transport',         876_000,   987_000, 1_087_000],
    ['All Other',                     3_690_936, 5_107_237, 5_310_877],
    ['TOTAL COGS',                    14_477_936,18_807_237,20_306_877],
    [],
    // Note: FY2020 missing intentionally
    ['Note: FY2020 data not available at vendor level. AP detail extracted from QuickBooks aging report.'],
    ['Summit Pacific Lumber: Primary lumber/plywood supplier. FOB origin. Net 45 days. Volume rebate 2%.'],
    ['Mountain West Transport: Third-party trucking for customer deliveries.'],
  ], 'Vendor Spend', [32,14,14,14])

  // Sheet 5: Working capital — monthly, mixed format
  add_sheet(wb, [
    ['CASCADE BUILDING MATERIALS INC. — Working Capital'],
    ['Key balance sheet metrics at period end'],
    [],
    ['', 'Dec 2020', 'Dec 2021', 'Dec 2022', 'Sep 2023'],
    [],
    ['CURRENT ASSETS'],
    ['Cash',                    ...bs_cash],
    ['Receivables (net)',        ...bs_ar],
    ['Inventory',                ...bs_inv],
    ['Prepaid',                  ...bs_prep],
    ['Total Current',            ...bs_cash.map((_,i)=>bs_cash[i]+bs_ar[i]+bs_inv[i]+bs_prep[i])],
    [],
    ['CURRENT LIABILITIES'],
    ['Accounts Payable',         ...bs_ap],
    ['Accrued Liabilities',      ...bs_accr],
    ['ST Debt Portion',          ...bs_ap_curr],
    ['Total Current',            ...bs_ap.map((_,i)=>bs_ap[i]+bs_accr[i]+bs_ap_curr[i])],
    [],
    // Adjusted NWC not explicitly calculated — Gemini must infer
    ['NET WORKING CAPITAL (excl cash & debt)',  ...bs_ar.map((_,i)=>bs_ar[i]+bs_inv[i]+bs_prep[i]-bs_ap[i]-bs_accr[i])],
    [],
    // DSO/DPO buried in narrative
    ['METRICS NOTE: Based on management estimates, DSO averaged approximately 30-35 days across periods reviewed.'],
    ['Inventory turns approximately 4.5x annually based on trailing 12 months COGS.'],
    ['DPO approximately 35-42 days. Vendor terms vary from Net 30 to Net 60.'],
  ], 'Working Capital', [35,16,16,16,16])

  // Sheet 6: Monthly revenue (only TTM, no COGS breakdown)
  const m_rev_b = [2_456_000,2_512_000,2_987_000,2_345_000,2_489_000,2_612_000,2_456_000,2_587_000,2_634_000,2_478_000,2_512_000,2_887_789]
  add_sheet(wb, [
    ['CASCADE BUILDING MATERIALS INC. — Monthly Revenue (TTM Oct 2022 – Sep 2023)'],
    ['Note: Monthly COGS detail not available. Gross margin estimated at 35% based on annual financials.'],
    [],
    ['Month','Net Revenue','Est. COGS (35%)','Est. Gross Profit','Est. GP %'],
    ...['Oct-22','Nov-22','Dec-22','Jan-23','Feb-23','Mar-23','Apr-23','May-23','Jun-23','Jul-23','Aug-23','Sep-23'].map((m,i)=>[
      m, m_rev_b[i], Math.round(m_rev_b[i]*0.65), Math.round(m_rev_b[i]*0.35), '35.0%'
    ]),
    ['TOTAL TTM', m_rev_b.reduce((a,b)=>a+b,0), '', '', ''],
    [],
    ['Note: Operating expense monthly detail not available. See annual P&L for full breakout.'],
  ], 'Monthly Revenue', [12,16,16,16,10])

  // Sheet 7: Bank statements — only 1 year, partially formatted
  const bank_b_2022 = [2_187_000,2_134_000,2_345_000,2_234_000,2_289_000,2_412_000,2_356_000,2_478_000,2_523_000,2_467_000,2_389_000,2_721_210]
  add_sheet(wb, [
    ['Cascade Building Materials — Bank Statement Reconciliation FY2022'],
    ['FirstState Bank Business Checking — Account No. xxxx-2847'],
    [],
    ['','Jan-22','Feb-22','Mar-22','Apr-22','May-22','Jun-22','Jul-22','Aug-22','Sep-22','Oct-22','Nov-22','Dec-22','TOTAL'],
    ['Deposits',...bank_b_2022, bank_b_2022.reduce((a,b)=>a+b,0)],
    [],
    // Narrative proof of cash
    ['CASH-TO-REVENUE RECONCILIATION (FY2022):'],
    ['Total bank deposits:',          bank_b_2022.reduce((a,b)=>a+b,0)],
    ['Less: beginning AR (12/31/21):', -2_345_678],
    ['Add: ending AR (12/31/22):',     2_987_654],
    ['Less: non-revenue deposits:',    -198_000],  // LOC draw, not labeled clearly
    ['Estimated revenue (cash basis):', bank_b_2022.reduce((a,b)=>a+b,0)-2_345_678+2_987_654-198_000],
    ['Revenue per P&L:',               28_934_210],
    ['Variance:',                      bank_b_2022.reduce((a,b)=>a+b,0)-2_345_678+2_987_654-198_000-28_934_210],
    [],
    ['Note: $198,000 non-revenue deposit represents a draw on our revolving line of credit (July 2022, FirstState Bank).'],
    ['All other deposits represent customer payments on trade receivables.'],
  ], 'Bank Stmt FY22', [28,10,10,10,10,10,10,10,10,10,10,10,10,14])

  wb_write(wb, 'B_MESSY_Cascade_Building_Materials.xlsx')
  console.log('  → SET B complete. Revenue: $18.4M–$31.5M | Multiple terminology inconsistencies | Partial data')
}


// ─────────────────────────────────────────────────────────────────────────────
// ══════════════════════════════════════════════════════════════════════════════
// SET C: VERY MESSY — Heartland Food Services LLC
// Multi-location restaurant/food service distributor. Raw QuickBooks export
// dump, inconsistent periods across tabs, numbers embedded in text,
// missing labels, duplicate entries, narrative-heavy. Max RAG stress test.
// ══════════════════════════════════════════════════════════════════════════════
// ─────────────────────────────────────────────────────────────────────────────
function buildVeryMessySet() {
  console.log('\n── SET C: VERY MESSY (Heartland Food Services LLC) ──')

  const rev   = [14_234_567, 17_890_123, 22_456_789, 24_789_012]
  const cogs  = [ 9_956_197, 12_523_086, 15_719_752, 17_352_308]  // ~70% COGS, 30% GM
  const gp    = rev.map((r,i)=>r-cogs[i])
  const opex  = gp.map(g=>Math.round(g*0.72))   // ~8.4% EBITDA
  const ebitda= gp.map((g,i)=>g-opex[i])
  const da    = [187_000, 198_000, 234_000, 212_000]
  const int_ex= [123_456, 112_345, 98_765, 87_654]
  const tax   = [ 23_456,  32_789,  43_210,  51_234]
  const ni    = ebitda.map((e,i)=>e-da[i]-int_ex[i]-tax[i])

  const wb = new_wb()

  // Sheet 1: Raw QB export — completely unstandardized
  add_sheet(wb, [
    // No clear header
    ['Heartland Food Services LLC'],
    ['QuickBooks Online — Profit and Loss — Custom'],
    ['January 2020 through September 2023'],
    ['Cash Basis'],   // Note: says Cash Basis but it's actually accrual — common error
    ['Run Date: October 5 2023'],
    [],
    // Periods in different columns, no standard labels
    ['', 'Jan 2020-Dec 2020', 'Jan 2021-Dec 2021', 'Jan 2022-Dec 2022', 'Oct 2022-Sep 2023'],
    [],
    ['Income'],
    // Revenue broken into many small lines, no subtotals visible
    ['  Food Product Sales - Wholesale',  ...rev.map(r=>Math.round(r*0.62))],
    ['  Food Product Sales - Retail',     ...rev.map(r=>Math.round(r*0.19))],
    ['  Catering & Event Services',       ...rev.map(r=>Math.round(r*0.11))],
    ['  Delivery Fees Collected',         ...rev.map(r=>Math.round(r*0.054))],
    ['  Other Income (misc)',             ...rev.map(r=>Math.round(r*0.026))],
    ['Total Income',                      ...rev],
    [],
    ['Cost of Goods Sold'],
    ['  Food & Beverage Purchases',       ...cogs.map(c=>Math.round(c*0.748))],
    ['  Kitchen Labor (variable)',        ...cogs.map(c=>Math.round(c*0.089))],
    ['  Packaging Supplies',              ...cogs.map(c=>Math.round(c*0.067))],
    ['  Cold Storage & Refrigeration',    ...cogs.map(c=>Math.round(c*0.054))],
    ['  Delivery Cost of Sales',         ...cogs.map(c=>Math.round(c*0.042))],
    // Line without clear label — raw QB export artifact
    ['  Cost of Goods Sold - Other',      ...cogs.map(c=>Math.round(c*0.000))],
    ['  (COGS Adjustments/Write-offs)',   ...cogs.map(c=>Math.round(c*(-0.000)))],
    ['Total Cost of Goods Sold',          ...cogs],
    [],
    ['Gross Profit',                      ...gp],
    [],
    ['Expenses'],
    // Granular QB accounts — many lines
    ['  Advertising',                     ...opex.map(o=>Math.round(o*0.098))],
    ['  Auto & Truck Expense',            ...opex.map(o=>Math.round(o*0.043))],
    ['  Bank Charges & Fees',             ...opex.map(o=>Math.round(o*0.028))],
    ['  Computer & Internet',             ...opex.map(o=>Math.round(o*0.019))],
    ['  Consulting Fees',                 ...opex.map(o=>Math.round(o*0.067))],
    ['  Depreciation Expense',            ...da],  // Note: D&A inside expenses (not below EBITDA)
    ['  Dues & Subscriptions',            ...opex.map(o=>Math.round(o*0.008))],
    ['  Employee Benefits',               ...opex.map(o=>Math.round(o*0.072))],
    ['  Insurance Expense',               ...opex.map(o=>Math.round(o*0.054))],
    ['  Interest Expense',                ...int_ex],  // Also inside expenses
    ['  Meals & Entertainment',           ...opex.map(o=>Math.round(o*0.031))],
    ['  Office Supplies',                 ...opex.map(o=>Math.round(o*0.012))],
    ['  Officer Salary',                  ...opex.map(o=>Math.round(o*0.089))],  // Owner comp buried here
    ['  Payroll Expenses',                ...opex.map(o=>Math.round(o*0.278))],
    ['  Payroll Tax Expense',             ...opex.map(o=>Math.round(o*0.034))],
    ['  Rent Expense',                    ...opex.map(o=>Math.round(o*0.112))],
    ['  Repairs & Maintenance',           ...opex.map(o=>Math.round(o*0.023))],
    ['  Telephone & Utilities',           ...opex.map(o=>Math.round(o*0.019))],
    ['  Travel',                          ...opex.map(o=>Math.round(o*0.011))],
    // Some lines that are clearly non-business
    ['  Charitable Contributions',        ...opex.map(o=>Math.round(o*0.028))],
    ['  Personal - Cell Phone (owner)',   ...opex.map(_=>9_600)],
    ['  Personal - Vehicle (owner)',      ...opex.map(_=>14_400)],
    // One-time items buried without annotation
    ['  Legal Settlement - 2021',          0, 67_890, 0, 0],
    ['  Kitchen Equipment Move',           0, 0, 43_210, 0],
    ['  Recruiting Fees',                  0, 0, 38_000, 22_500],
    ['Total Expenses',                    ...opex.map((o,i)=>o+da[i]+int_ex[i]+tax[i])],  // Includes D&A, interest, tax
    [],
    // Net income is after everything since D&A and interest are IN expenses
    ['Net Income',                        ...ni],
    [],
    // Narrative embedded in the spreadsheet
    ['IMPORTANT NOTES FROM OWNER (Sarah Whitfield):'],
    ['- The 2021 Legal Settlement of $67,890 was for a slip-and-fall at our Omaha location. One time. Paid to plaintiff.'],
    ['- Kitchen Equipment Move ($43,210 in 2022): we relocated our Topeka commissary kitchen to a larger space. This is done.'],
    ['- My salary (Officer Salary above) was $287,500 in 2022. Industry comp for this role would be $225,000-$240,000.'],
    ['- Personal Cell and Vehicle above are mine. Accountant said to include them for tax purposes.'],
    ['- Charitable Contributions are to local food banks. We likely will continue this post-sale but buyer can decide.'],
    ['- Recruiting fees: hired a new District Manager in 2022 and an Operations Director in 2023. Non-recurring.'],
    ['- 2021 was impacted by COVID supply chain issues which inflated our food costs by approximately $180,000.'],
  ], 'QB P&L Export Raw', [38,18,18,18,18])

  // Sheet 2: Balance sheet in completely different format
  add_sheet(wb, [
    ['HEARTLAND FOOD SERVICES LLC'],
    ['Balance Sheet as of dates listed below'],
    ['(management prepared, unaudited)'],
    [],
    // Only 2 periods shown, labeled inconsistently
    ['', '', 'Dec 31 2022', 'As of Sept 30 2023'],
    [],
    ['WHAT WE OWN:'],
    ['Cash in bank (operating acct)',     '',  1_234_567,  1_456_789],
    ['Cash in bank (payroll acct)',       '',     87_654,     92_345],
    ['Accounts receivable (customers)',   '', 2_123_456, 2_345_678],
    ['Inventory (food & supplies)',       '', 1_456_789, 1_587_654],
    ['Prepaid expenses',                  '',   167_890,   189_234],
    ['Equipment & FF&E (net)',            '',   987_654, 1_012_345],
    ['Vehicles (net)',                    '',   234_567,   198_765],
    ['Leasehold improvements (net)',      '',   145_678,   134_567],
    ['Goodwill',                          '', 1_500_000, 1_500_000],  // Note: goodwill present
    ['Other assets',                      '',    34_567,    28_901],
    ['TOTAL ASSETS',                      '', 7_972_822, 8_546_278],
    [],
    ['WHAT WE OWE:'],
    ['Accounts payable (vendors)',        '', 1_123_456, 1_234_567],
    ['Credit card balances',              '',   134_567,   145_678],
    ['Payroll taxes owed',                '',    78_901,    82_345],
    ['Sales tax payable',                 '',    45_678,    49_234],
    ['Line of credit (operating)',        '',   500_000,   250_000],
    ['Equipment loan balance',            '',   456_789,   312_345],
    ['SBA loan (7a)',                     '',   789_012,   667_890],
    ['Other accrued expenses',            '',    89_012,    92_345],
    ['TOTAL LIABILITIES',                 '', 3_217_415, 2_834_404],
    [],
    ["OWNER'S EQUITY:"],
    ['Total equity (plug)',                '', 4_755_407, 5_711_874],
    ['TOTAL LIAB + EQUITY',               '', 7_972_822, 8_546_278],
    [],
    ['Note: Balance sheets for FY2020 and FY2021 not available in final form. See prior year QB snapshots in separate file.'],
    ['Equipment net = Gross $2.1M less accumulated depreciation of $1.1M as of 12/31/2022.'],
    ['Goodwill from 2019 acquisition of Omaha Commissary Partners LLC. Not amortized (indefinite life).'],
    ['LOC with Heartland Federal Credit Union, $750k limit, prime+1%. Drew down in seasonal months.'],
  ], 'Balance Sheet', [30,5,16,16])

  // Sheet 3: Vendor analysis — very partial, wrong format
  add_sheet(wb, [
    ['Heartland Food Services — Top Vendor Spend'],
    ['This is approximate. Exact figures would need to be pulled from QB if required.'],
    [],
    ['Vendor Name', 'What They Supply', 'Approx Annual Spend (FY2022)', 'Notes'],
    ['Sysco Corporation',         'Broadline food distributor',    8_234_000, 'Main supplier. Net 14 days.'],
    ['US Foods',                  'Broadline food, backup',        2_456_000, 'Secondary supplier.'],
    ['Gordon Food Service',       'Specialty items',                 876_000, 'Regional. 30-day terms.'],
    ['Bunzl Distribution',        'Packaging/supplies',              654_000, 'Monthly account.'],
    ['AmeriQuest Food Service',   'Paper goods, disposables',        345_000, 'Net 30.'],
    ['Various (refrigeration)',   'Cold storage partners',           456_000, 'Variable, 3 locations.'],
    ['All Other',                 'Misc vendors < $100k each',     2_697_752, 'Multiple small vendors'],
    ['TOTAL COGS',                '',                             15_718_752, 'Approximately agrees to P&L'],
    [],
    ['Note: Sysco and US Foods together represent about 68% of our food purchases. We have no long-term contracts with either.'],
    ['We are in negotiations to extend our Sysco agreement with better pricing. No guarantee of outcome.'],
    ['FY2020 and FY2021 vendor detail not readily available. Sysco/US Foods have been primary suppliers since 2019.'],
  ], 'Vendor Spend', [30,28,20,38])

  // Sheet 4: AR aging — different format, narrative mixed in
  add_sheet(wb, [
    ['Heartland Food Svc — Customer AR Summary (pulled from QB, 12/31/22)'],
    [],
    ['Customer', 'Amount Owed', 'Days Outstanding', 'Status'],
    ['Midwest Grocery Chain (corporate)', 678_234, '< 30 days', 'Current'],
    ['Heartland School District',         234_567, '< 30 days', 'Current'],
    ['Omaha Convention Center',           189_012, '31-45 days', 'Slightly past due'],
    ['Kansas City Caterers LLC',          123_456, '< 30 days', 'Current'],
    ['Various (many small accounts)',     898_187, '< 60 days', 'Mix'],
    ['TOTAL AR',                        2_123_456, '', ''],
    [],
    ['Note from owner Sarah: We have very few bad debts. Most customers are institutions or established businesses.'],
    ['The $898k "various" category is approximately 340 separate customer accounts, all < $5,000 each.'],
    ['We do not maintain a formal allowance for doubtful accounts. Bad debt has been < 0.1% historically.'],
    ['DSO is approximately 28-33 days based on our receivable collection terms (Net 30 standard).'],
  ], 'AR Aging Dec22', [35,16,16,20])

  // Sheet 5: Working capital — just a narrative text dump
  add_sheet(wb, [
    ['HEARTLAND FOOD SERVICES — WORKING CAPITAL DISCUSSION'],
    [],
    ['The following is management commentary on our working capital position. Formal schedules available on request.'],
    [],
    ['ACCOUNTS RECEIVABLE:'],
    ['As of 12/31/2022, our total accounts receivable was $2,123,456. This is collected primarily within 30 days as most'],
    ['of our customers are institutional (schools, hospitals, municipal cafeterias) with reliable payment. We have not had'],
    ['significant bad debt issues. Our largest receivable at year end was Midwest Grocery Chain at $678,234 which was'],
    ['paid in January 2023 per their normal terms.'],
    [],
    ['INVENTORY:'],
    ['Our food inventory at 12/31/2022 was approximately $1,456,789. We hold approximately 18-22 days of inventory'],
    ['on hand at any time due to the perishable nature of our products. We do monthly physical counts at all 3 locations.'],
    ['There were no significant inventory write-offs in FY2022 beyond normal spoilage ($12,340 written off).'],
    [],
    ['ACCOUNTS PAYABLE:'],
    ['AP at year end was $1,123,456. Most of our vendor terms are Net 14 (Sysco) to Net 30. We pay on time.'],
    ['There are no past-due amounts. DPO is approximately 12-16 days driven by Sysco short terms.'],
    [],
    ['NET WORKING CAPITAL ESTIMATE:'],
    ['A rough NWC calculation (excl cash and debt): AR $2,123,456 + Inventory $1,456,789 + Prepaid $167,890'],
    ['minus AP $1,123,456 minus Accrued $89,012 minus Sales tax $45,678 minus Credit cards $134,567 = ~$2,355,422'],
    [],
    ['As of 9/30/2023 (TTM period end): AR $2,345,678 + Inventory $1,587,654 + Prepaid $189,234'],
    ['minus AP $1,234,567 minus Accrued $92,345 minus Sales tax $49,234 minus CC $145,678 = ~$2,600,742'],
  ], 'Working Capital', [100])

  // Sheet 6: Monthly data — only 6 months, inconsistent
  add_sheet(wb, [
    ['Heartland Food Services — Monthly P&L Snapshot (Selected Months 2022-2023)'],
    ['NOTE: Monthly data only available for the 6 months shown. Other months estimated by management.'],
    [],
    ['', 'Jan-23', 'Feb-23', 'Mar-23', 'Apr-23', 'May-23', 'Jun-23'],
    ['Revenue',    1_987_000, 1_876_000, 2_098_000, 2_134_000, 2_212_000, 2_098_000],
    ['Food COGS',  1_356_000, 1_289_000, 1_434_000, 1_458_000, 1_512_000, 1_436_000],
    ['Gross Profit', 631_000,   587_000,   664_000,   676_000,   700_000,   662_000],
    ['Operating Expenses', 498_000, 478_000, 512_000, 523_000, 534_000, 514_000],
    // Inconsistent: some months include D&A, some don't
    ['(incl D&A and interest in opex above for some months)'],
    ['Net Income (approx)', 133_000, 109_000, 152_000, 153_000, 166_000, 148_000],
    [],
    ['Revenue Note: July-December 2022 and July-September 2023 not available monthly. Annual totals in main P&L.'],
    ['These 6 months represent approximately 51% of TTM revenue ($24.8M TTM).'],
    [],
    ['GM % ranged from 31.2% to 31.7% in these months — consistent with full-year 30% gross margin.'],
  ], 'Monthly Snapshot', [35,14,14,14,14,14,14])

  // Sheet 7: EBITDA bridge — scattered narrative, no clear table
  add_sheet(wb, [
    ['HEARTLAND FOOD SERVICES LLC — EBITDA NORMALIZATION'],
    ['Prepared by: Sarah Whitfield (owner) with help from our CPA Jim Anderson'],
    ['Date: November 2023'],
    [],
    ['We are providing the following analysis of our normalized EBITDA for the buyer\'s review.'],
    ['These are our best estimates and we are happy to discuss.'],
    [],
    ['Starting point — Net Income from QB P&L:'],
    ['FY2020 net income: $', ni[0]],
    ['FY2021 net income: $', ni[1]],
    ['FY2022 net income: $', ni[2]],
    ['Trailing twelve months (Oct 2022-Sep 2023) net income: $', ni[3]],
    [],
    ['Adding back non-cash and financing items to get to EBITDA:'],
    ['Depreciation (included in expenses in our QB): FY20 $187k, FY21 $198k, FY22 $234k, TTM $212k'],
    ['Interest expense: FY20 $123k, FY21 $112k, FY22 $99k, TTM $88k'],
    ['Income taxes: minimal (we\'re an LLC, pass-through). FY20 $23k, FY21 $33k, FY22 $43k, TTM $51k'],
    [],
    ['EBITDA (approx): FY2020 ~$', ebitda[0], '  FY2021 ~$', ebitda[1]],
    ['                  FY2022 ~$', ebitda[2], '  TTM ~$', ebitda[3]],
    [],
    ['Items we believe should be added back:'],
    ['1) My salary - I pay myself $287,500 per year (Officer Salary line in QB). A hired CEO for this type and size'],
    ['   of business would cost $225,000-$240,000. So add-back is approximately $50,000-$60,000 per year.'],
    ['2) My personal cell phone on company account: $9,600/year. 4 phones for family.'],
    ['3) My personal vehicle: $14,400/year. Personal car, not a company vehicle.'],
    ['4) Charitable contributions: $28k-$40k/year. These are my personal charitable choices.'],
    ['5) Legal settlement (2021 only): $67,890. Slip and fall. Done.'],
    ['6) Kitchen relocation (2022 only): $43,210. Moved commissary to bigger space. Done.'],
    ['7) Recruiting fees (2022: $38k, 2023: $22.5k): Hired District Manager and Ops Director. Non-recurring.'],
    ['8) Covid supply chain impact (2021 only): Food cost inflation approximately $180,000 above normal.'],
    [],
    ['ESTIMATED NORMALIZED EBITDA:'],
    // No structured table — just text
    ['If you add back items 1-4 for all years, and items 5-8 for relevant years:'],
    ['FY2020: EBITDA $', ebitda[0], ' + addbacks ~$', 62_000, ' = approx $', ebitda[0]+62_000],
    ['FY2021: EBITDA $', ebitda[1], ' + addbacks ~$', 328_890, ' = approx $', ebitda[1]+328_890],
    ['FY2022: EBITDA $', ebitda[2], ' + addbacks ~$', 153_210, ' = approx $', ebitda[2]+153_210],
    ['TTM: EBITDA $', ebitda[3], ' + addbacks ~$', 104_000, ' = approx $', ebitda[3]+104_000],
    [],
    ['These are rough estimates. We expect the buyer\'s financial diligence firm will refine these numbers.'],
    ['Happy to provide QB access, bank statements, tax returns for any period needed.'],
  ], 'EBITDA Bridge (Notes)', [55,18,10,18,10,18])

  // Sheet 8: Totally unrelated raw data dump — challenges the RAG to ignore noise
  add_sheet(wb, [
    ['Heartland Food Services — Miscellaneous Financial Notes & Documents'],
    [],
    ['This sheet contains various financial notes that don\'t fit neatly elsewhere.'],
    [],
    ['TAX RETURN SUMMARY (from CPA Jim Anderson):'],
    ['2022 Form 1065 (Partnership return): Total revenue $22,456,789. Ordinary business income $879,234.'],
    ['2021 Form 1065: Total revenue $17,890,123. Ordinary business income $432,890.'],
    ['2020 Form 1065: Total revenue $14,234,567. Ordinary business income $212,340.'],
    [],
    ['LOAN SUMMARY:'],
    ['SBA 7(a) Loan - Heartland Federal Credit Union: Original $1,200,000 (2019). Balance 12/31/2022: $789,012.'],
    ['Equipment Loan - John Deere Financial: Balance 12/31/2022: $456,789. For refrigerated delivery vehicles.'],
    ['Operating Line of Credit - Heartland FCU: $750,000 limit. Drawn $500,000 at 12/31/22. Prime + 1%.'],
    [],
    ['KEY PERFORMANCE INDICATORS (owner-generated):'],
    ['Average order value: $1,240 (FY2022)'],
    ['Number of active accounts (FY2022): approximately 387'],
    ['Revenue per employee (FY2022): approximately $253,000 (89 FTE)'],
    ['Locations: Omaha NE (headquarters + commissary), Kansas City MO (satellite), Topeka KS (commissary)'],
    [],
    ['LEASE SUMMARY:'],
    ['Omaha HQ/Commissary: 28,000 sqft. $18,500/month. Lease expires June 2026.'],
    ['Kansas City Satellite: 8,500 sqft. $6,200/month. Month-to-month (negotiating extension).'],
    ['Topeka Commissary: 12,000 sqft. $8,800/month. Lease expires December 2024.'],
    ['Total annual rent: approximately $404,400.'],
    [],
    ['INSURANCE:'],
    ['General Liability: $2M/$4M limits. Annual premium: $67,890.'],
    ['Workers Comp: Premium FY2022 $123,456. Claim-free for 3 years.'],
    ['Commercial Auto: 12 vehicles. Annual premium $34,567.'],
    [],
    ['EMPLOYEES:'],
    ['FY2022: 89 FTE (43 full-time kitchen/warehouse, 28 drivers, 12 sales/admin, 6 management).'],
    ['FY2021: 82 FTE. FY2020: 71 FTE.'],
    ['No union. No pending labor disputes.'],
    ['Turnover has been high in kitchen staff (~40%) but lower for drivers (~15%) and management (~8%).'],
    [],
    ['MISCELLANEOUS:'],
    ['Company was originally founded as Whitfield Family Catering in 2012, rebranded to Heartland Food Services in 2016.'],
    ['Sarah Whitfield is 100% owner. No other equity holders.'],
    ['No related party transactions except: Sarah\'s husband Kevin Whitfield is a driver (W-2 income $62,400/yr, market rate).'],
    ['The company has $0 deferred revenue at any balance sheet date (all services delivered before billing).'],
  ], 'Misc Notes', [90])

  wb_write(wb, 'C_VERY_MESSY_Heartland_Food_Services.xlsx')
  console.log('  → SET C complete. Revenue: $14.2M–$24.8M | Raw QB dump | Narrative-heavy | Maximum RAG stress')
}


// ─────────────────────────────────────────────────────────────────────────────
// MAIN
// ─────────────────────────────────────────────────────────────────────────────
console.log('\n╔══════════════════════════════════════════════════════════════╗')
console.log('║  QoE Platform — Mock Financial Data Generator                ║')
console.log('║  Generating 3 test sets (Clean / Messy / Very Messy)         ║')
console.log('╚══════════════════════════════════════════════════════════════╝')

buildCleanSet()
buildMessySet()
buildVeryMessySet()

console.log('\n╔══════════════════════════════════════════════════════════════╗')
console.log('║  ALL FILES GENERATED IN: mock-data/                          ║')
console.log('╠══════════════════════════════════════════════════════════════╣')
console.log('║  A_CLEAN_Pinnacle_Financial_Statements.xlsx — 9 sheets        ║')
console.log('║  A_CLEAN_Pinnacle_General_Ledger.xlsx — 3 sheets              ║')
console.log('║  B_MESSY_Cascade_Building_Materials.xlsx — 7 sheets           ║')
console.log('║  C_VERY_MESSY_Heartland_Food_Services.xlsx — 8 sheets         ║')
console.log('╠══════════════════════════════════════════════════════════════╣')
console.log('║  UPLOAD ORDER FOR TESTING:                                    ║')
console.log('║  1. New workbook "Pinnacle Distribution" → upload A_CLEAN*   ║')
console.log('║  2. New workbook "Cascade Building" → upload B_MESSY*        ║')
console.log('║  3. New workbook "Heartland Food" → upload C_VERY_MESSY*     ║')
console.log('╚══════════════════════════════════════════════════════════════╝\n')
