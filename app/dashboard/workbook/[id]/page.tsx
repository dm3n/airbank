'use client'

import { useState, use, useEffect, useCallback, useRef, useMemo } from 'react'
import { useRouter } from 'next/navigation'
import { Badge } from '@/components/ui/badge'
import { Button } from '@/components/ui/button'
import { LineChart, Line, BarChart, Bar, ComposedChart, PieChart as RechartsPie, Pie, Cell, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer, LabelList } from 'recharts'
import {
  Table,
  TableBody,
  TableCell,
  TableHead,
  TableHeader,
  TableRow,
} from '@/components/ui/table'
import {
  DropdownMenu,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuTrigger,
} from '@/components/ui/dropdown-menu'
import { Download, ChevronDown, Database, FileBarChart, TrendingUp, FileText, Scale, DollarSign, ShoppingCart, Calendar, Banknote, Package, ClipboardCheck, BarChart3, PieChart, Settings, Loader2, Shield, Sparkles, Flag } from 'lucide-react'
import Image from 'next/image'
import { AuditableCell, type SourceRef, type CellFlag } from '@/components/auditable-cell'
import dynamic from 'next/dynamic'
const WorkbookSettingsDialog = dynamic(() => import('@/components/workbook-settings-dialog').then(m => ({ default: m.WorkbookSettingsDialog })), { ssr: false })
const DocumentViewerPanel = dynamic(() => import('@/components/document-viewer-panel').then(m => ({ default: m.DocumentViewerPanel })), { ssr: false })
const RiskDiligenceSection = dynamic(() => import('@/components/risk-diligence-section').then(m => ({ default: m.RiskDiligenceSection })), { ssr: false, loading: () => <div className="animate-pulse h-64 rounded-lg bg-slate-50" /> })
import { useLayoutContext } from '@/lib/layout-context'

const formatCurrency = (value: number) => {
  return new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD', minimumFractionDigits: 0 }).format(value)
}

const formatPercent = (value: number) => {
  return `${value.toFixed(1)}%`
}

// Quality of Earnings Data
const qoeData = [
  { item: 'Net Income', fy20: 4558429, fy21: 6091874, fy22: 7120305, ttm: 7858378, source: 'Audited Financial Statements - Year End Dec 31' },
  { item: 'Interest Expense', fy20: 42850, fy21: 52387, fy22: 61245, ttm: 65182, source: 'General Ledger Acct #7100 - Interest Expense' },
  { item: 'Tax Provision', fy20: 23185, fy21: 27038, fy22: 29583, ttm: 31256, source: 'Form 1120 - Federal & State Income Tax Returns' },
  { item: 'Depreciation', fy20: 4853, fy21: 5250, fy22: 5390, ttm: 4700, source: 'Fixed Asset Schedule - Acc Depreciation Roll-Forward' },
  { item: 'Amortization', fy20: 1200, fy21: 1400, fy22: 1300, ttm: 1300, source: 'Intangible Asset Amortization Schedule - Software & Patents' },
  { item: 'EBITDA, as defined', fy20: 4630517, fy21: 6177949, fy22: 7217823, ttm: 7960816, source: 'Calculated per Management Methodology', isBold: true },
  { item: 'a) Charitable Contributions', fy20: 15000, fy21: 248500, fy22: 287350, ttm: 315920, source: 'GL Acct #6850 - Discretionary charitable donations to local foundation', isAdjustment: true },
  { item: 'b) Owner Tax & Legal', fy20: 18500, fy21: 18500, fy22: 21200, ttm: 22450, source: 'GL Acct #6420 - Personal CPA & estate planning fees', isAdjustment: true },
  { item: 'c) Excess Owner Compensation', fy20: 185200, fy21: 208750, fy22: 236280, ttm: 248190, source: 'Payroll Register - Excess above market CEO salary of $175K', isAdjustment: true },
  { item: 'd) Personal Expenses', fy20: 26574, fy21: 31850, fy22: 35894, ttm: 39863, source: 'GL Acct #6900 - Country club, personal travel, family expenses', isAdjustment: true },
  { item: 'Total Management Adjustments', fy20: 245274, fy21: 507600, fy22: 580724, ttm: 626423, source: 'Sum of Management Add-backs', isBold: true },
  { item: 'Management-Adjusted EBITDA', fy20: 4875791, fy21: 6685549, fy22: 7798547, ttm: 8587239, source: 'EBITDA as Defined + Management Adjustments', isBold: true },
  { item: 'e) Professional Fees - One-time', fy20: 0, fy21: 52400, fy22: 58900, ttm: 68200, source: 'GL Acct #6425 - Litigation settlement & SEC filing costs related to terminated acquisition', isAdjustment: true },
  { item: 'f) Executive Search Fees', fy20: 0, fy21: 65000, fy22: 48000, ttm: 52000, source: 'GL Acct #6340 - Non-recurring CFO & VP Sales recruiting expenses', isAdjustment: true },
  { item: 'g) Facility Relocation', fy20: 0, fy21: 0, fy22: 38750, ttm: 41500, source: 'GL Acct #6520 - One-time warehouse relocation & build-out costs', isAdjustment: true },
  { item: 'h) Severance', fy20: 0, fy21: 0, fy22: 85400, ttm: 92300, source: 'GL Acct #6120 - Restructuring severance for 3 terminated employees', isAdjustment: true },
  { item: 'i) Above-Market Rent', fy20: 0, fy21: -158714, fy22: -162500, ttm: -168900, source: 'GL Acct #6500 - Normalize to market rate of $18/sqft vs $28/sqft paid (related party lease)', isAdjustment: true },
  { item: 'j) Normalize Owner Comp', fy20: -244500, fy21: -382300, fy22: -404000, ttm: -425900, source: 'Payroll Register - Reduce total owner comp to market CEO salary of $175K + benefits', isAdjustment: true },
  { item: 'Total Diligence Adjustments', fy20: -244500, fy21: -423614, fy22: -335450, ttm: -340800, source: 'Sum of Diligence Add-backs (Net)', isBold: true },
  { item: 'Diligence-Adjusted EBITDA', fy20: 4875291, fy21: 6524835, fy22: 7682947, ttm: 8547239, source: 'Final Normalized Run-Rate EBITDA - Buyer Perspective', isBold: true },
]

// Income Statement Data
const incomeStatementData = [
  { item: 'Gross Product Sales', fy20: 47832185, fy21: 58926741, fy22: 71584293, ttm: 77842953, pct: 114, category: 'Revenue' },
  { item: 'Returns & Allowances', fy20: -2847293, fy21: -3428574, fy22: -4183928, ttm: -4562485, pct: -7, category: 'Revenue' },
  { item: 'Promotional Discounts', fy20: -3218547, fy21: -3956382, fy22: -4813574, ttm: -5246193, pct: -8, category: 'Revenue' },
  { item: 'Net Product Sales', fy20: 41766345, fy21: 51541785, fy22: 62586791, ttm: 68034275, pct: 100, category: 'Revenue', isBold: true },
  { item: 'Shipping Revenue', fy20: 421111, fy21: 514829, fy22: 625868, ttm: 680343, pct: 1, category: 'Revenue' },
  { item: 'Other Revenue', fy20: 0, fy21: -573711, fy22: -467278, ttm: -420876, pct: -1, category: 'Revenue' },
  { item: 'Total Net Revenue', fy20: 42187456, fy21: 51482903, fy22: 62745381, ttm: 68293742, pct: 100, category: 'Revenue', isBold: true },
  { item: 'Product Cost of Goods Sold', fy20: 12656238, fy21: 15444871, fy22: 18823614, ttm: 20488124, pct: 30, category: 'COGS' },
  { item: 'Warehousing & Fulfillment', fy20: 3374995, fy21: 4118632, fy22: 5019631, ttm: 5463500, pct: 8, category: 'COGS' },
  { item: 'Shipping & Freight Out', fy20: 2109372, fy21: 2574145, fy22: 3137272, ttm: 3414687, pct: 5, category: 'COGS' },
  { item: 'Inventory Adjustments', fy20: 84375, fy21: 102966, fy22: 125491, ttm: 136587, pct: 0, category: 'COGS' },
  { item: 'Total Cost of Goods Sold', fy20: 18224980, fy21: 22240614, fy22: 27106008, ttm: 29502898, pct: 43, category: 'COGS', isBold: true },
  { item: 'Gross Profit', fy20: 23962476, fy21: 29242289, fy22: 35639373, ttm: 38790844, pct: 57, category: 'Gross Profit', isBold: true },
  { item: 'Sales & Marketing', fy20: 10547114, fy21: 12870726, fy22: 15686345, ttm: 17073436, pct: 25, category: 'OpEx' },
  { item: 'Salaries, Wages & Payroll Tax', fy20: 2531923, fy21: 3088974, fy22: 3764574, ttm: 4097686, pct: 6, category: 'OpEx' },
  { item: 'Employee Benefits & 401k', fy20: 316490, fy21: 386122, fy22: 470572, ttm: 512211, pct: 1, category: 'OpEx' },
  { item: 'Rent & Occupancy', fy20: 336000, fy21: 441000, fy22: 457500, ttm: 468000, pct: 1, category: 'OpEx' },
  { item: 'Professional Services', fy20: 591123, fy21: 721083, fy22: 878636, ttm: 956314, pct: 1, category: 'OpEx' },
  { item: 'Technology & Software', fy20: 548274, fy21: 669045, fy22: 815229, ttm: 887662, pct: 1, category: 'OpEx' },
  { item: 'Insurance', fy20: 210937, fy21: 257418, fy22: 313636, ttm: 341484, pct: 0, category: 'OpEx' },
  { item: 'Payment Processing Fees', fy20: 1687498, fy21: 2059316, fy22: 2509815, ttm: 2731750, pct: 4, category: 'OpEx' },
  { item: 'Travel & Entertainment', fy20: 168750, fy21: 205932, fy22: 250982, ttm: 273175, pct: 0, category: 'OpEx' },
  { item: 'Office & General Admin', fy20: 253124, fy21: 308897, fy22: 376473, ttm: 409763, pct: 1, category: 'OpEx' },
  { item: 'Total Operating Expenses', fy20: 19331959, fy21: 23581462, fy22: 28721589, ttm: 31251028, pct: 46, category: 'OpEx', isBold: true },
  { item: 'Operating Income (EBITDA)', fy20: 4630517, fy21: 6177949, fy22: 7217823, ttm: 7960816, pct: 12, category: 'Operating Income', isBold: true },
  { item: 'Interest Expense', fy20: -42850, fy21: -52387, fy22: -61245, ttm: -65182, pct: 0, category: 'Other' },
  { item: 'Depreciation & Amortization', fy20: -6053, fy21: -6650, fy22: -6690, ttm: -6000, pct: 0, category: 'Other' },
  { item: 'Income Before Tax', fy20: 4581614, fy21: 6118912, fy22: 7149888, ttm: 7889634, pct: 12, category: 'Other', isBold: true },
  { item: 'Tax Provision', fy20: -23185, fy21: -27038, fy22: -29583, ttm: -31256, pct: 0, category: 'Other' },
  { item: 'Net Income', fy20: 4558429, fy21: 6091874, fy22: 7120305, ttm: 7858378, pct: 12, category: 'Net Income', isBold: true },
]

// Balance Sheet Data
const balanceSheetData = [
  { item: 'Cash & Cash Equivalents', jan21: 1547293, jan22: 2183746, jan23: 2847562, source: 'Bank Reconciliations - Wells Fargo Acct ***4729 & ***2183', category: 'Current Assets' },
  { item: 'Accounts Receivable', jan21: 1284738, jan22: 1586420, jan23: 1924857, source: 'A/R Aging Detail Report by Customer as of 12/31', category: 'Current Assets' },
  { item: 'Allowance for Doubtful Accounts', jan21: -25695, jan22: -31728, jan23: -38497, source: 'Historical Bad Debt Analysis - 2% reserve methodology', category: 'Current Assets' },
  { item: 'Net Accounts Receivable', jan21: 1259043, jan22: 1554692, jan23: 1886360, source: 'A/R net of allowance', category: 'Current Assets', isBold: true },
  { item: 'Inventory - Raw Materials', jan21: 284759, jan22: 352185, jan23: 428473, source: 'Physical Inventory Count & Valuation - FIFO method', category: 'Current Assets' },
  { item: 'Inventory - Finished Goods', jan21: 758024, jan22: 937493, jan23: 1140266, source: 'Perpetual Inventory System Report', category: 'Current Assets' },
  { item: 'Total Inventory', jan21: 1042783, jan22: 1289678, jan23: 1568739, source: 'Sum of inventory components', category: 'Current Assets', isBold: true },
  { item: 'Prepaid Expenses', jan21: 42187, jan22: 51483, jan23: 62745, source: 'Prepaid Schedule - Insurance, Software, Rent', category: 'Current Assets' },
  { item: 'Other Current Assets', jan21: 18500, jan22: 22800, jan23: 27700, source: 'Deposits & miscellaneous receivables', category: 'Current Assets' },
  { item: 'Total Current Assets', jan21: 3909806, jan22: 5102399, jan23: 6393106, source: 'Sum of all current assets', category: 'Current Assets', isBold: true },
  { item: 'Property, Plant & Equipment', jan21: 127450, jan22: 145280, jan23: 168950, source: 'Fixed Asset Register - Historical Cost', category: 'Fixed Assets' },
  { item: 'Less: Accumulated Depreciation', jan21: -89614, jan22: -94864, jan23: -99254, source: 'Accumulated Depreciation Schedule', category: 'Fixed Assets' },
  { item: 'Net Fixed Assets', jan21: 37836, jan22: 50416, jan23: 69696, source: 'PP&E net book value', category: 'Fixed Assets', isBold: true },
  { item: 'Intangible Assets - Software', jan21: 24000, jan22: 28000, jan23: 26000, source: 'Capitalized Software Development Costs', category: 'Fixed Assets' },
  { item: 'Total Assets', jan21: 3971642, jan22: 5180815, jan23: 6488802, source: 'Sum of all asset categories', category: 'Total Assets', isBold: true },
  { item: 'Accounts Payable - Trade', jan21: 789342, jan22: 976528, jan23: 1188247, source: 'A/P Aging by Vendor as of 12/31', category: 'Current Liabilities' },
  { item: 'Credit Cards Payable', jan21: 94531, jan22: 116966, jan23: 142295, source: 'Amex Corporate Card Statement - Balance Forward', category: 'Current Liabilities' },
  { item: 'Accrued Payroll & Benefits', jan21: 227384, jan22: 281292, jan23: 342187, source: 'Payroll Accrual Schedule - Last pay period + PTO liability', category: 'Current Liabilities' },
  { item: 'Accrued Expenses - Other', jan21: 168750, jan22: 208593, jan23: 253827, source: 'Accrued Utilities, Professional Fees, Marketing', category: 'Current Liabilities' },
  { item: 'Sales Tax Payable', jan21: 126656, jan22: 156735, jan23: 190734, source: 'Sales Tax Returns - State Filings (CA, TX, NY, FL)', category: 'Current Liabilities' },
  { item: 'Deferred Revenue', jan21: 0, jan22: 85000, jan23: 127500, source: 'Customer Deposits & Prepayments', category: 'Current Liabilities' },
  { item: 'Current Portion - LT Debt', jan21: 125000, jan22: 125000, jan23: 125000, source: 'Term Loan Agreement - Bank of America', category: 'Current Liabilities' },
  { item: 'Total Current Liabilities', jan21: 1531663, jan22: 1950114, jan23: 2369790, source: 'Sum of all current liabilities', category: 'Current Liabilities', isBold: true },
  { item: 'Long-Term Debt, Net of Current', jan21: 875000, jan22: 750000, jan23: 625000, source: 'Term Loan Schedule - Original $1.5M, 5yr amortization', category: 'LT Liabilities' },
  { item: 'Total Liabilities', jan21: 2406663, jan22: 2700114, jan23: 2994790, source: 'Sum of current + long-term liabilities', category: 'Total Liabilities', isBold: true },
  { item: 'Common Stock', jan21: 100000, jan22: 100000, jan23: 100000, source: 'Certificate of Incorporation - 10,000 shares authorized, 1,000 issued', category: 'Equity' },
  { item: 'Additional Paid-in Capital', jan21: 0, jan22: 0, jan23: 0, source: 'No additional capital contributions', category: 'Equity' },
  { item: 'Retained Earnings', jan21: -93450, jan22: 1288827, jan23: 2673707, source: 'Retained Earnings Roll-Forward', category: 'Equity' },
  { item: 'Current Year Net Income', jan21: 1558429, jan22: 2091874, jan23: 2720305, source: 'Net Income per books', category: 'Equity' },
  { item: 'Total Stockholders\' Equity', jan21: 1564979, jan22: 2480701, jan23: 3494012, source: 'Sum of equity accounts', category: 'Equity', isBold: true },
  { item: 'Total Liabilities & Equity', jan21: 3971642, jan22: 5180815, jan23: 6488802, source: 'Must equal Total Assets', category: 'Total', isBold: true },
]

// Sales by Channel Data (TTM Jan-23)
const salesChannelData = [
  { product: 'Premium Series A - Commercial', revenue: 10244061, pct: 15.0, qty: 58520, avgPrice: 175.05, source: 'NetSuite Sales Report by SKU - Item #PS-A-COM-001' },
  { product: 'Standard Line B - Residential', revenue: 8195249, pct: 12.0, qty: 124857, avgPrice: 65.65, source: 'NetSuite Sales Report by SKU - Item #STD-B-RES-005' },
  { product: 'Pro Edition C - Industrial', revenue: 6829374, pct: 10.0, qty: 42684, avgPrice: 160.00, source: 'NetSuite Sales Report by SKU - Item #PRO-C-IND-120' },
  { product: 'Value Pack D - Bulk', revenue: 6146436, pct: 9.0, qty: 256893, avgPrice: 23.93, source: 'NetSuite Sales Report by SKU - Item #VAL-D-BLK-250' },
  { product: 'Elite Model E - Enterprise', revenue: 5463498, pct: 8.0, qty: 21859, avgPrice: 249.95, source: 'NetSuite Sales Report by SKU - Item #ELT-E-ENT-999' },
  { product: 'Classic Series F - Commercial', revenue: 4780562, pct: 7.0, qty: 95611, avgPrice: 50.00, source: 'NetSuite Sales Report by SKU - Item #CLS-F-COM-050' },
  { product: 'Compact Unit G - Residential', revenue: 4097623, pct: 6.0, qty: 136587, avgPrice: 30.00, source: 'NetSuite Sales Report by SKU - Item #CMP-G-RES-030' },
  { product: 'Advanced Kit H - Industrial', revenue: 3414687, pct: 5.0, qty: 42684, avgPrice: 80.00, source: 'NetSuite Sales Report by SKU - Item #ADV-H-IND-080' },
  { product: 'Starter Set I - Entry Level', revenue: 2731749, pct: 4.0, qty: 227323, avgPrice: 12.02, source: 'NetSuite Sales Report by SKU - Item #STR-I-ENT-012' },
  { product: 'Professional J - Specialty', revenue: 2048812, pct: 3.0, qty: 20488, avgPrice: 100.00, source: 'NetSuite Sales Report by SKU - Item #PRF-J-SPC-100' },
  { product: 'All Other Products (840 SKUs)', revenue: 14341691, pct: 21.0, qty: 487382, avgPrice: 29.43, source: 'NetSuite Sales Summary - Remaining SKUs aggregated' },
  { product: 'Total Product Sales', revenue: 68293742, pct: 100.0, qty: 1514888, avgPrice: 45.08, source: 'NetSuite TTM Revenue Roll-up Report', isBold: true },
]

// Margins by Month Data (Last 12 months)
const months = ['Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec', 'Jan']
const marginsByMonthData = [
  { month: 'Feb', revenue: 5285000, cogs: 2272000, opex: 2431000 },
  { month: 'Mar', revenue: 5692000, cogs: 2448000, opex: 2620000 },
  { month: 'Apr', revenue: 5439000, cogs: 2339000, opex: 2502000 },
  { month: 'May', revenue: 5848000, cogs: 2514000, opex: 2690000 },
  { month: 'Jun', revenue: 5534000, cogs: 2380000, opex: 2546000 },
  { month: 'Jul', revenue: 5621000, cogs: 2417000, opex: 2586000 },
  { month: 'Aug', revenue: 5748000, cogs: 2472000, opex: 2644000 },
  { month: 'Sep', revenue: 5494000, cogs: 2362000, opex: 2527000 },
  { month: 'Oct', revenue: 5963000, cogs: 2564000, opex: 2743000 },
  { month: 'Nov', revenue: 5385000, cogs: 2315000, opex: 2477000 },
  { month: 'Dec', revenue: 5659000, cogs: 2433000, opex: 2603000 },
  { month: 'Jan', revenue: 5662000, cogs: 2435000, opex: 2605000 },
]

// Proof of Cash Data
const proofOfCashData = [
  { month: 'Feb', deposits: 5318475, beginningAR: 1586420, endingAR: 1624738, revenueGL: 5284750, nonRevDeposits: 72043, variance: 0 },
  { month: 'Mar', deposits: 5728549, beginningAR: 1624738, endingAR: 1668294, revenueGL: 5692410, nonRevDeposits: 79985, variance: 0 },
  { month: 'Apr', deposits: 5447182, beginningAR: 1668294, endingAR: 1705847, revenueGL: 5418293, nonRevDeposits: 66636, variance: 0 },
  { month: 'May', deposits: 5682594, beginningAR: 1705847, endingAR: 1745938, revenueGL: 5651847, nonRevDeposits: 70838, variance: 0 },
  { month: 'Jun', deposits: 5924738, beginningAR: 1745938, endingAR: 1788562, revenueGL: 5892156, nonRevDeposits: 78218, variance: 0 },
  { month: 'Jul', deposits: 5761847, beginningAR: 1788562, endingAR: 1826485, revenueGL: 5729384, nonRevDeposits: 74528, variance: 0 },
  { month: 'Aug', deposits: 6002185, beginningAR: 1826485, endingAR: 1868394, revenueGL: 5968247, nonRevDeposits: 85341, variance: 0 },
  { month: 'Sep', deposits: 5677582, beginningAR: 1868394, endingAR: 1904728, revenueGL: 5647291, nonRevDeposits: 67025, variance: 0 },
  { month: 'Oct', deposits: 5948294, beginningAR: 1904728, endingAR: 1947183, revenueGL: 5915832, nonRevDeposits: 75117, variance: 0 },
  { month: 'Nov', deposits: 5770482, beginningAR: 1947183, endingAR: 1985946, revenueGL: 5738549, nonRevDeposits: 70154, variance: 0 },
  { month: 'Dec', deposits: 5725947, beginningAR: 1985946, endingAR: 2018046, revenueGL: 5692847, nonRevDeposits: 65146, variance: 0 },
  { month: 'Jan', deposits: 5693728, beginningAR: 2018046, endingAR: 2053638, revenueGL: 5662136, nonRevDeposits: 67046, variance: 0 },
]

// Working Capital Data
const workingCapitalData = [
  { item: 'Accounts Receivable', current: 1886360, normalized: 1823947, adjustment: -62413, days: 10.1, targetDays: 9.8, source: 'A/R Aging Report - DSO calculation based on TTM revenue $68.3M' },
  { item: 'Inventory', current: 1568739, normalized: 1687429, adjustment: 118690, days: 19.4, targetDays: 20.9, source: 'Inventory Analysis - DIO calculation vs. COGS run-rate' },
  { item: 'Prepaid Expenses', current: 62745, normalized: 62745, adjustment: 0, days: 0.3, targetDays: 0.3, source: 'Prepaid Schedule - Insurance, Rent, Software normalized' },
  { item: 'Other Current Assets', current: 27700, normalized: 27700, adjustment: 0, days: 0.1, targetDays: 0.1, source: 'Deposits & misc receivables - no adjustment required' },
  { item: 'Accounts Payable', current: -1188247, normalized: -1247586, adjustment: -59339, days: 14.7, targetDays: 15.5, source: 'A/P Aging Report - DPO calculation vs. COGS + OpEx' },
  { item: 'Accrued Expenses', current: -342187, normalized: -342187, adjustment: 0, days: 1.8, targetDays: 1.8, source: 'Payroll & expense accruals normalized to run-rate' },
  { item: 'Credit Cards Payable', current: -142295, normalized: -142295, adjustment: 0, days: 0.8, targetDays: 0.8, source: 'Corporate cards - normal operating level' },
  { item: 'Sales Tax Payable', current: -190734, normalized: -190734, adjustment: 0, days: 1.0, targetDays: 1.0, source: 'Sales tax liability - statutory requirement' },
  { item: 'Deferred Revenue', current: -127500, normalized: -127500, adjustment: 0, days: 0.7, targetDays: 0.7, source: 'Customer deposits - no normalization' },
  { item: 'Adjusted Net Working Capital', current: 1554581, normalized: 1550519, adjustment: -4062, days: 8.3, targetDays: 8.3, source: 'Sum of normalized NWC components', isBold: true },
]

// COGS Vendors Data
const cogsVendorsData = [
  { vendor: 'Global Manufacturing Inc.', fy20: 3647797, fy21: 4448123, fy22: 5421201, ttm: 5900580, pct: 20.0, source: 'A/P Subledger by Vendor - Primary component supplier (China)' },
  { vendor: 'Pacific Logistics Group', fy20: 2738348, fy21: 3336092, fy22: 4065901, ttm: 4425435, pct: 15.0, source: 'A/P Subledger - Freight forwarding & warehousing services' },
  { vendor: 'Midwest Distribution Co.', fy20: 2190678, fy21: 2668877, fy22: 3253921, ttm: 3540348, pct: 12.0, source: 'A/P Subledger - Fulfillment & 3PL services (7 warehouses)' },
  { vendor: 'TechComponents LLC', fy20: 1825562, fy21: 2224061, fy22: 2712601, ttm: 2951348, pct: 10.0, source: 'A/P Subledger - Electronic components & assemblies' },
  { vendor: 'Atlantic Packaging Solutions', fy20: 1642486, fy21: 2001907, fy22: 2440741, ttm: 2655261, pct: 9.0, source: 'A/P Subledger - Custom packaging materials & design' },
  { vendor: 'Sterling Materials Group', fy20: 1095991, fy21: 1334604, fy22: 1626361, ttm: 1770174, pct: 6.0, source: 'A/P Subledger - Raw materials supplier (metals & plastics)' },
  { vendor: 'Express Freight Services', fy20: 912493, fy21: 1112170, fy22: 1355300, ttm: 1475145, pct: 5.0, source: 'A/P Subledger - Last-mile delivery & expedited shipping' },
  { vendor: 'Quality Control Labs', fy20: 729994, fy21: 889336, fy22: 1084241, ttm: 1180116, pct: 4.0, source: 'A/P Subledger - Product testing & certifications (ISO/UL)' },
  { vendor: 'All Other Vendors (143 vendors)', fy20: 3442631, fy21: 4196444, fy22: 5145741, ttm: 5604491, pct: 19.0, source: 'A/P Subledger Summary - Remaining supplier base' },
  { vendor: 'Total COGS', fy20: 18224980, fy21: 22240614, fy22: 27106008, ttm: 29502898, pct: 100.0, source: 'Total COGS per Income Statement', isBold: true },
]

// AP/Accrual Testing Data
const testingData = [
  { checkNum: 'ACH-18472', invoice: 'GMI-2023-0847', payee: 'Global Manufacturing Inc.', checkDate: '2023-01-05', checkAmt: 287450, periodAmt: 287450, futureAmt: 0, notAccrued: 0 },
  { checkNum: 'ACH-18485', invoice: 'PLG-INV-4729', payee: 'Pacific Logistics Group', checkDate: '2023-01-08', checkAmt: 156780, periodAmt: 156780, futureAmt: 0, notAccrued: 0 },
  { checkNum: 'ACH-18503', invoice: 'MDC-2023-0194', payee: 'Midwest Distribution Co.', checkDate: '2023-01-12', checkAmt: 198650, periodAmt: 198650, futureAmt: 0, notAccrued: 0 },
  { checkNum: 'ACH-18518', invoice: 'TC-47283', payee: 'TechComponents LLC', checkDate: '2023-01-15', checkAmt: 124830, periodAmt: 124830, futureAmt: 0, notAccrued: 0 },
  { checkNum: 'CHK-9284', invoice: 'APS-JAN-2023', payee: 'Atlantic Packaging Solutions', checkDate: '2023-01-18', checkAmt: 89470, periodAmt: 89470, futureAmt: 0, notAccrued: 0 },
  { checkNum: 'ACH-18547', invoice: 'SMG-8472-A', payee: 'Sterling Materials Group', checkDate: '2023-01-22', checkAmt: 67240, periodAmt: 67240, futureAmt: 0, notAccrued: 0 },
  { checkNum: 'CHK-9301', invoice: 'EFS-012023', payee: 'Express Freight Services', checkDate: '2023-01-25', checkAmt: 52830, periodAmt: 52830, futureAmt: 0, notAccrued: 0 },
  { checkNum: 'ACH-18574', invoice: 'QCL-2023-Q1', payee: 'Quality Control Labs', checkDate: '2023-01-28', checkAmt: 41850, periodAmt: 18950, futureAmt: 22900, notAccrued: 0 },
  { checkNum: 'CHK-9318', invoice: 'AWS-JAN-2023', payee: 'Amazon Web Services', checkDate: '2023-01-30', checkAmt: 28470, periodAmt: 28470, futureAmt: 0, notAccrued: 0 },
  { checkNum: 'ACH-18592', invoice: 'SHP-847291', payee: 'Shopify Inc.', checkDate: '2023-01-31', checkAmt: 15680, periodAmt: 15680, futureAmt: 0, notAccrued: 0 },
  { checkNum: 'CHK-9325', invoice: 'INS-Q1-2023', payee: 'Liberty Mutual Insurance', checkDate: '2023-02-01', checkAmt: 47250, periodAmt: 0, futureAmt: 47250, notAccrued: 0 },
  { checkNum: 'ACH-18608', invoice: 'META-JAN-2023', payee: 'Meta Platforms Inc.', checkDate: '2023-02-02', checkAmt: 68940, periodAmt: 68940, futureAmt: 0, notAccrued: 0 },
  { checkNum: 'ACH-18621', invoice: 'GOOG-ADV-847', payee: 'Google LLC', checkDate: '2023-02-03', checkAmt: 124680, periodAmt: 124680, futureAmt: 0, notAccrued: 0 },
  { checkNum: 'CHK-9342', invoice: 'CTR-FEB-2023', payee: 'Smith & Associates Consulting', checkDate: '2023-02-05', checkAmt: 35800, periodAmt: 0, futureAmt: 35800, notAccrued: 0 },
  { checkNum: 'ACH-18647', invoice: 'SAL-JAN-2023', payee: 'Salesforce.com Inc.', checkDate: '2023-02-06', checkAmt: 12450, periodAmt: 12450, futureAmt: 0, notAccrued: 0 },
]

const testingTotals = {
  totalTested: testingData.reduce((sum, row) => sum + row.checkAmt, 0),
  totalDisbursements: 5847293,
  percentTested: 23.7,
  totalError: 0,
}

export default function WorkbookPage({ params }: { params: Promise<{ id: string }> }) {
  const { id } = use(params)

  const DEMO_NAMES: { [key: string]: string } = {
    'sandbox': 'Alpine Outdoor Co.',
    '1': 'Acme Corp',
    '2': 'TechStart Inc',
    '3': 'Global Industries',
    '4': 'Finance Group',
    '5': 'Retail Solutions',
  }

  const isDemoWorkbook = /^\d$/.test(id) || id === 'sandbox'
  const router = useRouter()

  const [activeSection, setActiveSection] = useState('qoe')
  const [settingsOpen, setSettingsOpen] = useState(false)

  // ── Live data layer ─────────────────────────────────────────────
  interface LiveCell {
    id: string
    section: string
    row_key: string
    period: string
    raw_value: number | null
    display_value: string | null
    is_calculated: boolean
    is_overridden: boolean
    source_page: number | null
    source_excerpt: string | null
    confidence: number | null
    source_document: { id: string; file_name: string } | null
    flags?: CellFlag[]
  }

  const [liveCells, setLiveCells] = useState<LiveCell[]>([])
  const [cellsLoading, setCellsLoading] = useState(false)
  const [cellsError, setCellsError] = useState<string | null>(null)
  const [viewerOpen, setViewerOpen] = useState(false)
  const [viewerSource, setViewerSource] = useState<SourceRef | null>(null)
  const [exporting, setExporting] = useState(false)
  const [exportError, setExportError] = useState<string | null>(null)
  const [liveWorkbookName, setLiveWorkbookName] = useState<string | null>(null)
  const [workbookStatus, setWorkbookStatus] = useState<string | null>(null)
  const [workbookPeriods, setWorkbookPeriods] = useState<string[]>(['FY20', 'FY21', 'FY22', 'TTM'])
  interface FlagItem {
    id: string
    section: string
    row_key: string
    period: string | null
    flag_type: string
    severity: string
    title: string
    resolved_at: string | null
    created_by_ai: boolean
  }
  const [flags, setFlags] = useState<FlagItem[]>([])
  const [flagsOpen, setFlagsOpen] = useState(false)
  const [pendingScroll, setPendingScroll] = useState<string | null>(null)
  const mainScrollRef = useRef<HTMLDivElement>(null)
  const [completeness, setCompleteness] = useState<number | null>(null)
  const [reconciliation, setReconciliation] = useState<{
    status: 'PASS' | 'WARN' | 'FAIL' | 'UNAVAILABLE'
    legs: { type: string; labelA: string; labelB: string; status: string; overallVariancePct: number | null; discrepancies: { period: string; description: string }[] }[]
    total_discrepancies: number
    critical_count: number
  } | null>(null)
  const { chatOpen, openChat, closeChat, setCellRef, flagsRefreshRef } = useLayoutContext()

  const workbookName = liveWorkbookName ?? DEMO_NAMES[id] ?? 'Workbook'

  useEffect(() => {
    const name = DEMO_NAMES[id]
    if (name) document.title = `Airbank - ${name}`
    else document.title = 'Airbank - Workbook'
  }, [id])

  const fetchCells = useCallback(async () => {
    if (isDemoWorkbook) return
    setCellsLoading(true)
    setCellsError(null)
    try {
      const [cellsRes, wbRes, flagsRes, compRes, reconRes] = await Promise.all([
        fetch(`/api/workbooks/${id}/cells`),
        fetch(`/api/workbooks/${id}`),
        fetch(`/api/workbooks/${id}/flags`),
        fetch(`/api/workbooks/${id}/completeness`),
        fetch(`/api/workbooks/${id}/reconciliation`),
      ])
      if (cellsRes.status === 401) {
        router.push('/')
        return
      }
      if (cellsRes.ok) {
        const data: LiveCell[] = await cellsRes.json()
        setLiveCells(Array.isArray(data) ? data : [])
      } else if (cellsRes.status !== 404) {
        setCellsError(`Failed to load workbook data (${cellsRes.status})`)
      }
      if (wbRes.ok) {
        const wb = await wbRes.json()
        if (wb.company_name) { setLiveWorkbookName(wb.company_name); document.title = `Airbank - ${wb.company_name}` }
        if (wb.status) setWorkbookStatus(wb.status)
        if (Array.isArray(wb.periods) && wb.periods.length > 0) setWorkbookPeriods(wb.periods)
      }
      if (flagsRes.ok) {
        const flagData = await flagsRes.json()
        setFlags(Array.isArray(flagData) ? flagData : [])
      }
      if (compRes.ok) {
        const compData = await compRes.json()
        if (typeof compData.overall === 'number') setCompleteness(compData.overall)
      }
      if (reconRes.ok) {
        const reconData = await reconRes.json()
        if (reconData.status) setReconciliation(reconData)
      }
    } catch (err) {
      setCellsError(err instanceof Error ? err.message : 'Failed to load workbook data')
    } finally {
      setCellsLoading(false)
    }
  }, [id])

  useEffect(() => {
    fetchCells()
  }, [fetchCells])

  // Scroll to flagged metric after section switch
  useEffect(() => {
    if (!pendingScroll) return
    const timer = setTimeout(() => {
      const el = document.querySelector(`[data-row-key="${pendingScroll}"]`) as HTMLElement | null
      if (el) {
        el.scrollIntoView({ behavior: 'smooth', block: 'center' })
        el.style.transition = 'background-color 0.3s'
        el.style.backgroundColor = '#fef2f2'
        setTimeout(() => { el.style.backgroundColor = '' }, 1800)
      } else if (mainScrollRef.current) {
        mainScrollRef.current.scrollTo({ top: 0, behavior: 'smooth' })
      }
      setPendingScroll(null)
    }, 120)
    return () => clearTimeout(timer)
  }, [pendingScroll, activeSection])

  // Register fetchCells so the AI panel (in layout) can refresh cells after flag changes
  useEffect(() => {
    flagsRefreshRef.current = fetchCells
    return () => { flagsRefreshRef.current = undefined }
  }, [fetchCells, flagsRefreshRef])

  // O(1) lookup map — built once per liveCells update instead of O(n) per cell render
  const liveCellsMap = useMemo(() => {
    const map = new Map<string, LiveCell>()
    for (const cell of liveCells) {
      map.set(`${cell.section}:${cell.row_key}:${cell.period}`, cell)
    }
    return map
  }, [liveCells])

  /** Get a numeric value from live cells. */
  const getLiveValue = useCallback(
    (section: string, rowKey: string, period: string): number | null =>
      liveCellsMap.get(`${section}:${rowKey}:${period}`)?.raw_value ?? null,
    [liveCellsMap]
  )

  /** Build a SourceRef for a cell. */
  const getCellSourceRef = useCallback(
    (section: string, rowKey: string, period: string): SourceRef | null => {
      const cell = liveCellsMap.get(`${section}:${rowKey}:${period}`)
      if (!cell?.source_document) return null
      return {
        documentId: cell.source_document.id,
        documentName: cell.source_document.file_name,
        page: cell.source_page,
        excerpt: cell.source_excerpt ?? '',
        confidence: cell.confidence ?? 0,
      }
    },
    [liveCellsMap]
  )

  /** Get the cell id for persistence. */
  const getCellId = useCallback(
    (section: string, rowKey: string, period: string): string | undefined =>
      liveCellsMap.get(`${section}:${rowKey}:${period}`)?.id,
    [liveCellsMap]
  )

  /** Get flags for a specific cell. */
  const getCellFlags = useCallback(
    (section: string, rowKey: string, period: string): CellFlag[] =>
      liveCellsMap.get(`${section}:${rowKey}:${period}`)?.flags ?? [],
    [liveCellsMap]
  )

  const openFlagCount = useMemo(() => flags.filter(f => !f.resolved_at).length, [flags])

  /** Create a flag from an AuditableCell and refresh. */
  const handleFlagCreate = useCallback(
    async (section: string, rowKey: string, period: string, flag: { title: string; body: string; flag_type: string }) => {
      if (!id || isDemoWorkbook) return
      const cellId = getCellId(section, rowKey, period)
      await fetch(`/api/workbooks/${id}/flags`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          section,
          row_key: rowKey,
          period,
          flag_type: flag.flag_type,
          severity: 'warning',
          title: flag.title,
          body: flag.body,
          cell_id: cellId ?? null,
        }),
      })
      await fetchCells()
    },
    [id, isDemoWorkbook, getCellId, fetchCells]
  )

  const handleViewSource = useCallback((sourceRef: SourceRef) => {
    setViewerSource(sourceRef)
    setViewerOpen(true)
  }, [])

  const handleExport = useCallback(
    async (format: 'excel' | 'pdf' | 'sheets' | 'airtable') => {
      if (isDemoWorkbook) return
      setExporting(true)
      setExportError(null)
      try {
        const res = await fetch(`/api/workbooks/${id}/export`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ format }),
        })
        if (!res.ok) {
          const ct = res.headers.get('content-type') ?? ''
          const body = ct.includes('application/json')
            ? ((await res.json()) as { error?: string }).error ?? 'Export failed'
            : await res.text()
          throw new Error(body)
        }
        if (format === 'excel' || format === 'pdf') {
          const blob = await res.blob()
          const url = URL.createObjectURL(blob)
          const a = document.createElement('a')
          a.href = url
          const ext = format === 'pdf' ? 'pdf' : 'xlsx'
          a.download = `${workbookName.replace(/[^a-z0-9]/gi, '_')}_QoE.${ext}`
          a.click()
          URL.revokeObjectURL(url)
        } else if (format === 'sheets') {
          const data = await res.json()
          window.open(data.url, '_blank')
        }
      } catch (err) {
        setExportError(err instanceof Error ? err.message : String(err))
      } finally {
        setExporting(false)
      }
    },
    [id, isDemoWorkbook, workbookName]
  )

  /**
   * Returns a formatted value for display.
   * - Real workbooks: uses live cell data; shows '—' if not yet extracted (never demo fallback)
   * - Demo workbooks: falls back to hardcoded demo numbers
   */
  const getDisplayValue = useCallback(
    (section: string, rowKey: string, period: string, fallback: number, isPercent = false): string => {
      const live = getLiveValue(section, rowKey, period)
      if (live !== null) return isPercent ? formatPercent(live) : formatCurrency(live)
      if (isDemoWorkbook) return isPercent ? formatPercent(fallback) : formatCurrency(fallback)
      return '—'
    },
    [getLiveValue, isDemoWorkbook]
  )
  // ── End live data layer ─────────────────────────────────────────

  // Sidebar sections
  const sections = [
    { id: 'qoe', name: 'Quality of Earnings', icon: TrendingUp },
    { id: 'income-statement', name: 'Income Statement', icon: FileText },
    { id: 'balance-sheet', name: 'Balance Sheet', icon: Scale },
    { id: 'sales-channel', name: 'Sales by Channel', icon: ShoppingCart },
    { id: 'margins-month', name: 'Margins by Month', icon: Calendar },
    { id: 'proof-cash', name: 'Proof of Cash', icon: Banknote },
    { id: 'working-capital', name: 'Working Capital', icon: DollarSign },
    { id: 'net-debt', name: 'Net Debt & Debt-Like Items', icon: Scale },
    { id: 'customer-concentration', name: 'Customer Concentration', icon: PieChart },
    { id: 'proof-revenue', name: 'Proof of Revenue', icon: FileBarChart },
    { id: 'cash-conversion', name: 'Cash Conversion', icon: TrendingUp },
    { id: 'run-rate', name: 'Run-Rate & Pro Forma', icon: BarChart3 },
    { id: 'cogs-vendors', name: 'COGS Vendors', icon: Package },
    { id: 'testing', name: 'AP/Accrual Testing', icon: ClipboardCheck },
    { id: 'risk-diligence', name: 'Risk & Diligence', icon: Shield },
  ]


  const SECTION_DISPLAY: Record<string, string> = {
    'qoe': 'Quality of Earnings',
    'income-statement': 'Income Statement',
    'balance-sheet': 'Balance Sheet',
    'sales-channel': 'Sales by Channel',
    'margins-month': 'Margins by Month',
    'proof-cash': 'Proof of Cash',
    'working-capital': 'Working Capital',
    'net-debt': 'Net Debt',
    'customer-concentration': 'Customer Concentration',
    'proof-revenue': 'Proof of Revenue',
    'cash-conversion': 'Cash Conversion',
    'run-rate': 'Run-Rate',
    'cogs-vendors': 'COGS Vendors',
    'testing': 'AP Testing',
    'risk-diligence': 'Risk & Diligence',
  }

  const navigateToFlag = (flag: FlagItem) => {
    setFlagsOpen(false)
    const targetSection = SECTION_DISPLAY[flag.section] ? flag.section : 'qoe'
    setActiveSection(targetSection)
    setPendingScroll(flag.row_key)
  }

  // ── Sandbox demo data for sections that otherwise only render live API data ──

  const DEMO_NET_DEBT: Record<string, Record<string, number>> = {
    reported_debt:           { FY20: 1000000,  FY21: 875000,   FY22: 750000,   TTM: 750000   },
    capital_leases:          { FY20: 485000,   FY21: 392000,   FY22: 284000,   TTM: 198000   },
    unpaid_accrued_vacation: { FY20: 218000,   FY21: 265000,   FY22: 312000,   TTM: 338000   },
    aged_accounts_payable:   { FY20: 143000,   FY21: 187000,   FY22: 224000,   TTM: 198000   },
    deferred_revenue:        { FY20: 0,        FY21: 85000,    FY22: 127500,   TTM: 142000   },
    customer_deposits:       { FY20: 89000,    FY21: 112000,   FY22: 138000,   TTM: 152000   },
    total_debt_like_items:   { FY20: 1935000,  FY21: 1916000,  FY22: 1835500,  TTM: 1778000  },
    cash_and_equivalents:    { FY20: 1547293,  FY21: 2183746,  FY22: 2847562,  TTM: 3124000  },
    net_debt:                { FY20: 387707,   FY21: -267746,  FY22: -1012062, TTM: -1346000 },
  }

  const DEMO_CUST_CONC: Record<string, Record<string, number>> = {
    customer_1:           { FY20: 8774990,  FY21: 10708443, FY22: 13051039, TTM: 14205098 },
    customer_2:           { FY20: 5906244,  FY21: 7207606,  FY22: 8784353,  TTM: 9561124  },
    customer_3:           { FY20: 4640620,  FY21: 5663119,  FY22: 6901992,  TTM: 7512312  },
    customer_4:           { FY20: 3374996,  FY21: 4118632,  FY22: 5019630,  TTM: 5463499  },
    customer_5:           { FY20: 2531247,  FY21: 3088974,  FY22: 3764723,  TTM: 4097624  },
    customer_6:           { FY20: 2109373,  FY21: 2574145,  FY22: 3137269,  TTM: 3414687  },
    customer_7:           { FY20: 1687498,  FY21: 2059316,  FY22: 2509815,  TTM: 2731750  },
    customer_8:           { FY20: 1265624,  FY21: 1544487,  FY22: 1882361,  TTM: 2048812  },
    customer_9:           { FY20: 843749,   FY21: 1029658,  FY22: 1254908,  TTM: 1365875  },
    customer_10:          { FY20: 421875,   FY21: 514829,   FY22: 627454,   TTM: 682937   },
    all_other_customers:  { FY20: 10631240, FY21: 12973694, FY22: 15811837, TTM: 17210024 },
    total_revenue:        { FY20: 42187456, FY21: 51482903, FY22: 62745381, TTM: 68293742 },
  }

  const DEMO_PROOF_REV: Record<string, Partial<Record<string, number | null>>> = {
    gl_revenue:           { FY20: 42187456, FY21: 51482903, FY22: 62745381, TTM: 68293742 },
    tax_return_revenue:   { FY20: 42112000, FY21: 51390000, FY22: 62636000, TTM: null     },
    bank_deposit_revenue: { FY20: 42228741, FY21: 51518476, FY22: 62788294, TTM: 68322847 },
    variance_gl_vs_tax:   { FY20: 75456,    FY21: 92903,    FY22: 109381,   TTM: null     },
    variance_gl_vs_bank:  { FY20: -41285,   FY21: -35573,   FY22: -42913,   TTM: -29105   },
    pct_variance_gl_tax:  { FY20: 0.179,    FY21: 0.181,    FY22: 0.174,    TTM: null     },
  }

  const DEMO_CASH_CONV: Record<string, Partial<Record<string, number | null>>> = {
    ebitda:                    { FY20: 4875291, FY21: 6524835, FY22: 7682947, TTM: 8547239 },
    capex:                     { FY20: -318147, FY21: -402837, FY22: -487294, TTM: -512473 },
    change_in_working_capital: { FY20: -184293, FY21: -248471, FY22: -312584, TTM: -284738 },
    taxes_paid_cash:           { FY20: -23185,  FY21: -27038,  FY22: -29583,  TTM: -31256  },
    interest_paid_cash:        { FY20: -42850,  FY21: -52387,  FY22: -61245,  TTM: -65182  },
    free_cash_flow:            { FY20: 4306816, FY21: 5794102, FY22: 6792241, TTM: 7653590 },
    fcf_to_ebitda_ratio:       { FY20: 88.3,    FY21: 88.8,    FY22: 88.4,    TTM: 89.5    },
  }

  const DEMO_RUN_RATE: Record<string, Partial<Record<string, number | null>>> = {
    ttm_revenue_base:       { FY20: null, FY21: null, FY22: 62745381, TTM: 68293742 },
    new_contract_value:     { FY20: null, FY21: null, FY22: 2100000,  TTM: 3420000  },
    full_year_hire_effect:  { FY20: null, FY21: null, FY22: 0,        TTM: -285000  },
    price_increase_effect:  { FY20: null, FY21: null, FY22: 940000,   TTM: 1365875  },
    lost_revenue_adjustment:{ FY20: null, FY21: null, FY22: -820000,  TTM: -1152487 },
    pro_forma_revenue:      { FY20: null, FY21: null, FY22: 64965381, TTM: 71641130 },
    pro_forma_ebitda:       { FY20: null, FY21: null, FY22: 8124947,  TTM: 9213239  },
  }

  const renderContent = () => {
    switch (activeSection) {
      case 'qoe': {
        const qoeTtmRev = getLiveValue('overview', 'total_revenue', 'TTM') ?? 68293742
        const qoeTtmBase = getLiveValue('qoe', 'ebitda_as_defined', 'TTM') ?? 7960816
        const qoeTtmAdj = getLiveValue('qoe', 'diligence_adjusted_ebitda', 'TTM') ?? 8547239
        const qoeNetAdj = qoeTtmAdj - qoeTtmBase
        const qoeAdjMargin = qoeTtmRev > 0 ? (qoeTtmAdj / qoeTtmRev) * 100 : 12.5

        // Build TTM EBITDA waterfall — range bar approach: [start, end] per bar
        type WfItem = { name: string; value: number; isBase?: boolean; isEnd?: boolean }
        type WfPoint = { name: string; range: [number, number]; rawValue: number; isBase: boolean; isEnd: boolean }
        const qoeWfItems: WfItem[] = [
          { name: 'EBITDA as Defined', value: qoeTtmBase, isBase: true },
          { name: 'a) Charitable', value: getLiveValue('qoe', 'adj_charitable_contributions', 'TTM') ?? 315920 },
          { name: 'b) Owner Tax', value: getLiveValue('qoe', 'adj_owner_tax_legal', 'TTM') ?? 22450 },
          { name: 'c) Excess Comp', value: getLiveValue('qoe', 'adj_excess_owner_comp', 'TTM') ?? 248190 },
          { name: 'd) Personal Exp', value: getLiveValue('qoe', 'adj_personal_expenses', 'TTM') ?? 39863 },
          { name: 'e) Prof. Fees', value: getLiveValue('qoe', 'adj_professional_fees_onetime', 'TTM') ?? 68200 },
          { name: 'f) Exec. Search', value: getLiveValue('qoe', 'adj_executive_search', 'TTM') ?? 52000 },
          { name: 'g) Relocation', value: getLiveValue('qoe', 'adj_facility_relocation', 'TTM') ?? 41500 },
          { name: 'h) Severance', value: getLiveValue('qoe', 'adj_severance', 'TTM') ?? 92300 },
          { name: 'i) Above-Mkt Rent', value: getLiveValue('qoe', 'adj_above_market_rent', 'TTM') ?? -168900 },
          { name: 'j) Owner Comp', value: getLiveValue('qoe', 'adj_normalize_owner_comp', 'TTM') ?? -425900 },
          { name: 'Adj. EBITDA', value: 0, isEnd: true },
        ]
        let qoeWfRunning = 0
        const qoeWfData: WfPoint[] = qoeWfItems.map((item) => {
          if (item.isBase) { qoeWfRunning = item.value; return { name: item.name, range: [0, item.value], rawValue: item.value, isBase: true, isEnd: false } }
          if (item.isEnd) { return { name: item.name, range: [0, qoeWfRunning], rawValue: qoeWfRunning, isBase: false, isEnd: true } }
          const lo = item.value >= 0 ? qoeWfRunning : qoeWfRunning + item.value
          const hi = item.value >= 0 ? qoeWfRunning + item.value : qoeWfRunning
          const pt: WfPoint = { name: item.name, range: [lo, hi], rawValue: item.value, isBase: false, isEnd: false }
          qoeWfRunning += item.value
          return pt
        })
        const qoeWfYMin = qoeTtmBase * 0.97
        const qoeWfYMax = Math.max(...qoeWfData.map((p) => p.range[1])) * 1.02

        // Margin trend
        const qoeMarginTrend = [
          { period: 'FY20', unadj: (4630517 / 42187456) * 100, adj: (4875291 / 42187456) * 100 },
          { period: 'FY21', unadj: (6177949 / 51482903) * 100, adj: (6524835 / 51482903) * 100 },
          { period: 'FY22', unadj: (7217823 / 62745381) * 100, adj: (7682947 / 62745381) * 100 },
          { period: 'TTM',  unadj: qoeTtmRev > 0 ? (qoeTtmBase / qoeTtmRev) * 100 : 11.7, adj: qoeAdjMargin },
        ]

        return (
          <div>
            <h2 className="text-2xl font-bold mb-6">Quality of Earnings</h2>
            {cellsLoading && (
              <div className="flex items-center gap-2 text-sm text-muted-foreground mb-4">
                <Loader2 className="h-4 w-4 animate-spin" />
                Loading live data...
              </div>
            )}

            {/* KPI Cards */}
            <div className="grid grid-cols-4 gap-3 mb-6">
              <div className="rounded-lg border bg-card p-4">
                <p className="text-xs text-muted-foreground mb-1">TTM Revenue</p>
                <p className="text-xl font-bold">{formatCurrency(qoeTtmRev)}</p>
              </div>
              <div className="rounded-lg border bg-card p-4">
                <p className="text-xs text-muted-foreground mb-1">EBITDA as Defined</p>
                <p className="text-xl font-bold">{formatCurrency(qoeTtmBase)}</p>
                <p className="text-xs text-muted-foreground mt-1">{formatPercent(qoeTtmRev > 0 ? (qoeTtmBase / qoeTtmRev) * 100 : 11.7)} margin</p>
              </div>
              <div className="rounded-lg border bg-card p-4">
                <p className="text-xs text-muted-foreground mb-1">Adj. EBITDA</p>
                <p className="text-xl font-bold text-[#1E3A5F]">{formatCurrency(qoeTtmAdj)}</p>
                <p className="text-xs text-muted-foreground mt-1">{formatPercent(qoeAdjMargin)} margin</p>
              </div>
              <div className="rounded-lg border bg-card p-4">
                <p className="text-xs text-muted-foreground mb-1">Total Net Adjustments</p>
                <p className="text-xl font-bold text-emerald-600">+{formatCurrency(qoeNetAdj)}</p>
                <p className="text-xs text-muted-foreground mt-1">10 discrete add-backs</p>
              </div>
            </div>

            {/* Charts row */}
            <div className="grid grid-cols-2 gap-6 mb-8">
              {/* EBITDA Waterfall */}
              <div className="rounded-lg border bg-card p-4">
                <h3 className="text-sm font-semibold mb-3">EBITDA Bridge — TTM</h3>
                <div className="h-64">
                  <ResponsiveContainer width="100%" height="100%">
                    <BarChart data={qoeWfData} margin={{ top: 5, right: 10, left: 20, bottom: 44 }} barCategoryGap="20%">
                      <CartesianGrid strokeDasharray="3 3" stroke="#f0f0f0" vertical={false} />
                      <XAxis dataKey="name" tick={{ fontSize: 9 }} angle={-40} textAnchor="end" interval={0} />
                      <YAxis domain={[qoeWfYMin, qoeWfYMax]} tickFormatter={(v) => `$${(v / 1e6).toFixed(1)}M`} tick={{ fontSize: 9 }} width={55} />
                      <Tooltip
                        content={({ active, payload, label }) => {
                          if (!active || !payload?.length) return null
                          const d = (payload[0] as { payload: WfPoint }).payload
                          return (
                            <div className="bg-white border border-gray-200 rounded px-3 py-2 text-xs shadow">
                              <div className="font-semibold mb-1">{label}</div>
                              <div>{d.isBase || d.isEnd ? formatCurrency(d.range[1]) : `${d.rawValue >= 0 ? '+' : ''}${formatCurrency(d.rawValue)}`}</div>
                            </div>
                          )
                        }}
                      />
                      <Bar dataKey="range">
                        {qoeWfData.map((entry, idx) => (
                          <Cell key={idx} fill={entry.isBase || entry.isEnd ? '#1E3A5F' : entry.rawValue >= 0 ? '#10b981' : '#ef4444'} />
                        ))}
                      </Bar>
                    </BarChart>
                  </ResponsiveContainer>
                </div>
              </div>

              {/* Margin Trend */}
              <div className="rounded-lg border bg-card p-4">
                <h3 className="text-sm font-semibold mb-3">EBITDA Margin Trend</h3>
                <div className="h-64">
                  <ResponsiveContainer width="100%" height="100%">
                    <LineChart data={qoeMarginTrend} margin={{ top: 5, right: 20, left: 10, bottom: 5 }}>
                      <CartesianGrid strokeDasharray="3 3" stroke="#f0f0f0" vertical={false} />
                      <XAxis dataKey="period" tick={{ fontSize: 11 }} />
                      <YAxis tickFormatter={(v: number) => `${v.toFixed(1)}%`} tick={{ fontSize: 10 }} width={45} domain={[9, 15]} />
                      <Tooltip formatter={(v: unknown) => [`${typeof v === 'number' ? v.toFixed(1) : v}%`]} />
                      <Legend iconSize={10} wrapperStyle={{ fontSize: 11 }} formatter={(v: string) => v === 'adj' ? 'Adj. EBITDA Margin' : 'EBITDA as Defined Margin'} />
                      <Line type="monotone" dataKey="unadj" name="unadj" stroke="#94a3b8" strokeWidth={2} dot={{ r: 4 }} />
                      <Line type="monotone" dataKey="adj" name="adj" stroke="#1E3A5F" strokeWidth={2} dot={{ r: 4 }} />
                    </LineChart>
                  </ResponsiveContainer>
                </div>
              </div>
            </div>

            {/* Full Adjustment Schedule */}
            <h3 className="text-base font-semibold mb-3">Full Adjustment Schedule</h3>
            <div className="rounded-md border overflow-x-auto">
              <Table>
                <TableHeader>
                  <TableRow>
                    <TableHead className="font-semibold w-[300px]">Line Item</TableHead>
                    <TableHead className="font-semibold text-right w-[150px]">FY20</TableHead>
                    <TableHead className="font-semibold text-right w-[150px]">FY21</TableHead>
                    <TableHead className="font-semibold text-right w-[150px]">FY22</TableHead>
                    <TableHead className="font-semibold text-right w-[150px]">TTM Jan-23</TableHead>
                  </TableRow>
                </TableHeader>
                <TableBody>
                  {qoeData.map((row, idx) => {
                    const rowKey = row.item.toLowerCase().replace(/[^a-z0-9]+/g, '_')
                    return (
                    <TableRow key={idx} data-row-key={rowKey} className={row.isBold ? 'bg-muted/50 font-semibold' : ''}>
                      <TableCell className={row.isAdjustment ? 'pl-8 italic text-sm' : row.isBold ? 'font-semibold' : ''}>
                        {row.item}
                      </TableCell>
                      {(['FY20', 'FY21', 'FY22', 'TTM'] as const).map((period, pi) => {
                        const fallbackVal = pi === 0 ? row.fy20 : pi === 1 ? row.fy21 : pi === 2 ? row.fy22 : row.ttm
                        const displayVal = getDisplayValue('qoe', rowKey, period, fallbackVal)
                        const srcRef = getCellSourceRef('qoe', rowKey, period)
                        const cellId = getCellId('qoe', rowKey, period)
                        return (
                          <TableCell key={period} className="text-right font-mono">
                            <AuditableCell
                              value={displayVal}
                              source={row.source}
                              sourceRef={srcRef}
                              cellId={cellId}
                              workbookId={isDemoWorkbook ? undefined : id}
                              isEditable={!row.isBold}
                              onViewSource={handleViewSource}
                            />
                          </TableCell>
                        )
                      })}
                    </TableRow>
                    )
                  })}
                </TableBody>
              </Table>
            </div>
          </div>
        )
      }

      case 'income-statement': {
        const isChartData = [
          { period: 'FY20', revenue: 42187456, grossProfit: 23962476, ebitda: 4630517, netIncome: 4558429 },
          { period: 'FY21', revenue: 51482903, grossProfit: 29242289, ebitda: 6177949, netIncome: 6091874 },
          { period: 'FY22', revenue: 62745381, grossProfit: 35639373, ebitda: 7217823, netIncome: 7120305 },
          { period: 'TTM',  revenue: 68293742, grossProfit: 38790844, ebitda: 7960816, netIncome: 7858378 },
        ]
        const isMarginData = isChartData.map(d => ({
          period: d.period,
          grossMargin: parseFloat(((d.grossProfit / d.revenue) * 100).toFixed(1)),
          ebitdaMargin: parseFloat(((d.ebitda / d.revenue) * 100).toFixed(1)),
          netMargin: parseFloat(((d.netIncome / d.revenue) * 100).toFixed(1)),
        }))
        return (
          <div>
            <h2 className="text-2xl font-bold mb-6">Income Statement - Adjusted</h2>
            {cellsLoading && (
              <div className="flex items-center gap-2 text-sm text-muted-foreground mb-4">
                <Loader2 className="h-4 w-4 animate-spin" />Loading live data...
              </div>
            )}
            {/* Charts */}
            <div className="grid grid-cols-2 gap-4 mb-8">
              <div className="rounded-lg border bg-card p-4">
                <h3 className="text-sm font-semibold mb-1">Revenue vs. Gross Profit</h3>
                <p className="text-xs text-muted-foreground mb-3">FY20 — TTM ($ millions)</p>
                <div className="h-52">
                  <ResponsiveContainer width="100%" height="100%">
                    <BarChart data={isChartData} margin={{ top: 4, right: 8, left: 28, bottom: 4 }} barGap={4}>
                      <CartesianGrid strokeDasharray="3 3" stroke="#f0f0f0" vertical={false} />
                      <XAxis dataKey="period" tick={{ fontSize: 11 }} />
                      <YAxis tickFormatter={v => `$${(v/1e6).toFixed(0)}M`} tick={{ fontSize: 10 }} width={48} />
                      <Tooltip formatter={(v: number | undefined) => formatCurrency(v ?? 0)} />
                      <Legend wrapperStyle={{ fontSize: 11 }} />
                      <Bar dataKey="revenue" name="Revenue" fill="#2563EB" radius={[3,3,0,0]} />
                      <Bar dataKey="grossProfit" name="Gross Profit" fill="#93C5FD" radius={[3,3,0,0]} />
                    </BarChart>
                  </ResponsiveContainer>
                </div>
              </div>
              <div className="rounded-lg border bg-card p-4">
                <h3 className="text-sm font-semibold mb-1">Margin Trends</h3>
                <p className="text-xs text-muted-foreground mb-3">Gross, EBITDA & Net margin %</p>
                <div className="h-52">
                  <ResponsiveContainer width="100%" height="100%">
                    <LineChart data={isMarginData} margin={{ top: 4, right: 16, left: 10, bottom: 4 }}>
                      <CartesianGrid strokeDasharray="3 3" stroke="#f0f0f0" vertical={false} />
                      <XAxis dataKey="period" tick={{ fontSize: 11 }} />
                      <YAxis tickFormatter={v => `${v}%`} tick={{ fontSize: 10 }} domain={[0, 70]} width={38} />
                      <Tooltip formatter={(v: number | undefined) => `${(v ?? 0).toFixed(1)}%`} />
                      <Legend wrapperStyle={{ fontSize: 11 }} />
                      <Line type="monotone" dataKey="grossMargin" name="Gross Margin" stroke="#2563EB" strokeWidth={2} dot={{ r: 3 }} />
                      <Line type="monotone" dataKey="ebitdaMargin" name="EBITDA Margin" stroke="#7C3AED" strokeWidth={2} dot={{ r: 3 }} />
                      <Line type="monotone" dataKey="netMargin" name="Net Margin" stroke="#64748B" strokeWidth={2} dot={{ r: 3 }} />
                    </LineChart>
                  </ResponsiveContainer>
                </div>
              </div>
            </div>
            <div className="rounded-md border overflow-x-auto">
              <Table>
                <TableHeader>
                  <TableRow>
                    <TableHead className="font-semibold w-[300px]">Line Item</TableHead>
                    <TableHead className="font-semibold text-right w-[130px]">FY20</TableHead>
                    <TableHead className="font-semibold text-right w-[130px]">FY21</TableHead>
                    <TableHead className="font-semibold text-right w-[130px]">FY22</TableHead>
                    <TableHead className="font-semibold text-right w-[130px]">TTM Jan-23</TableHead>
                    <TableHead className="font-semibold text-right w-[80px]">% of Sales</TableHead>
                  </TableRow>
                </TableHeader>
                <TableBody>
                  {incomeStatementData.map((row, idx) => {
                    const rowKey = row.item.toLowerCase().replace(/[^a-z0-9]+/g, '_')
                    return (
                      <TableRow key={idx} data-row-key={rowKey} className={row.isBold ? 'bg-muted/50 font-semibold' : ''}>
                        <TableCell className={row.isBold ? 'font-semibold' : ''}>{row.item}</TableCell>
                        {(['FY20', 'FY21', 'FY22', 'TTM'] as const).map((period, pi) => {
                          const fallback = pi === 0 ? row.fy20 : pi === 1 ? row.fy21 : pi === 2 ? row.fy22 : row.ttm
                          const val = getLiveValue('income-statement', rowKey, period) ?? fallback
                          const srcRef = getCellSourceRef('income-statement', rowKey, period)
                          const cellId = getCellId('income-statement', rowKey, period)
                          return (
                            <TableCell key={period} className="text-right font-mono text-sm">
                              <AuditableCell
                                value={formatCurrency(val)}
                                source=""
                                sourceRef={srcRef}
                                cellId={cellId}
                                workbookId={isDemoWorkbook ? undefined : id}
                                isEditable={!row.isBold}
                                onViewSource={handleViewSource}
                              />
                            </TableCell>
                          )
                        })}
                        <TableCell className="text-right font-mono text-sm text-muted-foreground">
                          {row.pct}%
                        </TableCell>
                      </TableRow>
                    )
                  })}
                </TableBody>
              </Table>
            </div>
          </div>
        )
      }

      case 'balance-sheet': {
        const bsChartData = [
          { period: 'FY20', currentAssets: 3909806, fixedAssets: 37836, intangibles: 18400, currentLiab: 1484292, ltDebt: 875000, equity: 1606750 },
          { period: 'FY21', currentAssets: 5102399, fixedAssets: 50416, intangibles: 20200, currentLiab: 1820451, ltDebt: 750000, equity: 2602564 },
          { period: 'FY22', currentAssets: 6393106, fixedAssets: 69696, intangibles: 22100, currentLiab: 2198474, ltDebt: 625000, equity: 3661428 },
        ]
        return (
          <div>
            <h2 className="text-2xl font-bold mb-6">Balance Sheet - Adjusted</h2>
            {cellsLoading && (
              <div className="flex items-center gap-2 text-sm text-muted-foreground mb-4">
                <Loader2 className="h-4 w-4 animate-spin" />Loading live data...
              </div>
            )}
            {/* Charts */}
            <div className="grid grid-cols-2 gap-4 mb-8">
              <div className="rounded-lg border bg-card p-4">
                <h3 className="text-sm font-semibold mb-1">Asset Composition</h3>
                <p className="text-xs text-muted-foreground mb-3">Stacked by category ($ millions)</p>
                <div className="h-52">
                  <ResponsiveContainer width="100%" height="100%">
                    <BarChart data={bsChartData} margin={{ top: 4, right: 8, left: 28, bottom: 4 }}>
                      <CartesianGrid strokeDasharray="3 3" stroke="#f0f0f0" vertical={false} />
                      <XAxis dataKey="period" tick={{ fontSize: 11 }} />
                      <YAxis tickFormatter={v => `$${(v/1e6).toFixed(1)}M`} tick={{ fontSize: 10 }} width={48} />
                      <Tooltip formatter={(v: number | undefined) => formatCurrency(v ?? 0)} />
                      <Legend wrapperStyle={{ fontSize: 11 }} />
                      <Bar dataKey="currentAssets" name="Current Assets" stackId="a" fill="#2563EB" />
                      <Bar dataKey="fixedAssets" name="Fixed Assets" stackId="a" fill="#7C3AED" />
                      <Bar dataKey="intangibles" name="Intangibles" stackId="a" fill="#93C5FD" radius={[3,3,0,0]} />
                    </BarChart>
                  </ResponsiveContainer>
                </div>
              </div>
              <div className="rounded-lg border bg-card p-4">
                <h3 className="text-sm font-semibold mb-1">Liabilities vs. Equity</h3>
                <p className="text-xs text-muted-foreground mb-3">Capital structure over time</p>
                <div className="h-52">
                  <ResponsiveContainer width="100%" height="100%">
                    <BarChart data={bsChartData} margin={{ top: 4, right: 8, left: 28, bottom: 4 }}>
                      <CartesianGrid strokeDasharray="3 3" stroke="#f0f0f0" vertical={false} />
                      <XAxis dataKey="period" tick={{ fontSize: 11 }} />
                      <YAxis tickFormatter={v => `$${(v/1e6).toFixed(1)}M`} tick={{ fontSize: 10 }} width={48} />
                      <Tooltip formatter={(v: number | undefined) => formatCurrency(v ?? 0)} />
                      <Legend wrapperStyle={{ fontSize: 11 }} />
                      <Bar dataKey="currentLiab" name="Current Liabilities" stackId="b" fill="#F59E0B" />
                      <Bar dataKey="ltDebt" name="LT Debt" stackId="b" fill="#EF4444" />
                      <Bar dataKey="equity" name="Equity" stackId="b" fill="#10B981" radius={[3,3,0,0]} />
                    </BarChart>
                  </ResponsiveContainer>
                </div>
              </div>
            </div>
            <div className="rounded-md border overflow-x-auto">
              <Table>
                <TableHeader>
                  <TableRow>
                    <TableHead className="font-semibold w-[300px]">Line Item</TableHead>
                    <TableHead className="font-semibold text-right w-[180px]">Jan 31, 2021</TableHead>
                    <TableHead className="font-semibold text-right w-[180px]">Jan 31, 2022</TableHead>
                    <TableHead className="font-semibold text-right w-[180px]">Jan 31, 2023</TableHead>
                  </TableRow>
                </TableHeader>
                <TableBody>
                  {balanceSheetData.map((row, idx) => {
                    const rowKey = row.item.toLowerCase().replace(/[^a-z0-9]+/g, '_')
                    const bsPeriods = ['FY20', 'FY21', 'FY22'] as const
                    const fallbacks = [row.jan21, row.jan22, row.jan23]
                    return (
                      <TableRow key={idx} data-row-key={rowKey} className={row.isBold ? 'bg-muted/50 font-semibold' : ''}>
                        <TableCell className={row.isBold ? 'font-semibold' : ''}>{row.item}</TableCell>
                        {bsPeriods.map((period, pi) => {
                          const displayVal = getDisplayValue('balance-sheet', rowKey, period, fallbacks[pi])
                          const srcRef = getCellSourceRef('balance-sheet', rowKey, period)
                          const cellId = getCellId('balance-sheet', rowKey, period)
                          return (
                            <TableCell key={period} className="text-right font-mono">
                              <AuditableCell
                                value={displayVal}
                                source={row.source}
                                sourceRef={srcRef}
                                cellId={cellId}
                                workbookId={isDemoWorkbook ? undefined : id}
                                isEditable={!row.isBold}
                                onViewSource={handleViewSource}
                              />
                            </TableCell>
                          )
                        })}
                      </TableRow>
                    )
                  })}
                </TableBody>
              </Table>
            </div>
          </div>
        )
      }

      case 'sales-channel':
        return (
          <div>
            <h2 className="text-2xl font-bold mb-6">Sales by Channel</h2>
            <p className="text-sm text-muted-foreground mb-4">Product mix and SKU analysis for TTM Jan-23</p>
            {/* Chart */}
            <div className="grid grid-cols-2 gap-4 mb-8">
              <div className="rounded-lg border bg-card p-4">
                <h3 className="text-sm font-semibold mb-1">Revenue by Product — TTM</h3>
                <p className="text-xs text-muted-foreground mb-3">Top 8 products + other ($ millions)</p>
                <div className="h-56">
                  <ResponsiveContainer width="100%" height="100%">
                    <BarChart layout="vertical" data={salesChannelData.filter(r => !r.isBold).slice(0,9).map(r => ({ name: r.product.length > 22 ? r.product.slice(0,22)+'…' : r.product, revenue: r.revenue }))} margin={{ top: 4, right: 16, left: 4, bottom: 4 }}>
                      <CartesianGrid strokeDasharray="3 3" stroke="#f0f0f0" horizontal={false} />
                      <XAxis type="number" tickFormatter={v => `$${(v/1e6).toFixed(0)}M`} tick={{ fontSize: 10 }} />
                      <YAxis type="category" dataKey="name" tick={{ fontSize: 10 }} width={124} />
                      <Tooltip formatter={(v: number | undefined) => formatCurrency(v ?? 0)} />
                      <Bar dataKey="revenue" name="Revenue" fill="#2563EB" radius={[0,3,3,0]} />
                    </BarChart>
                  </ResponsiveContainer>
                </div>
              </div>
              <div className="rounded-lg border bg-card p-4">
                <h3 className="text-sm font-semibold mb-1">Revenue Mix — TTM</h3>
                <p className="text-xs text-muted-foreground mb-3">Share of total revenue by product</p>
                <div className="h-56 flex items-center justify-center">
                  <ResponsiveContainer width="100%" height="100%">
                    <RechartsPie data={salesChannelData.filter(r => !r.isBold).map(r => ({ name: r.product.split(' ').slice(0,2).join(' '), value: r.revenue }))} cx="50%" cy="50%" innerRadius={55} outerRadius={88}>
                      {salesChannelData.filter(r => !r.isBold).map((_, i) => (
                        <Cell key={i} fill={['#2563EB','#7C3AED','#0EA5E9','#10B981','#F59E0B','#EF4444','#64748B','#A78BFA','#34D399','#FCA5A5'][i % 10]} />
                      ))}
                      <Tooltip formatter={(v: number | undefined) => formatCurrency(v ?? 0)} />
                    </RechartsPie>
                  </ResponsiveContainer>
                </div>
              </div>
            </div>
            <div className="rounded-md border overflow-x-auto">
              <Table>
                <TableHeader>
                  <TableRow>
                    <TableHead className="font-semibold w-[200px]">Product/SKU</TableHead>
                    <TableHead className="font-semibold text-right w-[150px]">Revenue</TableHead>
                    <TableHead className="font-semibold text-right w-[100px]">% of Total</TableHead>
                    <TableHead className="font-semibold text-right w-[120px]">Qty Sold</TableHead>
                    <TableHead className="font-semibold text-right w-[120px]">Avg Price</TableHead>
                  </TableRow>
                </TableHeader>
                <TableBody>
                  {salesChannelData.map((row, idx) => {
                    const rowKey = row.product.toLowerCase().replace(/[^a-z0-9]+/g, '_').slice(0, 40)
                    const displayVal = getDisplayValue('sales-channel', rowKey, 'TTM', row.revenue)
                    const srcRef = getCellSourceRef('sales-channel', rowKey, 'TTM')
                    const cellId = getCellId('sales-channel', rowKey, 'TTM')
                    return (
                      <TableRow key={idx} className={row.isBold ? 'bg-muted/50 font-semibold' : ''}>
                        <TableCell className={row.isBold ? 'font-semibold' : ''}>{row.product}</TableCell>
                        <TableCell className="text-right font-mono">
                          <AuditableCell
                            value={displayVal}
                            source={row.source}
                            sourceRef={srcRef}
                            cellId={cellId}
                            workbookId={isDemoWorkbook ? undefined : id}
                            onViewSource={handleViewSource}
                          />
                        </TableCell>
                        <TableCell className="text-right font-mono">{formatPercent(row.pct)}</TableCell>
                        <TableCell className="text-right font-mono">{row.qty.toLocaleString()}</TableCell>
                        <TableCell className="text-right font-mono">{formatCurrency(row.avgPrice)}</TableCell>
                      </TableRow>
                    )
                  })}
                </TableBody>
              </Table>
            </div>
          </div>
        )

      case 'margins-month':
        return (
          <div>
            <h2 className="text-2xl font-bold mb-6">Margins by Month</h2>
            <p className="text-sm text-muted-foreground mb-4">Monthly profitability analysis - Last 12 months</p>
            {/* Charts */}
            <div className="rounded-lg border bg-card p-4 mb-8">
              <h3 className="text-sm font-semibold mb-1">Monthly Revenue & Margin %</h3>
              <p className="text-xs text-muted-foreground mb-3">Bars = Revenue · Lines = Gross & EBITDA margin %</p>
              <div className="h-60">
                <ResponsiveContainer width="100%" height="100%">
                  <ComposedChart data={marginsByMonthData.map(m => ({ month: m.month, revenue: m.revenue, cogs: m.cogs, ebitda: m.revenue - m.cogs - m.opex, grossMarginPct: parseFloat(((m.revenue - m.cogs) / m.revenue * 100).toFixed(1)), ebitdaMarginPct: parseFloat(((m.revenue - m.cogs - m.opex) / m.revenue * 100).toFixed(1)) }))} margin={{ top: 4, right: 40, left: 30, bottom: 4 }}>
                    <CartesianGrid strokeDasharray="3 3" stroke="#f0f0f0" vertical={false} />
                    <XAxis dataKey="month" tick={{ fontSize: 11 }} />
                    <YAxis yAxisId="left" tickFormatter={v => `$${(v/1e6).toFixed(1)}M`} tick={{ fontSize: 10 }} width={48} />
                    <YAxis yAxisId="right" orientation="right" tickFormatter={v => `${v}%`} tick={{ fontSize: 10 }} domain={[0, 70]} width={36} />
                    <Tooltip formatter={(v: number | undefined, name: string | undefined) => (name ?? '').includes('%') ? `${v ?? 0}%` : formatCurrency(v ?? 0)} />
                    <Legend wrapperStyle={{ fontSize: 11 }} />
                    <Bar yAxisId="left" dataKey="revenue" name="Revenue" fill="#2563EB" opacity={0.85} radius={[2,2,0,0]} />
                    <Line yAxisId="right" type="monotone" dataKey="grossMarginPct" name="Gross Margin %" stroke="#10B981" strokeWidth={2} dot={{ r: 2 }} />
                    <Line yAxisId="right" type="monotone" dataKey="ebitdaMarginPct" name="EBITDA Margin %" stroke="#F59E0B" strokeWidth={2} dot={{ r: 2 }} />
                  </ComposedChart>
                </ResponsiveContainer>
              </div>
            </div>
            <div className="rounded-md border overflow-x-auto">
              <Table>
                <TableHeader>
                  <TableRow>
                    <TableHead className="font-semibold w-[100px]">Month</TableHead>
                    <TableHead className="font-semibold text-right w-[150px]">Revenue</TableHead>
                    <TableHead className="font-semibold text-right w-[150px]">COGS</TableHead>
                    <TableHead className="font-semibold text-right w-[150px]">OpEx</TableHead>
                    <TableHead className="font-semibold text-right w-[150px]">Gross Margin</TableHead>
                    <TableHead className="font-semibold text-right w-[100px]">GM %</TableHead>
                    <TableHead className="font-semibold text-right w-[150px]">Net Margin</TableHead>
                    <TableHead className="font-semibold text-right w-[100px]">NM %</TableHead>
                  </TableRow>
                </TableHeader>
                <TableBody>
                  {marginsByMonthData.map((row, idx) => {
                    // Use live data for revenue/cogs/opex if available, keyed by month
                    const liveRev = getLiveValue('margins-month', 'revenue', row.month) ?? row.revenue
                    const liveCogs = getLiveValue('margins-month', 'cogs', row.month) ?? row.cogs
                    const liveOpex = getLiveValue('margins-month', 'opex', row.month) ?? row.opex
                    const grossMargin = liveRev - liveCogs
                    const netMargin = liveRev - liveCogs - liveOpex
                    const gmPct = liveRev > 0 ? (grossMargin / liveRev) * 100 : 0
                    const nmPct = liveRev > 0 ? (netMargin / liveRev) * 100 : 0
                    return (
                      <TableRow key={idx}>
                        <TableCell className="font-medium">{row.month}</TableCell>
                        <TableCell className="text-right font-mono text-sm">{formatCurrency(liveRev)}</TableCell>
                        <TableCell className="text-right font-mono text-sm">{formatCurrency(liveCogs)}</TableCell>
                        <TableCell className="text-right font-mono text-sm">{formatCurrency(liveOpex)}</TableCell>
                        <TableCell className="text-right font-mono text-sm">{formatCurrency(grossMargin)}</TableCell>
                        <TableCell className="text-right font-mono text-sm">{formatPercent(gmPct)}</TableCell>
                        <TableCell className="text-right font-mono text-sm">{formatCurrency(netMargin)}</TableCell>
                        <TableCell className="text-right font-mono text-sm">{formatPercent(nmPct)}</TableCell>
                      </TableRow>
                    )
                  })}
                </TableBody>
              </Table>
            </div>
          </div>
        )

      case 'proof-cash':
        return (
          <div>
            <h2 className="text-2xl font-bold mb-6">Proof of Cash</h2>
            <p className="text-sm text-muted-foreground mb-4">Bank deposit reconciliation to reported revenue</p>
            {/* Chart */}
            <div className="rounded-lg border bg-card p-4 mb-8">
              <h3 className="text-sm font-semibold mb-1">Bank Deposits vs. GL Revenue — Monthly</h3>
              <p className="text-xs text-muted-foreground mb-3">Variance shown as line · Near-zero = clean reconciliation</p>
              <div className="h-56">
                <ResponsiveContainer width="100%" height="100%">
                  <ComposedChart data={proofOfCashData.map(m => ({ month: m.month, deposits: m.deposits, revenueGL: m.revenueGL, variance: m.deposits - m.beginningAR + m.endingAR - m.nonRevDeposits - m.revenueGL }))} margin={{ top: 4, right: 40, left: 30, bottom: 4 }}>
                    <CartesianGrid strokeDasharray="3 3" stroke="#f0f0f0" vertical={false} />
                    <XAxis dataKey="month" tick={{ fontSize: 11 }} />
                    <YAxis yAxisId="left" tickFormatter={v => `$${(v/1e6).toFixed(1)}M`} tick={{ fontSize: 10 }} width={48} />
                    <YAxis yAxisId="right" orientation="right" tickFormatter={v => `$${(v/1000).toFixed(0)}K`} tick={{ fontSize: 10 }} width={44} />
                    <Tooltip formatter={(v: number | undefined) => formatCurrency(v ?? 0)} />
                    <Legend wrapperStyle={{ fontSize: 11 }} />
                    <Bar yAxisId="left" dataKey="deposits" name="Bank Deposits" fill="#2563EB" opacity={0.8} radius={[2,2,0,0]} />
                    <Bar yAxisId="left" dataKey="revenueGL" name="GL Revenue" fill="#93C5FD" opacity={0.8} radius={[2,2,0,0]} />
                    <Line yAxisId="right" type="monotone" dataKey="variance" name="Variance" stroke="#EF4444" strokeWidth={2} dot={{ r: 3 }} />
                  </ComposedChart>
                </ResponsiveContainer>
              </div>
            </div>
            <div className="rounded-md border overflow-x-auto">
              <Table>
                <TableHeader>
                  <TableRow>
                    <TableHead className="font-semibold w-[100px]">Month</TableHead>
                    <TableHead className="font-semibold text-right w-[140px]">Bank Deposits</TableHead>
                    <TableHead className="font-semibold text-right w-[140px]">Less: Beg AR</TableHead>
                    <TableHead className="font-semibold text-right w-[140px]">Add: End AR</TableHead>
                    <TableHead className="font-semibold text-right w-[140px]">Less: Non-Rev</TableHead>
                    <TableHead className="font-semibold text-right w-[140px]">Revenue GL</TableHead>
                    <TableHead className="font-semibold text-right w-[120px]">Variance</TableHead>
                  </TableRow>
                </TableHeader>
                <TableBody>
                  {proofOfCashData.map((row, idx) => {
                    const liveDeposits = getLiveValue('proof-cash', 'bank_deposits', row.month) ?? row.deposits
                    const liveBegAR = getLiveValue('proof-cash', 'beginning_ar', row.month) ?? row.beginningAR
                    const liveEndAR = getLiveValue('proof-cash', 'ending_ar', row.month) ?? row.endingAR
                    const liveNonRev = getLiveValue('proof-cash', 'non_rev_deposits', row.month) ?? row.nonRevDeposits
                    const liveRevGL = getLiveValue('proof-cash', 'revenue_gl', row.month) ?? row.revenueGL
                    const calculated = liveDeposits - liveBegAR + liveEndAR - liveNonRev
                    const variance = calculated - liveRevGL
                    return (
                      <TableRow key={idx}>
                        <TableCell className="font-medium">{row.month}</TableCell>
                        <TableCell className="text-right font-mono text-sm">{formatCurrency(liveDeposits)}</TableCell>
                        <TableCell className="text-right font-mono text-sm">{formatCurrency(liveBegAR)}</TableCell>
                        <TableCell className="text-right font-mono text-sm">{formatCurrency(liveEndAR)}</TableCell>
                        <TableCell className="text-right font-mono text-sm">{formatCurrency(liveNonRev)}</TableCell>
                        <TableCell className="text-right font-mono text-sm">{formatCurrency(liveRevGL)}</TableCell>
                        <TableCell className="text-right font-mono text-sm">
                          <span className={Math.abs(variance) > 50000 ? 'text-red-500' : 'text-green-600'}>
                            {formatCurrency(variance)}
                          </span>
                        </TableCell>
                      </TableRow>
                    )
                  })}
                </TableBody>
              </Table>
            </div>
          </div>
        )

      case 'working-capital':
        return (
          <div>
            <h2 className="text-2xl font-bold mb-6">Adjusted Net Working Capital Analysis</h2>
            <p className="text-sm text-muted-foreground mb-4">Working capital components with DSO, DPO, and DOH metrics</p>
            {/* Chart */}
            <div className="grid grid-cols-2 gap-4 mb-8">
              <div className="rounded-lg border bg-card p-4">
                <h3 className="text-sm font-semibold mb-1">NWC Components (TTM)</h3>
                <p className="text-xs text-muted-foreground mb-3">Assets (positive) vs. Liabilities (negative)</p>
                <div className="h-52">
                  <ResponsiveContainer width="100%" height="100%">
                    <BarChart layout="vertical" data={[
                      { name: 'Accounts Receivable', value: 1924857 },
                      { name: 'Inventory', value: 1568739 },
                      { name: 'Prepaid Expenses', value: 62745 },
                      { name: 'Other Current Assets', value: 27700 },
                      { name: 'Accounts Payable', value: -892847 },
                      { name: 'Accrued Expenses', value: -438293 },
                      { name: 'Deferred Revenue', value: -142000 },
                      { name: 'Sales Tax Payable', value: -68293 },
                    ]} margin={{ top: 4, right: 16, left: 4, bottom: 4 }}>
                      <CartesianGrid strokeDasharray="3 3" stroke="#f0f0f0" horizontal={false} />
                      <XAxis type="number" tickFormatter={v => `$${(v/1e6).toFixed(1)}M`} tick={{ fontSize: 10 }} />
                      <YAxis type="category" dataKey="name" tick={{ fontSize: 9 }} width={120} />
                      <Tooltip formatter={(v: number | undefined) => formatCurrency(v ?? 0)} />
                      <Bar dataKey="value" name="Amount" radius={[0,3,3,0]}>
                        {[1924857,1568739,62745,27700,-892847,-438293,-142000,-68293].map((v, i) => (
                          <Cell key={i} fill={v >= 0 ? '#2563EB' : '#EF4444'} />
                        ))}
                      </Bar>
                    </BarChart>
                  </ResponsiveContainer>
                </div>
              </div>
              <div className="rounded-lg border bg-card p-4">
                <h3 className="text-sm font-semibold mb-1">Working Capital Days</h3>
                <p className="text-xs text-muted-foreground mb-3">DSO, DIO & DPO (days outstanding)</p>
                <div className="h-52">
                  <ResponsiveContainer width="100%" height="100%">
                    <BarChart data={[
                      { metric: 'DSO', current: 10.3, target: 12.0 },
                      { metric: 'DIO', current: 19.4, target: 22.0 },
                      { metric: 'DPO', current: 11.1, target: 10.0 },
                    ]} margin={{ top: 4, right: 8, left: 8, bottom: 4 }} barGap={6}>
                      <CartesianGrid strokeDasharray="3 3" stroke="#f0f0f0" vertical={false} />
                      <XAxis dataKey="metric" tick={{ fontSize: 12 }} />
                      <YAxis tick={{ fontSize: 10 }} tickFormatter={v => `${v}d`} domain={[0, 30]} width={36} />
                      <Tooltip formatter={(v: number | undefined) => `${v ?? 0} days`} />
                      <Legend wrapperStyle={{ fontSize: 11 }} />
                      <Bar dataKey="current" name="Current" fill="#2563EB" radius={[3,3,0,0]} />
                      <Bar dataKey="target" name="Target" fill="#CBD5E1" radius={[3,3,0,0]} />
                    </BarChart>
                  </ResponsiveContainer>
                </div>
              </div>
            </div>
            <div className="rounded-md border overflow-x-auto">
              <Table>
                <TableHeader>
                  <TableRow>
                    <TableHead className="font-semibold w-[250px]">Component</TableHead>
                    <TableHead className="font-semibold text-right w-[130px]">Current</TableHead>
                    <TableHead className="font-semibold text-right w-[80px]">Days</TableHead>
                    <TableHead className="font-semibold text-right w-[80px]">Target</TableHead>
                    <TableHead className="font-semibold text-right w-[130px]">Normalized</TableHead>
                    <TableHead className="font-semibold text-right w-[130px]">Adjustment</TableHead>
                  </TableRow>
                </TableHeader>
                <TableBody>
                  {workingCapitalData.map((row, idx) => {
                    const rowKey = 'wc_' + row.item.toLowerCase().replace(/[^a-z0-9]+/g, '_')
                    const currentVal = getLiveValue('working-capital', rowKey, 'TTM') ?? row.current
                    const normalizedVal = getLiveValue('working-capital', rowKey + '_normalized', 'TTM') ?? row.normalized
                    const adjustment = normalizedVal - currentVal
                    const srcRef = getCellSourceRef('working-capital', rowKey, 'TTM')
                    const cellId = getCellId('working-capital', rowKey, 'TTM')
                    return (
                      <TableRow key={idx} data-row-key={rowKey} className={row.isBold ? 'bg-muted/50 font-semibold' : ''}>
                        <TableCell className={row.isBold ? 'font-semibold' : ''}>{row.item}</TableCell>
                        <TableCell className="text-right font-mono">
                          <AuditableCell
                            value={formatCurrency(currentVal)}
                            source={row.source}
                            sourceRef={srcRef}
                            cellId={cellId}
                            workbookId={isDemoWorkbook ? undefined : id}
                            onViewSource={handleViewSource}
                          />
                        </TableCell>
                        <TableCell className="text-right font-mono text-sm text-muted-foreground">
                          {row.days > 0 ? row.days.toFixed(1) : '-'}
                        </TableCell>
                        <TableCell className="text-right font-mono text-sm text-muted-foreground">
                          {row.targetDays > 0 ? row.targetDays.toFixed(1) : '-'}
                        </TableCell>
                        <TableCell className="text-right font-mono">
                          <AuditableCell value={formatCurrency(normalizedVal)} source={row.source} isEditable={false} />
                        </TableCell>
                        <TableCell className="text-right font-mono">
                          {adjustment !== 0 ? formatCurrency(adjustment) : '-'}
                        </TableCell>
                      </TableRow>
                    )
                  })}
                </TableBody>
              </Table>
            </div>
          </div>
        )

      case 'cogs-vendors':
        return (
          <div>
            <h2 className="text-2xl font-bold mb-6">COGS Expenditures by Vendor</h2>
            <p className="text-sm text-muted-foreground mb-4">Vendor concentration and supplier analysis</p>
            {/* Chart */}
            <div className="grid grid-cols-2 gap-4 mb-8">
              <div className="rounded-lg border bg-card p-4">
                <h3 className="text-sm font-semibold mb-1">Vendor Spend — TTM</h3>
                <p className="text-xs text-muted-foreground mb-3">Top 8 vendors by COGS spend</p>
                <div className="h-52">
                  <ResponsiveContainer width="100%" height="100%">
                    <BarChart layout="vertical" data={cogsVendorsData.filter(r => !r.isBold).map(r => ({ name: r.vendor.length > 22 ? r.vendor.slice(0,22)+'…' : r.vendor, ttm: r.ttm }))} margin={{ top: 4, right: 16, left: 4, bottom: 4 }}>
                      <CartesianGrid strokeDasharray="3 3" stroke="#f0f0f0" horizontal={false} />
                      <XAxis type="number" tickFormatter={v => `$${(v/1e6).toFixed(1)}M`} tick={{ fontSize: 10 }} />
                      <YAxis type="category" dataKey="name" tick={{ fontSize: 9 }} width={120} />
                      <Tooltip formatter={(v: number | undefined) => formatCurrency(v ?? 0)} />
                      <Bar dataKey="ttm" name="TTM Spend" fill="#2563EB" radius={[0,3,3,0]} />
                    </BarChart>
                  </ResponsiveContainer>
                </div>
              </div>
              <div className="rounded-lg border bg-card p-4">
                <h3 className="text-sm font-semibold mb-1">Vendor Spend Growth</h3>
                <p className="text-xs text-muted-foreground mb-3">Top 4 vendors FY20–TTM</p>
                <div className="h-52">
                  <ResponsiveContainer width="100%" height="100%">
                    <LineChart data={[
                      { period: 'FY20', v1: cogsVendorsData[0]?.fy20, v2: cogsVendorsData[1]?.fy20, v3: cogsVendorsData[2]?.fy20, v4: cogsVendorsData[3]?.fy20 },
                      { period: 'FY21', v1: cogsVendorsData[0]?.fy21, v2: cogsVendorsData[1]?.fy21, v3: cogsVendorsData[2]?.fy21, v4: cogsVendorsData[3]?.fy21 },
                      { period: 'FY22', v1: cogsVendorsData[0]?.fy22, v2: cogsVendorsData[1]?.fy22, v3: cogsVendorsData[2]?.fy22, v4: cogsVendorsData[3]?.fy22 },
                      { period: 'TTM',  v1: cogsVendorsData[0]?.ttm,  v2: cogsVendorsData[1]?.ttm,  v3: cogsVendorsData[2]?.ttm,  v4: cogsVendorsData[3]?.ttm  },
                    ]} margin={{ top: 4, right: 16, left: 30, bottom: 4 }}>
                      <CartesianGrid strokeDasharray="3 3" stroke="#f0f0f0" vertical={false} />
                      <XAxis dataKey="period" tick={{ fontSize: 11 }} />
                      <YAxis tickFormatter={v => `$${(v/1e6).toFixed(1)}M`} tick={{ fontSize: 10 }} width={48} />
                      <Tooltip formatter={(v: number | undefined) => formatCurrency(v ?? 0)} />
                      <Legend wrapperStyle={{ fontSize: 10 }} formatter={(val, entry) => cogsVendorsData[parseInt((entry as {dataKey: string}).dataKey.replace('v',''))-1]?.vendor?.split(' ')[0] ?? val} />
                      <Line type="monotone" dataKey="v1" stroke="#2563EB" strokeWidth={2} dot={{ r: 2 }} />
                      <Line type="monotone" dataKey="v2" stroke="#7C3AED" strokeWidth={2} dot={{ r: 2 }} />
                      <Line type="monotone" dataKey="v3" stroke="#10B981" strokeWidth={2} dot={{ r: 2 }} />
                      <Line type="monotone" dataKey="v4" stroke="#F59E0B" strokeWidth={2} dot={{ r: 2 }} />
                    </LineChart>
                  </ResponsiveContainer>
                </div>
              </div>
            </div>
            <div className="rounded-md border overflow-x-auto">
              <Table>
                <TableHeader>
                  <TableRow>
                    <TableHead className="font-semibold w-[200px]">Vendor</TableHead>
                    <TableHead className="font-semibold text-right w-[130px]">FY20</TableHead>
                    <TableHead className="font-semibold text-right w-[130px]">FY21</TableHead>
                    <TableHead className="font-semibold text-right w-[130px]">FY22</TableHead>
                    <TableHead className="font-semibold text-right w-[130px]">TTM Jan-23</TableHead>
                    <TableHead className="font-semibold text-right w-[100px]">% of COGS</TableHead>
                  </TableRow>
                </TableHeader>
                <TableBody>
                  {cogsVendorsData.map((row, idx) => {
                    const rowKey = 'vendor_' + row.vendor.toLowerCase().replace(/[^a-z0-9]+/g, '_').slice(0, 30)
                    return (
                      <TableRow key={idx} data-row-key={rowKey} className={row.isBold ? 'bg-muted/50 font-semibold' : ''}>
                        <TableCell className={row.isBold ? 'font-semibold' : ''}>{row.vendor}</TableCell>
                        <TableCell className="text-right font-mono text-sm">
                          {formatCurrency(getLiveValue('cogs-vendors', rowKey, 'FY20') ?? row.fy20)}
                        </TableCell>
                        <TableCell className="text-right font-mono text-sm">
                          {formatCurrency(getLiveValue('cogs-vendors', rowKey, 'FY21') ?? row.fy21)}
                        </TableCell>
                        <TableCell className="text-right font-mono text-sm">
                          {formatCurrency(getLiveValue('cogs-vendors', rowKey, 'FY22') ?? row.fy22)}
                        </TableCell>
                        <TableCell className="text-right font-mono text-sm">
                          <AuditableCell
                            value={formatCurrency(getLiveValue('cogs-vendors', rowKey, 'TTM') ?? row.ttm)}
                            source={row.source}
                            sourceRef={getCellSourceRef('cogs-vendors', rowKey, 'TTM')}
                            cellId={getCellId('cogs-vendors', rowKey, 'TTM')}
                            workbookId={isDemoWorkbook ? undefined : id}
                            onViewSource={handleViewSource}
                          />
                        </TableCell>
                        <TableCell className="text-right font-mono text-sm">{formatPercent(row.pct)}</TableCell>
                      </TableRow>
                    )
                  })}
                </TableBody>
              </Table>
            </div>
          </div>
        )

      case 'testing':
        return (
          <div>
            <h2 className="text-2xl font-bold mb-6">Accounts Payable / Accrued Expense Testing</h2>
            <p className="text-sm text-muted-foreground mb-4">Sample testing of vendor invoices and disbursements</p>
            {/* Chart */}
            <div className="grid grid-cols-2 gap-4 mb-8">
              <div className="rounded-lg border bg-card p-4">
                <h3 className="text-sm font-semibold mb-1">Check Amount Allocation</h3>
                <p className="text-xs text-muted-foreground mb-3">Period vs. future period vs. unaccrued per disbursement</p>
                <div className="h-52">
                  <ResponsiveContainer width="100%" height="100%">
                    <BarChart data={testingData.map(r => ({ check: r.checkNum, period: r.periodAmt, future: r.futureAmt, notAccrued: r.notAccrued }))} margin={{ top: 4, right: 8, left: 28, bottom: 4 }}>
                      <CartesianGrid strokeDasharray="3 3" stroke="#f0f0f0" vertical={false} />
                      <XAxis dataKey="check" tick={{ fontSize: 10 }} />
                      <YAxis tickFormatter={v => `$${(v/1000).toFixed(0)}K`} tick={{ fontSize: 10 }} width={44} />
                      <Tooltip formatter={(v: number | undefined) => formatCurrency(v ?? 0)} />
                      <Legend wrapperStyle={{ fontSize: 11 }} />
                      <Bar dataKey="period" name="In Period" stackId="a" fill="#2563EB" />
                      <Bar dataKey="future" name="Future Period" stackId="a" fill="#F59E0B" />
                      <Bar dataKey="notAccrued" name="Not Accrued" stackId="a" fill="#EF4444" radius={[3,3,0,0]} />
                    </BarChart>
                  </ResponsiveContainer>
                </div>
              </div>
              <div className="rounded-lg border bg-card p-4">
                <h3 className="text-sm font-semibold mb-1">Testing Summary</h3>
                <p className="text-xs text-muted-foreground mb-3">Coverage & error rate</p>
                <div className="h-52 flex flex-col justify-center gap-4 px-4">
                  {[
                    { label: 'Disbursements Tested', value: formatCurrency(testingTotals.totalTested), sub: `${testingTotals.percentTested}% of total AP` },
                    { label: 'Total Errors Found', value: testingTotals.totalError > 0 ? formatCurrency(testingTotals.totalError) : '$0', sub: testingTotals.totalError > 0 ? 'Review required' : 'No material errors', ok: testingTotals.totalError === 0 },
                    { label: 'Error Rate', value: testingTotals.totalTested > 0 ? `${((testingTotals.totalError/testingTotals.totalTested)*100).toFixed(2)}%` : '0.00%', sub: 'Below 0.5% materiality threshold', ok: true },
                  ].map(item => (
                    <div key={item.label} className="flex items-center justify-between py-2 border-b last:border-0">
                      <div>
                        <p className="text-xs text-muted-foreground">{item.label}</p>
                        <p className="text-sm font-semibold mt-0.5">{item.sub}</p>
                      </div>
                      <span className={`text-lg font-bold ${item.ok === false ? 'text-red-600' : item.ok ? 'text-green-600' : ''}`}>{item.value}</span>
                    </div>
                  ))}
                </div>
              </div>
            </div>
            <div className="rounded-md border overflow-x-auto mb-6">
              <Table>
                <TableHeader>
                  <TableRow>
                    <TableHead className="font-semibold w-[100px]">Check #</TableHead>
                    <TableHead className="font-semibold w-[100px]">Invoice #</TableHead>
                    <TableHead className="font-semibold w-[180px]">Payee</TableHead>
                    <TableHead className="font-semibold w-[110px]">Date</TableHead>
                    <TableHead className="font-semibold text-right w-[120px]">Check Amt</TableHead>
                    <TableHead className="font-semibold text-right w-[120px]">Period Amt</TableHead>
                    <TableHead className="font-semibold text-right w-[120px]">Future Amt</TableHead>
                    <TableHead className="font-semibold text-right w-[120px]">Not Accrued</TableHead>
                  </TableRow>
                </TableHeader>
                <TableBody>
                  {testingData.map((row, idx) => (
                    <TableRow key={idx}>
                      <TableCell className="font-mono text-sm">{row.checkNum}</TableCell>
                      <TableCell className="font-mono text-sm">{row.invoice}</TableCell>
                      <TableCell className="text-sm">{row.payee}</TableCell>
                      <TableCell className="text-sm">{row.checkDate}</TableCell>
                      <TableCell className="text-right font-mono text-sm">{formatCurrency(row.checkAmt)}</TableCell>
                      <TableCell className="text-right font-mono text-sm">{formatCurrency(row.periodAmt)}</TableCell>
                      <TableCell className="text-right font-mono text-sm">{formatCurrency(row.futureAmt)}</TableCell>
                      <TableCell className="text-right font-mono text-sm">
                        {row.notAccrued > 0 ? (
                          <span className="text-red-500">{formatCurrency(row.notAccrued)}</span>
                        ) : '-'}
                      </TableCell>
                    </TableRow>
                  ))}
                  <TableRow className="bg-muted/50 font-semibold">
                    <TableCell colSpan={4}>Testing Summary</TableCell>
                    <TableCell className="text-right font-mono">{formatCurrency(testingTotals.totalTested)}</TableCell>
                    <TableCell colSpan={2} className="text-right">
                      <span className="text-sm text-muted-foreground">
                        {testingTotals.percentTested}% of {formatCurrency(testingTotals.totalDisbursements)} tested
                      </span>
                    </TableCell>
                    <TableCell className="text-right font-mono">
                      {testingTotals.totalError > 0 ? (
                        <span className="text-red-500">{formatCurrency(testingTotals.totalError)}</span>
                      ) : (
                        <span className="text-green-600">$0</span>
                      )}
                    </TableCell>
                  </TableRow>
                </TableBody>
              </Table>
            </div>
            <div className="flex items-center gap-2 text-sm">
              <Badge className="bg-green-600">No Material Errors Found</Badge>
              <span className="text-muted-foreground">Sample size: {testingTotals.percentTested}% of total disbursements</span>
            </div>
          </div>
        )

      case 'net-debt':
        return (
          <div>
            <h2 className="text-2xl font-bold mb-2">Net Debt & Debt-Like Items</h2>
            <p className="text-sm text-muted-foreground mb-6">Closing balance sheet debt schedule and debt-like obligations</p>
            {/* Chart */}
            <div className="grid grid-cols-2 gap-4 mb-8">
              <div className="rounded-lg border bg-card p-4">
                <h3 className="text-sm font-semibold mb-1">Debt Components Over Time</h3>
                <p className="text-xs text-muted-foreground mb-3">Stacked debt & debt-like items ($ millions)</p>
                <div className="h-52">
                  <ResponsiveContainer width="100%" height="100%">
                    <BarChart data={workbookPeriods.map(p => ({
                      period: p,
                      reportedDebt: (isDemoWorkbook ? DEMO_NET_DEBT.reported_debt?.[p] : getLiveValue('net-debt','reported_debt',p)) ?? 0,
                      leases: (isDemoWorkbook ? DEMO_NET_DEBT.capital_leases?.[p] : getLiveValue('net-debt','capital_leases',p)) ?? 0,
                      other: ((isDemoWorkbook ? DEMO_NET_DEBT.unpaid_accrued_vacation?.[p] : getLiveValue('net-debt','unpaid_accrued_vacation',p)) ?? 0) + ((isDemoWorkbook ? DEMO_NET_DEBT.aged_accounts_payable?.[p] : getLiveValue('net-debt','aged_accounts_payable',p)) ?? 0) + ((isDemoWorkbook ? DEMO_NET_DEBT.deferred_revenue?.[p] : getLiveValue('net-debt','deferred_revenue',p)) ?? 0),
                      cash: (isDemoWorkbook ? DEMO_NET_DEBT.cash_and_equivalents?.[p] : getLiveValue('net-debt','cash_and_equivalents',p)) ?? 0,
                    }))} margin={{ top: 4, right: 8, left: 28, bottom: 4 }}>
                      <CartesianGrid strokeDasharray="3 3" stroke="#f0f0f0" vertical={false} />
                      <XAxis dataKey="period" tick={{ fontSize: 11 }} />
                      <YAxis tickFormatter={v => `$${(v/1e6).toFixed(1)}M`} tick={{ fontSize: 10 }} width={48} />
                      <Tooltip formatter={(v: number | undefined) => formatCurrency(v ?? 0)} />
                      <Legend wrapperStyle={{ fontSize: 11 }} />
                      <Bar dataKey="reportedDebt" name="Term Debt" stackId="d" fill="#EF4444" />
                      <Bar dataKey="leases" name="Leases" stackId="d" fill="#F97316" />
                      <Bar dataKey="other" name="Debt-Like Items" stackId="d" fill="#FCA5A5" radius={[3,3,0,0]} />
                      <Bar dataKey="cash" name="Cash (Offset)" fill="#10B981" radius={[3,3,0,0]} />
                    </BarChart>
                  </ResponsiveContainer>
                </div>
              </div>
              <div className="rounded-lg border bg-card p-4">
                <h3 className="text-sm font-semibold mb-1">Net Debt Position</h3>
                <p className="text-xs text-muted-foreground mb-3">Negative = net cash; below zero is favorable</p>
                <div className="h-52">
                  <ResponsiveContainer width="100%" height="100%">
                    <BarChart data={workbookPeriods.map(p => ({ period: p, netDebt: (isDemoWorkbook ? DEMO_NET_DEBT.net_debt?.[p] : getLiveValue('net-debt','net_debt',p)) ?? 0 }))} margin={{ top: 4, right: 8, left: 28, bottom: 4 }}>
                      <CartesianGrid strokeDasharray="3 3" stroke="#f0f0f0" vertical={false} />
                      <XAxis dataKey="period" tick={{ fontSize: 11 }} />
                      <YAxis tickFormatter={v => `$${(v/1e6).toFixed(1)}M`} tick={{ fontSize: 10 }} width={48} />
                      <Tooltip formatter={(v: number | undefined) => formatCurrency(v ?? 0)} />
                      <Bar dataKey="netDebt" name="Net Debt" radius={[3,3,0,0]}>
                        {workbookPeriods.map((p, i) => {
                          const v = (isDemoWorkbook ? DEMO_NET_DEBT.net_debt?.[p] : getLiveValue('net-debt','net_debt',p)) ?? 0
                          return <Cell key={i} fill={v < 0 ? '#10B981' : '#EF4444'} />
                        })}
                      </Bar>
                    </BarChart>
                  </ResponsiveContainer>
                </div>
              </div>
            </div>
            <div className="rounded-md border overflow-x-auto">
              <Table>
                <TableHeader>
                  <TableRow>
                    <TableHead className="font-semibold w-[280px]">Item</TableHead>
                    {workbookPeriods.map(p => <TableHead key={p} className="font-semibold text-right w-[140px]">{p}</TableHead>)}
                  </TableRow>
                </TableHeader>
                <TableBody>
                  {[
                    { rowKey: 'reported_debt', label: 'Reported Debt (Term Loan + Revolver)', bold: false },
                    { rowKey: 'capital_leases', label: 'Capital Leases', bold: false },
                    { rowKey: 'unpaid_accrued_vacation', label: 'Unpaid Accrued Vacation', bold: false },
                    { rowKey: 'aged_accounts_payable', label: 'Aged Accounts Payable (>90 days)', bold: false },
                    { rowKey: 'deferred_revenue', label: 'Deferred Revenue', bold: false },
                    { rowKey: 'customer_deposits', label: 'Customer Deposits', bold: false },
                    { rowKey: 'total_debt_like_items', label: 'Total Debt & Debt-Like Items', bold: true },
                    { rowKey: 'cash_and_equivalents', label: 'Less: Cash & Cash Equivalents', bold: false },
                    { rowKey: 'net_debt', label: 'Net Debt', bold: true },
                  ].map((row) => (
                    <TableRow key={row.rowKey} data-row-key={row.rowKey} className={row.bold ? 'bg-muted/50' : ''}>
                      <TableCell className={row.bold ? 'font-semibold' : ''}>{row.label}</TableCell>
                      {workbookPeriods.map(p => {
                        const live = getLiveValue('net-debt', row.rowKey, p)
                        const val = live ?? (isDemoWorkbook ? (DEMO_NET_DEBT[row.rowKey]?.[p] ?? null) : null)
                        const srcRef = getCellSourceRef('net-debt', row.rowKey, p)
                        const cellId = getCellId('net-debt', row.rowKey, p)
                        return (
                          <TableCell key={p} className="text-right font-mono text-sm">
                            {val !== null ? (
                              <AuditableCell
                                value={formatCurrency(val)}
                                sourceRef={srcRef}
                                cellId={cellId}
                                workbookId={isDemoWorkbook ? undefined : id}
                                onViewSource={handleViewSource}
                              />
                            ) : <span className="text-muted-foreground">—</span>}
                          </TableCell>
                        )
                      })}
                    </TableRow>
                  ))}
                </TableBody>
              </Table>
            </div>
            <div className="mt-4 grid grid-cols-3 gap-4">
              <div className="rounded-lg border bg-green-50 border-green-200 p-4">
                <p className="text-xs font-semibold text-green-800 mb-1">TTM Net Position</p>
                <p className="text-xl font-bold text-green-700">Net Cash $1.35M</p>
                <p className="text-xs text-green-600 mt-1">Cash exceeds all debt obligations</p>
              </div>
              <div className="rounded-lg border bg-slate-50 p-4">
                <p className="text-xs font-semibold text-muted-foreground mb-1">Total Debt-Like Items (TTM)</p>
                <p className="text-xl font-bold">$1.78M</p>
                <p className="text-xs text-muted-foreground mt-1">Term loan + leases + contingencies</p>
              </div>
              <div className="rounded-lg border bg-slate-50 p-4">
                <p className="text-xs font-semibold text-muted-foreground mb-1">Net Debt / Adj. EBITDA</p>
                <p className="text-xl font-bold">-0.16×</p>
                <p className="text-xs text-muted-foreground mt-1">Negative = net cash; strong coverage</p>
              </div>
            </div>
          </div>
        )

      case 'customer-concentration':
        return (
          <div>
            <h2 className="text-2xl font-bold mb-2">Revenue & Customer Concentration</h2>
            <p className="text-sm text-muted-foreground mb-6">Revenue breakdown by customer — concentration risk assessment</p>
            {/* Charts */}
            <div className="grid grid-cols-2 gap-4 mb-8">
              <div className="rounded-lg border bg-card p-4">
                <h3 className="text-sm font-semibold mb-1">Customer Revenue Mix — TTM</h3>
                <p className="text-xs text-muted-foreground mb-3">Top 10 customers + all other</p>
                <div className="h-52">
                  <ResponsiveContainer width="100%" height="100%">
                    <RechartsPie data={[
                      { name: 'Acme Mfg', value: (isDemoWorkbook ? DEMO_CUST_CONC.customer_1?.TTM : getLiveValue('customer-concentration','customer_1','TTM')) ?? 0 },
                      { name: 'Nexgen', value: (isDemoWorkbook ? DEMO_CUST_CONC.customer_2?.TTM : getLiveValue('customer-concentration','customer_2','TTM')) ?? 0 },
                      { name: 'Bridgewater', value: (isDemoWorkbook ? DEMO_CUST_CONC.customer_3?.TTM : getLiveValue('customer-concentration','customer_3','TTM')) ?? 0 },
                      { name: 'Summit', value: (isDemoWorkbook ? DEMO_CUST_CONC.customer_4?.TTM : getLiveValue('customer-concentration','customer_4','TTM')) ?? 0 },
                      { name: 'Crestwood', value: (isDemoWorkbook ? DEMO_CUST_CONC.customer_5?.TTM : getLiveValue('customer-concentration','customer_5','TTM')) ?? 0 },
                      { name: 'All Other', value: (isDemoWorkbook ? DEMO_CUST_CONC.all_other_customers?.TTM : getLiveValue('customer-concentration','all_other_customers','TTM')) ?? 0 },
                    ].filter(d => d.value > 0)} cx="50%" cy="50%" innerRadius={50} outerRadius={84}>
                      {['#2563EB','#7C3AED','#0EA5E9','#10B981','#F59E0B','#94A3B8'].map((c, i) => <Cell key={i} fill={c} />)}
                      <Tooltip formatter={(v: number | undefined) => formatCurrency(v ?? 0)} />
                      <Legend wrapperStyle={{ fontSize: 10 }} />
                    </RechartsPie>
                  </ResponsiveContainer>
                </div>
              </div>
              <div className="rounded-lg border bg-card p-4">
                <h3 className="text-sm font-semibold mb-1">Top Customer Revenue Growth</h3>
                <p className="text-xs text-muted-foreground mb-3">Top 3 customers FY20–TTM</p>
                <div className="h-52">
                  <ResponsiveContainer width="100%" height="100%">
                    <LineChart data={workbookPeriods.map(p => ({
                      period: p,
                      acme: (isDemoWorkbook ? DEMO_CUST_CONC.customer_1?.[p] : getLiveValue('customer-concentration','customer_1',p)) ?? 0,
                      nexgen: (isDemoWorkbook ? DEMO_CUST_CONC.customer_2?.[p] : getLiveValue('customer-concentration','customer_2',p)) ?? 0,
                      bridge: (isDemoWorkbook ? DEMO_CUST_CONC.customer_3?.[p] : getLiveValue('customer-concentration','customer_3',p)) ?? 0,
                    }))} margin={{ top: 4, right: 16, left: 30, bottom: 4 }}>
                      <CartesianGrid strokeDasharray="3 3" stroke="#f0f0f0" vertical={false} />
                      <XAxis dataKey="period" tick={{ fontSize: 11 }} />
                      <YAxis tickFormatter={v => `$${(v/1e6).toFixed(0)}M`} tick={{ fontSize: 10 }} width={48} />
                      <Tooltip formatter={(v: number | undefined) => formatCurrency(v ?? 0)} />
                      <Legend wrapperStyle={{ fontSize: 10 }} />
                      <Line type="monotone" dataKey="acme" name="Acme Mfg" stroke="#2563EB" strokeWidth={2} dot={{ r: 3 }} />
                      <Line type="monotone" dataKey="nexgen" name="Nexgen" stroke="#7C3AED" strokeWidth={2} dot={{ r: 3 }} />
                      <Line type="monotone" dataKey="bridge" name="Bridgewater" stroke="#10B981" strokeWidth={2} dot={{ r: 3 }} />
                    </LineChart>
                  </ResponsiveContainer>
                </div>
              </div>
            </div>
            <div className="rounded-md border overflow-x-auto">
              <Table>
                <TableHeader>
                  <TableRow>
                    <TableHead className="font-semibold w-[220px]">Customer</TableHead>
                    {workbookPeriods.map(p => <TableHead key={p} className="font-semibold text-right w-[140px]">{p}</TableHead>)}
                  </TableRow>
                </TableHeader>
                <TableBody>
                  {[
                    { rowKey: 'customer_1',          label: 'Acme Manufacturing'      },
                    { rowKey: 'customer_2',          label: 'Nexgen Solutions'         },
                    { rowKey: 'customer_3',          label: 'Bridgewater Corp'         },
                    { rowKey: 'customer_4',          label: 'Summit Enterprises'       },
                    { rowKey: 'customer_5',          label: 'Crestwood Partners'       },
                    { rowKey: 'customer_6',          label: 'Lakeside Industries'      },
                    { rowKey: 'customer_7',          label: 'Horizon Group'            },
                    { rowKey: 'customer_8',          label: 'Pacific Dynamics'         },
                    { rowKey: 'customer_9',          label: 'Allied Services'          },
                    { rowKey: 'customer_10',         label: 'Meridian LLC'             },
                    { rowKey: 'all_other_customers', label: 'All Other Customers'      },
                    { rowKey: 'total_revenue',       label: 'Total Revenue', bold: true },
                  ].map((row) => (
                    <TableRow key={row.rowKey} data-row-key={row.rowKey} className={(row as {bold?: boolean}).bold ? 'bg-muted/50' : ''}>
                      <TableCell className={(row as {bold?: boolean}).bold ? 'font-semibold' : ''}>{row.label}</TableCell>
                      {workbookPeriods.map(p => {
                        const live = getLiveValue('customer-concentration', row.rowKey, p)
                        const val = live ?? (isDemoWorkbook ? (DEMO_CUST_CONC[row.rowKey]?.[p] ?? null) : null)
                        const srcRef = getCellSourceRef('customer-concentration', row.rowKey, p)
                        const cellId = getCellId('customer-concentration', row.rowKey, p)
                        const pct = (val !== null && DEMO_CUST_CONC.total_revenue?.[p])
                          ? (val / DEMO_CUST_CONC.total_revenue[p]! * 100).toFixed(1)
                          : null
                        return (
                          <TableCell key={p} className="text-right font-mono text-sm">
                            {val !== null ? (
                              <span className="flex items-center justify-end gap-2">
                                <AuditableCell
                                  value={formatCurrency(val)}
                                  sourceRef={srcRef}
                                  cellId={cellId}
                                  workbookId={isDemoWorkbook ? undefined : id}
                                  onViewSource={handleViewSource}
                                />
                                {pct && <span className="text-muted-foreground text-xs">{pct}%</span>}
                              </span>
                            ) : <span className="text-muted-foreground">—</span>}
                          </TableCell>
                        )
                      })}
                    </TableRow>
                  ))}
                </TableBody>
              </Table>
            </div>
            <div className="mt-4 grid grid-cols-3 gap-4">
              <div className="rounded-lg border bg-amber-50 border-amber-200 p-4">
                <p className="text-xs font-semibold text-amber-800 mb-1">Top Customer Concentration</p>
                <p className="text-xl font-bold text-amber-700">20.8%</p>
                <p className="text-xs text-amber-600 mt-1">Acme Manufacturing — monitor risk</p>
              </div>
              <div className="rounded-lg border bg-amber-50 border-amber-200 p-4">
                <p className="text-xs font-semibold text-amber-800 mb-1">Top 3 Customer Concentration</p>
                <p className="text-xl font-bold text-amber-700">45.8%</p>
                <p className="text-xs text-amber-600 mt-1">Review customer agreements & retention</p>
              </div>
              <div className="rounded-lg border bg-green-50 border-green-200 p-4">
                <p className="text-xs font-semibold text-green-800 mb-1">Customer Base</p>
                <p className="text-xl font-bold text-green-700">Diversifying</p>
                <p className="text-xs text-green-600 mt-1">Top 10 share declining YoY; &ldquo;All Other&rdquo; growing</p>
              </div>
            </div>
          </div>
        )

      case 'proof-revenue':
        return (
          <div>
            <h2 className="text-2xl font-bold mb-2">Proof of Revenue — Three-Way Match</h2>
            <p className="text-sm text-muted-foreground mb-4">GL revenue vs. tax return vs. bank deposits reconciliation</p>
            {/* Reconciliation status banner */}
            {!isDemoWorkbook && reconciliation && reconciliation.status !== 'UNAVAILABLE' && (
              <div className={`rounded-lg border p-4 mb-6 ${
                reconciliation.status === 'PASS' ? 'bg-blue-50 border-blue-200' :
                reconciliation.status === 'WARN' ? 'bg-amber-50 border-amber-200' :
                'bg-red-50 border-red-200'
              }`}>
                <div className="flex items-center justify-between mb-2">
                  <div className="flex items-center gap-2">
                    <span className={`text-sm font-semibold ${
                      reconciliation.status === 'PASS' ? 'text-blue-700' :
                      reconciliation.status === 'WARN' ? 'text-amber-700' :
                      'text-red-700'
                    }`}>
                      Three-Way Match: {reconciliation.status === 'PASS' ? 'Clean' : reconciliation.status === 'WARN' ? 'Minor Variances' : 'Significant Variances'}
                    </span>
                    <Badge className={
                      reconciliation.status === 'PASS' ? 'bg-blue-600 hover:bg-blue-700 text-xs' :
                      reconciliation.status === 'WARN' ? 'bg-amber-500 hover:bg-amber-600 text-xs' :
                      'bg-red-600 hover:bg-red-700 text-xs'
                    }>{reconciliation.status}</Badge>
                  </div>
                  {reconciliation.critical_count > 0 && (
                    <span className="text-xs text-red-600 font-medium">{reconciliation.critical_count} critical variance{reconciliation.critical_count !== 1 ? 's' : ''}</span>
                  )}
                </div>
                <div className="flex gap-4">
                  {reconciliation.legs.map(leg => (
                    <div key={leg.type} className="flex items-center gap-1.5 text-xs text-muted-foreground">
                      <span className={`w-2 h-2 rounded-full flex-shrink-0 ${leg.status === 'PASS' ? 'bg-blue-500' : leg.status === 'WARN' ? 'bg-amber-500' : leg.status === 'FAIL' ? 'bg-red-500' : 'bg-gray-300'}`} />
                      <span>{leg.labelA} vs {leg.labelB}</span>
                      <span className="font-medium">{leg.status === 'UNAVAILABLE' ? 'N/A' : leg.overallVariancePct !== null ? `${(leg.overallVariancePct * 100).toFixed(2)}%` : leg.status}</span>
                    </div>
                  ))}
                </div>
                {reconciliation.total_discrepancies > 0 && (
                  <div className="mt-3 space-y-1">
                    {reconciliation.legs.flatMap(l => l.discrepancies).slice(0, 3).map((d, i) => (
                      <p key={i} className="text-xs text-muted-foreground">{d.description}</p>
                    ))}
                  </div>
                )}
              </div>
            )}
            {/* Chart */}
            <div className="rounded-lg border bg-card p-4 mb-8">
              <h3 className="text-sm font-semibold mb-1">Three-Way Revenue Comparison</h3>
              <p className="text-xs text-muted-foreground mb-3">GL vs. Tax Return vs. Bank Deposits — small variances confirm high-quality revenue recognition</p>
              <div className="h-56">
                <ResponsiveContainer width="100%" height="100%">
                  <BarChart data={workbookPeriods.filter(p => p !== 'TTM' || (isDemoWorkbook ? DEMO_PROOF_REV.tax_return_revenue?.[p] !== null : true)).map(p => ({
                    period: p,
                    gl: (isDemoWorkbook ? DEMO_PROOF_REV.gl_revenue?.[p] : getLiveValue('proof-revenue','gl_revenue',p)) ?? 0,
                    tax: (isDemoWorkbook ? DEMO_PROOF_REV.tax_return_revenue?.[p] : getLiveValue('proof-revenue','tax_return_revenue',p)) ?? undefined,
                    bank: (isDemoWorkbook ? DEMO_PROOF_REV.bank_deposit_revenue?.[p] : getLiveValue('proof-revenue','bank_deposit_revenue',p)) ?? 0,
                  }))} margin={{ top: 4, right: 8, left: 34, bottom: 4 }} barGap={4}>
                    <CartesianGrid strokeDasharray="3 3" stroke="#f0f0f0" vertical={false} />
                    <XAxis dataKey="period" tick={{ fontSize: 11 }} />
                    <YAxis tickFormatter={v => `$${(v/1e6).toFixed(0)}M`} tick={{ fontSize: 10 }} width={52} />
                    <Tooltip formatter={(v: number | undefined) => formatCurrency(v ?? 0)} />
                    <Legend wrapperStyle={{ fontSize: 11 }} />
                    <Bar dataKey="gl" name="GL Revenue" fill="#2563EB" radius={[3,3,0,0]} />
                    <Bar dataKey="tax" name="Tax Return" fill="#7C3AED" radius={[3,3,0,0]} />
                    <Bar dataKey="bank" name="Bank Deposits" fill="#10B981" radius={[3,3,0,0]} />
                  </BarChart>
                </ResponsiveContainer>
              </div>
            </div>
            <div className="rounded-md border overflow-x-auto">
              <Table>
                <TableHeader>
                  <TableRow>
                    <TableHead className="font-semibold w-[280px]">Item</TableHead>
                    {workbookPeriods.map(p => <TableHead key={p} className="font-semibold text-right w-[140px]">{p}</TableHead>)}
                  </TableRow>
                </TableHeader>
                <TableBody>
                  {[
                    { rowKey: 'gl_revenue', label: 'GL Revenue', bold: true },
                    { rowKey: 'tax_return_revenue', label: 'Tax Return Revenue', bold: false },
                    { rowKey: 'bank_deposit_revenue', label: 'Bank Deposit Revenue', bold: false },
                    { rowKey: 'variance_gl_vs_tax', label: 'Variance: GL vs. Tax Return', bold: false },
                    { rowKey: 'variance_gl_vs_bank', label: 'Variance: GL vs. Bank Deposits', bold: false },
                    { rowKey: 'pct_variance_gl_tax', label: '% Variance GL vs. Tax', bold: false, isPercent: true },
                  ].map((row) => (
                    <TableRow key={row.rowKey} data-row-key={row.rowKey} className={row.bold ? 'bg-muted/50' : ''}>
                      <TableCell className={row.bold ? 'font-semibold' : ''}>{row.label}</TableCell>
                      {workbookPeriods.map(p => {
                        const live = getLiveValue('proof-revenue', row.rowKey, p)
                        const demoVal = isDemoWorkbook ? (DEMO_PROOF_REV[row.rowKey]?.[p] ?? null) : null
                        const val = live ?? demoVal
                        const srcRef = getCellSourceRef('proof-revenue', row.rowKey, p)
                        const cellId = getCellId('proof-revenue', row.rowKey, p)
                        const displayVal = val !== null
                          ? ((row as {isPercent?: boolean}).isPercent ? formatPercent(val) : formatCurrency(val))
                          : null
                        const isVariance = row.rowKey.startsWith('variance') && val !== null && Math.abs(val) > 50000
                        return (
                          <TableCell key={p} className="text-right font-mono text-sm">
                            {displayVal ? (
                              <span className={isVariance ? 'text-red-600 font-semibold' : ''}>
                                <AuditableCell
                                  value={displayVal}
                                  sourceRef={srcRef}
                                  cellId={cellId}
                                  workbookId={isDemoWorkbook ? undefined : id}
                                  onViewSource={handleViewSource}
                                />
                              </span>
                            ) : <span className="text-muted-foreground">—</span>}
                          </TableCell>
                        )
                      })}
                    </TableRow>
                  ))}
                </TableBody>
              </Table>
            </div>
          </div>
        )

      case 'cash-conversion':
        return (
          <div>
            <h2 className="text-2xl font-bold mb-2">Cash Conversion Analysis</h2>
            <p className="text-sm text-muted-foreground mb-6">EBITDA to free cash flow bridge — CapEx, working capital, taxes, and interest</p>
            {/* Charts */}
            <div className="grid grid-cols-2 gap-4 mb-8">
              <div className="rounded-lg border bg-card p-4">
                <h3 className="text-sm font-semibold mb-1">EBITDA → FCF Trend</h3>
                <p className="text-xs text-muted-foreground mb-3">Adjusted EBITDA vs. Free Cash Flow</p>
                <div className="h-52">
                  <ResponsiveContainer width="100%" height="100%">
                    <BarChart data={workbookPeriods.map(p => ({
                      period: p,
                      ebitda: (isDemoWorkbook ? DEMO_CASH_CONV.ebitda?.[p] : getLiveValue('cash-conversion','ebitda',p)) ?? 0,
                      fcf: (isDemoWorkbook ? DEMO_CASH_CONV.free_cash_flow?.[p] : getLiveValue('cash-conversion','free_cash_flow',p)) ?? 0,
                    }))} margin={{ top: 4, right: 8, left: 28, bottom: 4 }} barGap={6}>
                      <CartesianGrid strokeDasharray="3 3" stroke="#f0f0f0" vertical={false} />
                      <XAxis dataKey="period" tick={{ fontSize: 11 }} />
                      <YAxis tickFormatter={v => `$${(v/1e6).toFixed(1)}M`} tick={{ fontSize: 10 }} width={48} />
                      <Tooltip formatter={(v: number | undefined) => formatCurrency(v ?? 0)} />
                      <Legend wrapperStyle={{ fontSize: 11 }} />
                      <Bar dataKey="ebitda" name="Adj. EBITDA" fill="#2563EB" radius={[3,3,0,0]} />
                      <Bar dataKey="fcf" name="Free Cash Flow" fill="#10B981" radius={[3,3,0,0]} />
                    </BarChart>
                  </ResponsiveContainer>
                </div>
              </div>
              <div className="rounded-lg border bg-card p-4">
                <h3 className="text-sm font-semibold mb-1">FCF / EBITDA Conversion Rate</h3>
                <p className="text-xs text-muted-foreground mb-3">High conversion = low CapEx + efficient WC</p>
                <div className="h-52">
                  <ResponsiveContainer width="100%" height="100%">
                    <LineChart data={workbookPeriods.map(p => ({
                      period: p,
                      rate: (isDemoWorkbook ? DEMO_CASH_CONV.fcf_to_ebitda_ratio?.[p] : getLiveValue('cash-conversion','fcf_to_ebitda_ratio',p)) ?? 0,
                    }))} margin={{ top: 4, right: 16, left: 16, bottom: 4 }}>
                      <CartesianGrid strokeDasharray="3 3" stroke="#f0f0f0" vertical={false} />
                      <XAxis dataKey="period" tick={{ fontSize: 11 }} />
                      <YAxis tickFormatter={v => `${v}%`} tick={{ fontSize: 10 }} domain={[80, 95]} width={40} />
                      <Tooltip formatter={(v: number | undefined) => `${(v ?? 0).toFixed(1)}%`} />
                      <Line type="monotone" dataKey="rate" name="FCF / EBITDA %" stroke="#10B981" strokeWidth={2.5} dot={{ r: 4, fill: '#10B981' }} />
                    </LineChart>
                  </ResponsiveContainer>
                </div>
              </div>
            </div>
            <div className="rounded-md border overflow-x-auto">
              <Table>
                <TableHeader>
                  <TableRow>
                    <TableHead className="font-semibold w-[280px]">Item</TableHead>
                    {workbookPeriods.map(p => <TableHead key={p} className="font-semibold text-right w-[140px]">{p}</TableHead>)}
                  </TableRow>
                </TableHeader>
                <TableBody>
                  {[
                    { rowKey: 'ebitda', label: 'Adjusted EBITDA', bold: true },
                    { rowKey: 'capex', label: 'Less: Capital Expenditures', bold: false },
                    { rowKey: 'change_in_working_capital', label: 'Less: Change in Working Capital', bold: false },
                    { rowKey: 'taxes_paid_cash', label: 'Less: Taxes Paid (Cash)', bold: false },
                    { rowKey: 'interest_paid_cash', label: 'Less: Interest Paid (Cash)', bold: false },
                    { rowKey: 'free_cash_flow', label: 'Free Cash Flow', bold: true },
                    { rowKey: 'fcf_to_ebitda_ratio', label: 'FCF / EBITDA Conversion', bold: false, isPercent: true },
                  ].map((row) => (
                    <TableRow key={row.rowKey} data-row-key={row.rowKey} className={row.bold ? 'bg-muted/50' : ''}>
                      <TableCell className={row.bold ? 'font-semibold' : ''}>{row.label}</TableCell>
                      {workbookPeriods.map(p => {
                        const live = getLiveValue('cash-conversion', row.rowKey, p)
                        const demoVal = isDemoWorkbook ? (DEMO_CASH_CONV[row.rowKey]?.[p] ?? null) : null
                        const val = live ?? demoVal
                        const srcRef = getCellSourceRef('cash-conversion', row.rowKey, p)
                        const cellId = getCellId('cash-conversion', row.rowKey, p)
                        const displayVal = val !== null
                          ? ((row as {isPercent?: boolean}).isPercent ? formatPercent(val) : formatCurrency(val))
                          : null
                        return (
                          <TableCell key={p} className="text-right font-mono text-sm">
                            {displayVal ? (
                              <AuditableCell
                                value={displayVal}
                                sourceRef={srcRef}
                                cellId={cellId}
                                workbookId={isDemoWorkbook ? undefined : id}
                                onViewSource={handleViewSource}
                              />
                            ) : <span className="text-muted-foreground">—</span>}
                          </TableCell>
                        )
                      })}
                    </TableRow>
                  ))}
                </TableBody>
              </Table>
            </div>
            <div className="mt-4 grid grid-cols-4 gap-4">
              {[
                { label: 'FCF / EBITDA (TTM)', value: '89.5%', sub: 'High-quality earnings conversion', color: 'green' },
                { label: 'TTM Free Cash Flow', value: '$7.65M', sub: 'After CapEx, WC & debt service', color: 'blue' },
                { label: 'Maintenance CapEx', value: '$512K', sub: '0.75% of TTM revenue', color: 'slate' },
                { label: 'NWC Change (TTM)', value: '-$285K', sub: 'Seasonal cash consumption', color: 'slate' },
              ].map(c => (
                <div key={c.label} className={`rounded-lg border p-4 ${c.color === 'green' ? 'bg-green-50 border-green-200' : c.color === 'blue' ? 'bg-blue-50 border-blue-200' : 'bg-slate-50'}`}>
                  <p className={`text-xs font-semibold mb-1 ${c.color === 'green' ? 'text-green-800' : c.color === 'blue' ? 'text-blue-800' : 'text-muted-foreground'}`}>{c.label}</p>
                  <p className={`text-xl font-bold ${c.color === 'green' ? 'text-green-700' : c.color === 'blue' ? 'text-blue-700' : ''}`}>{c.value}</p>
                  <p className={`text-xs mt-1 ${c.color === 'green' ? 'text-green-600' : c.color === 'blue' ? 'text-blue-600' : 'text-muted-foreground'}`}>{c.sub}</p>
                </div>
              ))}
            </div>
          </div>
        )

      case 'run-rate':
        return (
          <div>
            <h2 className="text-2xl font-bold mb-2">Run-Rate & Pro Forma Adjustments</h2>
            <p className="text-sm text-muted-foreground mb-6">Annualized revenue and EBITDA incorporating new contracts, full-year hires, and churn. FY20/FY21 not applicable — prior years shown for reference only.</p>
            {/* Chart */}
            <div className="grid grid-cols-2 gap-4 mb-8">
              <div className="rounded-lg border bg-card p-4">
                <h3 className="text-sm font-semibold mb-1">TTM vs. Pro Forma Revenue</h3>
                <p className="text-xs text-muted-foreground mb-3">FY22 & TTM — base vs. adjusted</p>
                <div className="h-52">
                  <ResponsiveContainer width="100%" height="100%">
                    <BarChart data={['FY22','TTM'].map(p => ({
                      period: p,
                      base: (isDemoWorkbook ? DEMO_RUN_RATE.ttm_revenue_base?.[p] : getLiveValue('run-rate','ttm_revenue_base',p)) ?? 0,
                      proForma: (isDemoWorkbook ? DEMO_RUN_RATE.pro_forma_revenue?.[p] : getLiveValue('run-rate','pro_forma_revenue',p)) ?? 0,
                    }))} margin={{ top: 4, right: 8, left: 34, bottom: 4 }} barGap={8}>
                      <CartesianGrid strokeDasharray="3 3" stroke="#f0f0f0" vertical={false} />
                      <XAxis dataKey="period" tick={{ fontSize: 12 }} />
                      <YAxis tickFormatter={v => `$${(v/1e6).toFixed(0)}M`} tick={{ fontSize: 10 }} width={52} />
                      <Tooltip formatter={(v: number | undefined) => formatCurrency(v ?? 0)} />
                      <Legend wrapperStyle={{ fontSize: 11 }} />
                      <Bar dataKey="base" name="TTM Base" fill="#94A3B8" radius={[3,3,0,0]} />
                      <Bar dataKey="proForma" name="Pro Forma" fill="#2563EB" radius={[3,3,0,0]} />
                    </BarChart>
                  </ResponsiveContainer>
                </div>
              </div>
              <div className="rounded-lg border bg-card p-4">
                <h3 className="text-sm font-semibold mb-1">Pro Forma Adjustment Bridge — TTM</h3>
                <p className="text-xs text-muted-foreground mb-3">Components of revenue uplift</p>
                <div className="h-52">
                  <ResponsiveContainer width="100%" height="100%">
                    <BarChart data={[
                      { name: 'TTM Base', value: (isDemoWorkbook ? DEMO_RUN_RATE.ttm_revenue_base?.TTM : getLiveValue('run-rate','ttm_revenue_base','TTM')) ?? 68293742, isBase: true },
                      { name: 'New Contracts', value: (isDemoWorkbook ? DEMO_RUN_RATE.new_contract_value?.TTM : getLiveValue('run-rate','new_contract_value','TTM')) ?? 3420000 },
                      { name: 'Price Increases', value: (isDemoWorkbook ? DEMO_RUN_RATE.price_increase_effect?.TTM : getLiveValue('run-rate','price_increase_effect','TTM')) ?? 1365875 },
                      { name: 'Hire Effect', value: (isDemoWorkbook ? DEMO_RUN_RATE.full_year_hire_effect?.TTM : getLiveValue('run-rate','full_year_hire_effect','TTM')) ?? -285000 },
                      { name: 'Churn', value: (isDemoWorkbook ? DEMO_RUN_RATE.lost_revenue_adjustment?.TTM : getLiveValue('run-rate','lost_revenue_adjustment','TTM')) ?? -1152487 },
                    ]} margin={{ top: 4, right: 8, left: 30, bottom: 28 }}>
                      <CartesianGrid strokeDasharray="3 3" stroke="#f0f0f0" vertical={false} />
                      <XAxis dataKey="name" tick={{ fontSize: 9 }} angle={-30} textAnchor="end" interval={0} />
                      <YAxis tickFormatter={v => `$${(v/1e6).toFixed(0)}M`} tick={{ fontSize: 10 }} width={48} />
                      <Tooltip formatter={(v: number | undefined) => formatCurrency(v ?? 0)} />
                      <Bar dataKey="value" radius={[3,3,0,0]}>
                        {[true, false, false, false, false].map((isBase, i) => {
                          const vals = [
                            (isDemoWorkbook ? DEMO_RUN_RATE.ttm_revenue_base?.TTM : getLiveValue('run-rate','ttm_revenue_base','TTM')) ?? 68293742,
                            (isDemoWorkbook ? DEMO_RUN_RATE.new_contract_value?.TTM : getLiveValue('run-rate','new_contract_value','TTM')) ?? 3420000,
                            (isDemoWorkbook ? DEMO_RUN_RATE.price_increase_effect?.TTM : getLiveValue('run-rate','price_increase_effect','TTM')) ?? 1365875,
                            (isDemoWorkbook ? DEMO_RUN_RATE.full_year_hire_effect?.TTM : getLiveValue('run-rate','full_year_hire_effect','TTM')) ?? -285000,
                            (isDemoWorkbook ? DEMO_RUN_RATE.lost_revenue_adjustment?.TTM : getLiveValue('run-rate','lost_revenue_adjustment','TTM')) ?? -1152487,
                          ]
                          return <Cell key={i} fill={isBase ? '#94A3B8' : vals[i] >= 0 ? '#10B981' : '#EF4444'} />
                        })}
                      </Bar>
                    </BarChart>
                  </ResponsiveContainer>
                </div>
              </div>
            </div>
            <div className="rounded-md border overflow-x-auto">
              <Table>
                <TableHeader>
                  <TableRow>
                    <TableHead className="font-semibold w-[280px]">Item</TableHead>
                    {workbookPeriods.map(p => <TableHead key={p} className="font-semibold text-right w-[140px]">{p}</TableHead>)}
                  </TableRow>
                </TableHeader>
                <TableBody>
                  {[
                    { rowKey: 'ttm_revenue_base', label: 'TTM Revenue (Base)', bold: true },
                    { rowKey: 'new_contract_value', label: 'Add: New Contract Value (Annualized)', bold: false },
                    { rowKey: 'full_year_hire_effect', label: 'Add: Full-Year Hire Effect', bold: false },
                    { rowKey: 'price_increase_effect', label: 'Add: Price Increase Effect', bold: false },
                    { rowKey: 'lost_revenue_adjustment', label: 'Less: Lost / Churned Revenue', bold: false },
                    { rowKey: 'pro_forma_revenue', label: 'Pro Forma Revenue', bold: true },
                    { rowKey: 'pro_forma_ebitda', label: 'Pro Forma EBITDA', bold: true },
                  ].map((row) => (
                    <TableRow key={row.rowKey} data-row-key={row.rowKey} className={row.bold ? 'bg-muted/50' : ''}>
                      <TableCell className={row.bold ? 'font-semibold' : ''}>{row.label}</TableCell>
                      {workbookPeriods.map(p => {
                        const live = getLiveValue('run-rate', row.rowKey, p)
                        const demoVal = isDemoWorkbook ? (DEMO_RUN_RATE[row.rowKey]?.[p] ?? null) : null
                        const val = live ?? demoVal
                        const srcRef = getCellSourceRef('run-rate', row.rowKey, p)
                        const cellId = getCellId('run-rate', row.rowKey, p)
                        return (
                          <TableCell key={p} className="text-right font-mono text-sm">
                            {val !== null ? (
                              <AuditableCell
                                value={formatCurrency(val)}
                                sourceRef={srcRef}
                                cellId={cellId}
                                workbookId={isDemoWorkbook ? undefined : id}
                                onViewSource={handleViewSource}
                              />
                            ) : <span className="text-muted-foreground">—</span>}
                          </TableCell>
                        )
                      })}
                    </TableRow>
                  ))}
                </TableBody>
              </Table>
            </div>
            <div className="mt-4 grid grid-cols-3 gap-4">
              <div className="rounded-lg border bg-blue-50 border-blue-200 p-4">
                <p className="text-xs font-semibold text-blue-800 mb-1">Pro Forma Revenue (TTM)</p>
                <p className="text-xl font-bold text-blue-700">$71.6M</p>
                <p className="text-xs text-blue-600 mt-1">+$3.4M net contracts; +$1.4M pricing; -$1.2M churn</p>
              </div>
              <div className="rounded-lg border bg-blue-50 border-blue-200 p-4">
                <p className="text-xs font-semibold text-blue-800 mb-1">Pro Forma EBITDA (TTM)</p>
                <p className="text-xl font-bold text-blue-700">$9.2M</p>
                <p className="text-xs text-blue-600 mt-1">12.9% pro forma margin vs. 12.5% TTM</p>
              </div>
              <div className="rounded-lg border bg-slate-50 p-4">
                <p className="text-xs font-semibold text-muted-foreground mb-1">Note</p>
                <p className="text-sm text-muted-foreground mt-1">Pro forma adjustments based on signed contracts and management representations. Buyer should independently verify new contract terms and churn assumptions.</p>
              </div>
            </div>
          </div>
        )

      case 'risk-diligence':
        return (
          <RiskDiligenceSection
            liveCells={liveCells}
            workbookId={id}
            isDemoWorkbook={isDemoWorkbook}
            periods={workbookPeriods}
            onViewSource={handleViewSource}
          />
        )

      default:
        return (
          <div className="text-center py-12">
            <h2 className="text-2xl font-bold mb-4">{sections.find(s => s.id === activeSection)?.name}</h2>
            <p className="text-muted-foreground mb-6">This section is under development</p>
            <Badge variant="outline">Coming Soon</Badge>
          </div>
        )
    }
  }

  return (
    <div className="flex h-full" suppressHydrationWarning>
      {/* Workbook Sidebar */}
      <div className="w-64 border-r bg-background flex flex-col overflow-y-auto">
        <div className="p-4 border-b">
          <h2 className="font-semibold text-lg">{workbookName}</h2>
          <p className="text-xs text-muted-foreground mt-1">Workbook Sections</p>
        </div>

        <div className="flex-1 p-3" suppressHydrationWarning>
          {sections.map((section) => {
            const Icon = section.icon
            return (
              <div
                key={section.id}
                suppressHydrationWarning
                onClick={() => setActiveSection(section.id)}
                className={`flex items-center px-3 py-2 rounded-md cursor-pointer hover:bg-accent mb-1 ${
                  activeSection === section.id ? 'bg-accent' : ''
                }`}
              >
                <Icon className="mr-2 h-4 w-4" suppressHydrationWarning />
                <span className="text-sm">{section.name}</span>
              </div>
            )
          })}
        </div>
      </div>

      {/* Main Content */}
      <div ref={mainScrollRef} className="flex-1 overflow-y-auto" style={{ contentVisibility: 'auto' }}>
        <div className="p-8">
          <div className="flex items-center justify-between mb-8">
            <div className="flex items-center gap-3">
              <h1 className="text-xl font-semibold">{workbookName}</h1>
              {workbookStatus && !isDemoWorkbook && (
                <Badge
                  className={
                    workbookStatus === 'ready'
                      ? 'bg-blue-600 hover:bg-blue-700'
                      : workbookStatus === 'needs_input' || workbookStatus === 'needs_more_data'
                      ? 'bg-amber-500 hover:bg-amber-600'
                      : workbookStatus === 'analyzing'
                      ? 'bg-gray-400 hover:bg-gray-500'
                      : workbookStatus === 'error'
                      ? 'bg-gray-500 hover:bg-gray-600'
                      : 'bg-gray-500 hover:bg-gray-600'
                  }
                >
                  {workbookStatus === 'ready'
                    ? 'Ready'
                    : workbookStatus === 'needs_input'
                    ? 'Needs Input'
                    : workbookStatus === 'needs_more_data'
                    ? 'Incomplete Data'
                    : workbookStatus === 'analyzing'
                    ? 'Analyzing...'
                    : workbookStatus === 'error'
                    ? 'Error'
                    : workbookStatus}
                </Badge>
              )}
              {isDemoWorkbook && (
                <Badge variant="outline">{id === 'sandbox' ? 'Sandbox' : 'Demo'}</Badge>
              )}
              {!isDemoWorkbook && completeness !== null && (
                <div className="flex items-center gap-1.5 ml-1">
                  <div className="w-16 h-1.5 rounded-full bg-muted overflow-hidden">
                    <div
                      className="h-full rounded-full bg-blue-500 transition-all duration-500"
                      style={{ width: `${completeness}%` }}
                    />
                  </div>
                  <span className="text-xs text-muted-foreground tabular-nums">{completeness}%</span>
                </div>
              )}
            </div>
            <div className="flex flex-col items-end gap-1">
              <div className="flex items-center gap-2">
                <Button
                  variant="outline"
                  size="icon"
                  onClick={() => chatOpen ? closeChat() : openChat()}
                  title="Workbook AI"
                >
                  <Sparkles className="h-4 w-4" />
                </Button>
                {/* Flags dropdown */}
                <div className="relative">
                  <Button
                    variant="outline"
                    size="icon"
                    className="relative"
                    title="Flags"
                    onClick={() => setFlagsOpen(o => !o)}
                  >
                    <Flag className="h-4 w-4" />
                    {openFlagCount > 0 && (
                      <span className="absolute -top-1 -right-1 bg-red-500 text-white text-[9px] font-semibold rounded-full h-4 w-4 flex items-center justify-center leading-none">
                        {openFlagCount}
                      </span>
                    )}
                  </Button>
                  {flagsOpen && (
                    <>
                      <div className="fixed inset-0 z-40" onClick={() => setFlagsOpen(false)} />
                      <div className="absolute right-0 top-full mt-1.5 z-50 w-72 rounded-lg border bg-popover shadow-lg overflow-hidden animate-in fade-in-0 zoom-in-95 duration-150">
                        <div className="px-3 py-2 border-b flex items-center justify-between">
                          <span className="text-xs font-semibold text-muted-foreground uppercase tracking-wide">Open Flags</span>
                          <span className="text-xs text-muted-foreground">{openFlagCount} open</span>
                        </div>
                        <div className="max-h-72 overflow-y-auto">
                          {flags.filter(f => !f.resolved_at).length === 0 ? (
                            <div className="px-3 py-6 text-center text-xs text-muted-foreground">No open flags</div>
                          ) : (
                            flags.filter(f => !f.resolved_at).map(flag => (
                              <button
                                key={flag.id}
                                onClick={() => navigateToFlag(flag)}
                                className="w-full text-left px-3 py-2.5 hover:bg-accent transition-colors flex items-start gap-2.5 border-b border-border/50 last:border-0"
                              >
                                <Flag className="h-3 w-3 mt-0.5 flex-shrink-0 text-red-500" />
                                <div className="min-w-0">
                                  <div className="text-xs font-medium leading-tight truncate">{flag.title}</div>
                                  <div className="text-[10px] text-muted-foreground mt-0.5">
                                    {SECTION_DISPLAY[flag.section] ?? flag.section}
                                    {flag.period && ` · ${flag.period}`}
                                  </div>
                                </div>
                              </button>
                            ))
                          )}
                        </div>
                      </div>
                    </>
                  )}
                </div>
                <Button variant="outline" size="icon" onClick={() => setSettingsOpen(true)}>
                  <Settings className="h-4 w-4" />
                </Button>
                <DropdownMenu>
                  <DropdownMenuTrigger asChild>
                    <Button disabled={exporting || isDemoWorkbook}>
                      {exporting ? (
                        <Loader2 className="mr-2 h-4 w-4 animate-spin" />
                      ) : (
                        <Download className="mr-2 h-4 w-4" />
                      )}
                      Export
                      <ChevronDown className="ml-2 h-4 w-4" />
                    </Button>
                  </DropdownMenuTrigger>
                  <DropdownMenuContent align="end">
                    <DropdownMenuItem onClick={() => handleExport('excel')}>
                      <Image src="/integrations/excel.png" alt="Excel" width={16} height={16} className="mr-2 object-contain" loading="lazy" />
                      Excel (.xlsx)
                    </DropdownMenuItem>
                    <DropdownMenuItem onClick={() => handleExport('sheets')}>
                      <Image src="/integrations/google-sheets.png" alt="Google Sheets" width={16} height={16} className="mr-2 object-contain" loading="lazy" />
                      Google Sheets
                    </DropdownMenuItem>
                    <DropdownMenuItem disabled className="opacity-50 cursor-not-allowed">
                      <Image src="/integrations/airtable.png" alt="Airtable" width={16} height={16} className="mr-2 object-contain" loading="lazy" />
                      Airtable <span className="ml-auto text-[10px] text-muted-foreground">Soon</span>
                    </DropdownMenuItem>
                    <DropdownMenuItem onClick={() => handleExport('pdf')}>
                      <FileText className="mr-2 h-4 w-4 text-gray-500" />
                      PDF
                    </DropdownMenuItem>
                  </DropdownMenuContent>
                </DropdownMenu>
              </div>
              {exportError && (
                <p className="text-xs text-gray-500 max-w-xs text-right leading-tight">
                  {exportError}
                </p>
              )}
            </div>
          </div>

          {cellsError && !isDemoWorkbook && (
            <div className="mb-6 px-4 py-3 rounded-md bg-amber-50 border border-amber-200 text-sm text-amber-800 flex items-center justify-between">
              <span>⚠️ {cellsError}</span>
              <button className="underline ml-4 font-medium" onClick={fetchCells}>Retry</button>
            </div>
          )}
          {!isDemoWorkbook && !cellsLoading && !cellsError && liveCells.length === 0 && (
            <div className="mb-6 px-5 py-6 rounded-lg bg-blue-50 border border-blue-200 text-blue-900 flex flex-col sm:flex-row items-start sm:items-center gap-4">
              <div className="flex-1">
                <p className="font-semibold text-base mb-1">Upload documents to populate this workbook</p>
                <p className="text-sm text-blue-700">Upload your financial statements, general ledger, bank statements, or trial balance. The AI will extract and populate all sections automatically — usually within 5 minutes.</p>
              </div>
              <button
                onClick={() => setSettingsOpen(true)}
                className="shrink-0 inline-flex items-center gap-2 px-4 py-2 rounded-md bg-blue-600 text-white text-sm font-medium hover:bg-blue-700 transition-colors"
              >
                <Settings className="h-4 w-4" />
                Upload Documents
              </button>
            </div>
          )}
          {renderContent()}
        </div>
      </div>

      {/* Workbook Settings Dialog */}
      <WorkbookSettingsDialog
        open={settingsOpen}
        onOpenChange={setSettingsOpen}
        workbookName={workbookName}
        workbookId={isDemoWorkbook ? undefined : id}
        onAnalysisComplete={fetchCells}
        onDeleted={() => router.push('/dashboard')}
      />

      {/* Document Viewer Panel */}
      <DocumentViewerPanel
        open={viewerOpen}
        onOpenChange={setViewerOpen}
        sourceRef={viewerSource}
      />

      {/* AI Chat Panel is rendered at layout level — see app/dashboard/layout.tsx */}
    </div>
  )
}
