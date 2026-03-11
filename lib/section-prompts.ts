export interface RowDef {
  rowKey: string
  label: string
  valueType: 'currency' | 'percent' | 'count' | 'text' | 'date'
  isCalculated: boolean
  formula?: string
}

export interface SectionConfig {
  key: string
  displayName: string
  /** Primary RAG query (fallback if ragQueries not set) */
  ragQuery: string
  /**
   * Multiple targeted queries for higher recall.
   * All run in parallel; results are deduplicated by content hash.
   * Use 2–4 specific queries that together cover all rows in this section.
   */
  ragQueries?: string[]
  /**
   * Override workbook-level periods for this section.
   * Used by margins-by-month which needs month names, not fiscal years.
   */
  overridePeriods?: string[]
  requiredRows: RowDef[]
}

export const SECTION_CONFIGS: SectionConfig[] = [
  // ──────────────────────────────────────────────────────────────────
  // OVERVIEW
  // ──────────────────────────────────────────────────────────────────
  {
    key: 'overview',
    displayName: 'Overview',
    ragQuery:
      'total revenue net revenue annual revenue EBITDA net income gross profit operating income financial summary income statement annual totals',
    ragQueries: [
      'total revenue net revenue annual gross revenue financial summary highlights income statement',
      'EBITDA adjusted EBITDA gross profit operating income diligence adjusted management adjustments bridge',
      'net income annual profit loss company overview performance summary',
    ],
    requiredRows: [
      { rowKey: 'total_revenue', label: 'Total Revenue', valueType: 'currency', isCalculated: false },
      { rowKey: 'gross_profit', label: 'Gross Profit', valueType: 'currency', isCalculated: true },
      { rowKey: 'ebitda_as_defined', label: 'EBITDA, as defined', valueType: 'currency', isCalculated: true },
      { rowKey: 'diligence_adjusted_ebitda', label: 'Diligence-Adjusted EBITDA', valueType: 'currency', isCalculated: true },
      { rowKey: 'adjusted_ebitda_margin', label: 'Adjusted EBITDA Margin %', valueType: 'percent', isCalculated: true },
    ],
  },

  // ──────────────────────────────────────────────────────────────────
  // QUALITY OF EARNINGS
  // ──────────────────────────────────────────────────────────────────
  {
    key: 'qoe',
    displayName: 'Quality of Earnings',
    ragQuery:
      'net income EBITDA interest expense depreciation amortization tax add-back adjustments owner compensation charitable contributions one-time expenses quality of earnings',
    ragQueries: [
      'net income interest expense depreciation amortization tax provision EBITDA reconciliation bridge table',
      'management adjustments owner compensation personal expenses charitable contributions add-backs one-time',
      'diligence adjustments professional fees executive search severance relocation above-market rent normalize owner compensation',
    ],
    requiredRows: [
      { rowKey: 'net_income', label: 'Net Income', valueType: 'currency', isCalculated: false },
      { rowKey: 'interest_expense', label: 'Interest Expense', valueType: 'currency', isCalculated: false },
      { rowKey: 'tax_provision', label: 'Tax Provision', valueType: 'currency', isCalculated: false },
      { rowKey: 'depreciation', label: 'Depreciation', valueType: 'currency', isCalculated: false },
      { rowKey: 'amortization', label: 'Amortization', valueType: 'currency', isCalculated: false },
      { rowKey: 'ebitda_as_defined', label: 'EBITDA, as defined', valueType: 'currency', isCalculated: true, formula: 'net_income+interest_expense+tax_provision+depreciation+amortization' },
      { rowKey: 'adj_charitable_contributions', label: 'a) Charitable Contributions', valueType: 'currency', isCalculated: false },
      { rowKey: 'adj_owner_tax_legal', label: 'b) Owner Tax & Legal', valueType: 'currency', isCalculated: false },
      { rowKey: 'adj_excess_owner_comp', label: 'c) Excess Owner Compensation', valueType: 'currency', isCalculated: false },
      { rowKey: 'adj_personal_expenses', label: 'd) Personal Expenses', valueType: 'currency', isCalculated: false },
      { rowKey: 'total_mgmt_adjustments', label: 'Total Management Adjustments', valueType: 'currency', isCalculated: true, formula: 'adj_charitable_contributions+adj_owner_tax_legal+adj_excess_owner_comp+adj_personal_expenses' },
      { rowKey: 'mgmt_adjusted_ebitda', label: 'Management-Adjusted EBITDA', valueType: 'currency', isCalculated: true, formula: 'ebitda_as_defined+total_mgmt_adjustments' },
      { rowKey: 'adj_professional_fees_onetime', label: 'e) Professional Fees - One-time', valueType: 'currency', isCalculated: false },
      { rowKey: 'adj_executive_search', label: 'f) Executive Search Fees', valueType: 'currency', isCalculated: false },
      { rowKey: 'adj_facility_relocation', label: 'g) Facility Relocation', valueType: 'currency', isCalculated: false },
      { rowKey: 'adj_severance', label: 'h) Severance', valueType: 'currency', isCalculated: false },
      { rowKey: 'adj_above_market_rent', label: 'i) Above-Market Rent', valueType: 'currency', isCalculated: false },
      { rowKey: 'adj_normalize_owner_comp', label: 'j) Normalize Owner Comp', valueType: 'currency', isCalculated: false },
      { rowKey: 'total_diligence_adjustments', label: 'Total Diligence Adjustments', valueType: 'currency', isCalculated: true, formula: 'adj_professional_fees_onetime+adj_executive_search+adj_facility_relocation+adj_severance+adj_above_market_rent+adj_normalize_owner_comp' },
      { rowKey: 'diligence_adjusted_ebitda', label: 'Diligence-Adjusted EBITDA', valueType: 'currency', isCalculated: true, formula: 'mgmt_adjusted_ebitda+total_diligence_adjustments' },
    ],
  },

  // ──────────────────────────────────────────────────────────────────
  // INCOME STATEMENT
  // ──────────────────────────────────────────────────────────────────
  {
    key: 'income-statement',
    displayName: 'Income Statement',
    ragQuery:
      'income statement profit loss revenue gross profit operating expenses sales marketing payroll COGS cost of goods sold net income total revenue operating income',
    ragQueries: [
      'gross product sales returns allowances promotional discounts net product sales shipping revenue total net revenue income statement',
      'cost of goods sold COGS product cost warehousing fulfillment shipping freight inventory adjustments gross profit',
      'operating expenses payroll salaries wages benefits sales marketing rent technology insurance payment processing travel general admin',
      'operating income EBITDA interest expense depreciation amortization income before tax income tax net income',
    ],
    requiredRows: [
      { rowKey: 'gross_product_sales', label: 'Gross Product Sales', valueType: 'currency', isCalculated: false },
      { rowKey: 'returns_allowances', label: 'Returns & Allowances', valueType: 'currency', isCalculated: false },
      { rowKey: 'promotional_discounts', label: 'Promotional Discounts', valueType: 'currency', isCalculated: false },
      { rowKey: 'net_product_sales', label: 'Net Product Sales', valueType: 'currency', isCalculated: true, formula: 'gross_product_sales-returns_allowances-promotional_discounts' },
      { rowKey: 'shipping_revenue', label: 'Shipping Revenue', valueType: 'currency', isCalculated: false },
      { rowKey: 'other_revenue', label: 'Other Revenue', valueType: 'currency', isCalculated: false },
      { rowKey: 'total_net_revenue', label: 'Total Net Revenue', valueType: 'currency', isCalculated: true, formula: 'net_product_sales+shipping_revenue+other_revenue' },
      { rowKey: 'product_cogs', label: 'Product Cost of Goods Sold', valueType: 'currency', isCalculated: false },
      { rowKey: 'warehousing_fulfillment', label: 'Warehousing & Fulfillment', valueType: 'currency', isCalculated: false },
      { rowKey: 'shipping_freight_out', label: 'Shipping & Freight Out', valueType: 'currency', isCalculated: false },
      { rowKey: 'inventory_adjustments', label: 'Inventory Adjustments', valueType: 'currency', isCalculated: false },
      { rowKey: 'total_cogs', label: 'Total Cost of Goods Sold', valueType: 'currency', isCalculated: true, formula: 'product_cogs+warehousing_fulfillment+shipping_freight_out+inventory_adjustments' },
      { rowKey: 'gross_profit', label: 'Gross Profit', valueType: 'currency', isCalculated: true, formula: 'total_net_revenue-total_cogs' },
      { rowKey: 'sales_marketing', label: 'Sales & Marketing', valueType: 'currency', isCalculated: false },
      { rowKey: 'salaries_wages_payroll', label: 'Salaries, Wages & Payroll Tax', valueType: 'currency', isCalculated: false },
      { rowKey: 'employee_benefits_401k', label: 'Employee Benefits & 401k', valueType: 'currency', isCalculated: false },
      { rowKey: 'rent_occupancy', label: 'Rent & Occupancy', valueType: 'currency', isCalculated: false },
      { rowKey: 'professional_services', label: 'Professional Services', valueType: 'currency', isCalculated: false },
      { rowKey: 'technology_software', label: 'Technology & Software', valueType: 'currency', isCalculated: false },
      { rowKey: 'insurance', label: 'Insurance', valueType: 'currency', isCalculated: false },
      { rowKey: 'payment_processing_fees', label: 'Payment Processing Fees', valueType: 'currency', isCalculated: false },
      { rowKey: 'travel_entertainment', label: 'Travel & Entertainment', valueType: 'currency', isCalculated: false },
      { rowKey: 'office_general_admin', label: 'Office & General Admin', valueType: 'currency', isCalculated: false },
      { rowKey: 'total_operating_expenses', label: 'Total Operating Expenses', valueType: 'currency', isCalculated: true, formula: 'sales_marketing+salaries_wages_payroll+employee_benefits_401k+rent_occupancy+professional_services+technology_software+insurance+payment_processing_fees+travel_entertainment+office_general_admin' },
      { rowKey: 'operating_income_ebitda', label: 'Operating Income (EBITDA)', valueType: 'currency', isCalculated: true, formula: 'gross_profit-total_operating_expenses' },
      { rowKey: 'interest_expense', label: 'Interest Expense', valueType: 'currency', isCalculated: false },
      { rowKey: 'depreciation_amortization', label: 'Depreciation & Amortization', valueType: 'currency', isCalculated: false },
      { rowKey: 'income_before_tax', label: 'Income Before Tax', valueType: 'currency', isCalculated: true, formula: 'operating_income_ebitda-interest_expense-depreciation_amortization' },
      { rowKey: 'tax_provision', label: 'Tax Provision', valueType: 'currency', isCalculated: false },
      { rowKey: 'net_income', label: 'Net Income', valueType: 'currency', isCalculated: true, formula: 'income_before_tax-tax_provision' },
    ],
  },

  // ──────────────────────────────────────────────────────────────────
  // BALANCE SHEET
  // ──────────────────────────────────────────────────────────────────
  {
    key: 'balance-sheet',
    displayName: 'Balance Sheet',
    ragQuery:
      'balance sheet assets liabilities equity accounts receivable inventory accounts payable debt retained earnings cash current assets fixed assets',
    ragQueries: [
      'cash equivalents accounts receivable allowance doubtful inventory prepaid other current assets balance sheet',
      'property plant equipment accumulated depreciation fixed assets net intangible assets total assets',
      'accounts payable credit cards accrued payroll accrued expenses sales tax deferred revenue current portion long-term debt current liabilities',
      'long-term debt net equity common stock additional paid-in capital retained earnings net income total liabilities stockholders equity',
    ],
    requiredRows: [
      { rowKey: 'cash_equivalents', label: 'Cash & Cash Equivalents', valueType: 'currency', isCalculated: false },
      { rowKey: 'accounts_receivable', label: 'Accounts Receivable', valueType: 'currency', isCalculated: false },
      { rowKey: 'allowance_doubtful', label: 'Allowance for Doubtful Accounts', valueType: 'currency', isCalculated: false },
      { rowKey: 'net_accounts_receivable', label: 'Net Accounts Receivable', valueType: 'currency', isCalculated: true, formula: 'accounts_receivable-allowance_doubtful' },
      { rowKey: 'inventory_raw_materials', label: 'Inventory - Raw Materials', valueType: 'currency', isCalculated: false },
      { rowKey: 'inventory_finished_goods', label: 'Inventory - Finished Goods', valueType: 'currency', isCalculated: false },
      { rowKey: 'total_inventory', label: 'Total Inventory', valueType: 'currency', isCalculated: true, formula: 'inventory_raw_materials+inventory_finished_goods' },
      { rowKey: 'prepaid_expenses', label: 'Prepaid Expenses', valueType: 'currency', isCalculated: false },
      { rowKey: 'other_current_assets', label: 'Other Current Assets', valueType: 'currency', isCalculated: false },
      { rowKey: 'total_current_assets', label: 'Total Current Assets', valueType: 'currency', isCalculated: true, formula: 'cash_equivalents+net_accounts_receivable+total_inventory+prepaid_expenses+other_current_assets' },
      { rowKey: 'ppe_gross', label: 'Property, Plant & Equipment', valueType: 'currency', isCalculated: false },
      { rowKey: 'accumulated_depreciation', label: 'Less: Accumulated Depreciation', valueType: 'currency', isCalculated: false },
      { rowKey: 'net_fixed_assets', label: 'Net Fixed Assets', valueType: 'currency', isCalculated: true, formula: 'ppe_gross-accumulated_depreciation' },
      { rowKey: 'intangible_assets', label: 'Intangible Assets - Software', valueType: 'currency', isCalculated: false },
      { rowKey: 'total_assets', label: 'Total Assets', valueType: 'currency', isCalculated: true, formula: 'total_current_assets+net_fixed_assets+intangible_assets' },
      { rowKey: 'accounts_payable_trade', label: 'Accounts Payable - Trade', valueType: 'currency', isCalculated: false },
      { rowKey: 'credit_cards_payable', label: 'Credit Cards Payable', valueType: 'currency', isCalculated: false },
      { rowKey: 'accrued_payroll_benefits', label: 'Accrued Payroll & Benefits', valueType: 'currency', isCalculated: false },
      { rowKey: 'accrued_expenses_other', label: 'Accrued Expenses - Other', valueType: 'currency', isCalculated: false },
      { rowKey: 'sales_tax_payable', label: 'Sales Tax Payable', valueType: 'currency', isCalculated: false },
      { rowKey: 'deferred_revenue', label: 'Deferred Revenue', valueType: 'currency', isCalculated: false },
      { rowKey: 'current_portion_lt_debt', label: 'Current Portion - LT Debt', valueType: 'currency', isCalculated: false },
      { rowKey: 'total_current_liabilities', label: 'Total Current Liabilities', valueType: 'currency', isCalculated: true, formula: 'accounts_payable_trade+credit_cards_payable+accrued_payroll_benefits+accrued_expenses_other+sales_tax_payable+deferred_revenue+current_portion_lt_debt' },
      { rowKey: 'lt_debt_net', label: 'Long-Term Debt, Net of Current', valueType: 'currency', isCalculated: false },
      { rowKey: 'total_liabilities', label: 'Total Liabilities', valueType: 'currency', isCalculated: true, formula: 'total_current_liabilities+lt_debt_net' },
      { rowKey: 'common_stock', label: 'Common Stock', valueType: 'currency', isCalculated: false },
      { rowKey: 'additional_paid_in_capital', label: 'Additional Paid-in Capital', valueType: 'currency', isCalculated: false },
      { rowKey: 'retained_earnings', label: 'Retained Earnings', valueType: 'currency', isCalculated: false },
      { rowKey: 'current_year_net_income', label: 'Current Year Net Income', valueType: 'currency', isCalculated: false },
      { rowKey: 'total_equity', label: "Total Stockholders' Equity", valueType: 'currency', isCalculated: true, formula: 'common_stock+additional_paid_in_capital+retained_earnings+current_year_net_income' },
      { rowKey: 'total_liabilities_equity', label: 'Total Liabilities & Equity', valueType: 'currency', isCalculated: true, formula: 'total_liabilities+total_equity' },
    ],
  },

  // ──────────────────────────────────────────────────────────────────
  // SALES BY CHANNEL
  // ──────────────────────────────────────────────────────────────────
  {
    key: 'sales-channel',
    displayName: 'Sales by Channel',
    ragQuery:
      'revenue by product line channel SKU product mix sales breakdown by category segment service type product revenue breakdown top products units sold average selling price',
    ragQueries: [
      'revenue by product line channel SKU product mix top products sales breakdown by category',
      'product segment service type revenue percentage of total sales units sold average selling price',
    ],
    requiredRows: [
      { rowKey: 'product_line_1', label: 'Product / Channel 1 (largest)', valueType: 'currency', isCalculated: false },
      { rowKey: 'product_line_2', label: 'Product / Channel 2', valueType: 'currency', isCalculated: false },
      { rowKey: 'product_line_3', label: 'Product / Channel 3', valueType: 'currency', isCalculated: false },
      { rowKey: 'product_line_4', label: 'Product / Channel 4', valueType: 'currency', isCalculated: false },
      { rowKey: 'product_line_5', label: 'Product / Channel 5', valueType: 'currency', isCalculated: false },
      { rowKey: 'product_line_6', label: 'Product / Channel 6', valueType: 'currency', isCalculated: false },
      { rowKey: 'product_line_7', label: 'Product / Channel 7', valueType: 'currency', isCalculated: false },
      { rowKey: 'product_line_8', label: 'Product / Channel 8', valueType: 'currency', isCalculated: false },
      { rowKey: 'product_line_9', label: 'Product / Channel 9', valueType: 'currency', isCalculated: false },
      { rowKey: 'product_line_10', label: 'Product / Channel 10', valueType: 'currency', isCalculated: false },
      { rowKey: 'all_other_products', label: 'All Other Products / Channels', valueType: 'currency', isCalculated: false },
      { rowKey: 'total_product_sales', label: 'Total Product Sales', valueType: 'currency', isCalculated: true },
    ],
  },

  // ──────────────────────────────────────────────────────────────────
  // MARGINS BY MONTH
  // Use month names as periods instead of fiscal years.
  // ──────────────────────────────────────────────────────────────────
  {
    key: 'margins-month',
    displayName: 'Margins by Month',
    ragQuery:
      'monthly revenue gross margin operating expenses by month COGS monthly EBITDA monthly profit loss monthly income statement trailing twelve months',
    ragQueries: [
      'monthly revenue gross margin monthly COGS cost of goods sold by month trailing twelve months TTM',
      'monthly operating expenses payroll rent month-by-month gross profit net profit income',
    ],
    // Override to extract monthly data instead of fiscal-year periods
    overridePeriods: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],
    requiredRows: [
      { rowKey: 'revenue', label: 'Revenue', valueType: 'currency', isCalculated: false },
      { rowKey: 'cogs', label: 'Cost of Goods Sold', valueType: 'currency', isCalculated: false },
      { rowKey: 'opex', label: 'Operating Expenses', valueType: 'currency', isCalculated: false },
      { rowKey: 'gross_profit', label: 'Gross Profit', valueType: 'currency', isCalculated: true, formula: 'revenue-cogs' },
      { rowKey: 'gross_margin_pct', label: 'Gross Margin %', valueType: 'percent', isCalculated: true },
      { rowKey: 'net_profit', label: 'Net Profit', valueType: 'currency', isCalculated: true, formula: 'gross_profit-opex' },
      { rowKey: 'net_margin_pct', label: 'Net Margin %', valueType: 'percent', isCalculated: true },
    ],
  },

  // ──────────────────────────────────────────────────────────────────
  // PROOF OF CASH
  // ──────────────────────────────────────────────────────────────────
  {
    key: 'proof-cash',
    displayName: 'Proof of Cash',
    ragQuery:
      'bank deposits bank statements cash receipts accounts receivable beginning ending balance bank reconciliation proof of cash monthly deposits collections',
    ragQueries: [
      'bank deposits cash receipts monthly collections bank statements beginning ending cash balance',
      'accounts receivable beginning AR ending AR revenue per general ledger non-revenue deposits cash reconciliation',
    ],
    requiredRows: [
      { rowKey: 'bank_deposits', label: 'Bank Deposits', valueType: 'currency', isCalculated: false },
      { rowKey: 'beginning_ar', label: 'Beginning AR', valueType: 'currency', isCalculated: false },
      { rowKey: 'ending_ar', label: 'Ending AR', valueType: 'currency', isCalculated: false },
      { rowKey: 'revenue_gl', label: 'Revenue per GL', valueType: 'currency', isCalculated: false },
      { rowKey: 'non_rev_deposits', label: 'Non-Revenue Deposits', valueType: 'currency', isCalculated: false },
      { rowKey: 'variance', label: 'Variance', valueType: 'currency', isCalculated: true, formula: 'bank_deposits-revenue_gl-non_rev_deposits' },
    ],
  },

  // ──────────────────────────────────────────────────────────────────
  // WORKING CAPITAL
  // ──────────────────────────────────────────────────────────────────
  {
    key: 'working-capital',
    displayName: 'Working Capital',
    ragQuery:
      'working capital accounts receivable inventory accounts payable normalized working capital days outstanding DSO DIO DPO net working capital current assets current liabilities',
    ragQueries: [
      'accounts receivable inventory prepaid other current assets working capital normalized',
      'accounts payable accrued expenses credit cards sales tax deferred revenue net working capital DSO DIO DPO',
    ],
    requiredRows: [
      { rowKey: 'wc_accounts_receivable', label: 'Accounts Receivable', valueType: 'currency', isCalculated: false },
      { rowKey: 'wc_inventory', label: 'Inventory', valueType: 'currency', isCalculated: false },
      { rowKey: 'wc_prepaid_expenses', label: 'Prepaid Expenses', valueType: 'currency', isCalculated: false },
      { rowKey: 'wc_other_current_assets', label: 'Other Current Assets', valueType: 'currency', isCalculated: false },
      { rowKey: 'wc_accounts_payable', label: 'Accounts Payable', valueType: 'currency', isCalculated: false },
      { rowKey: 'wc_accrued_expenses', label: 'Accrued Expenses', valueType: 'currency', isCalculated: false },
      { rowKey: 'wc_credit_cards_payable', label: 'Credit Cards Payable', valueType: 'currency', isCalculated: false },
      { rowKey: 'wc_sales_tax_payable', label: 'Sales Tax Payable', valueType: 'currency', isCalculated: false },
      { rowKey: 'wc_deferred_revenue', label: 'Deferred Revenue', valueType: 'currency', isCalculated: false },
      { rowKey: 'wc_adjusted_nwc', label: 'Adjusted Net Working Capital', valueType: 'currency', isCalculated: true, formula: 'wc_accounts_receivable+wc_inventory+wc_prepaid_expenses+wc_other_current_assets-wc_accounts_payable-wc_accrued_expenses-wc_credit_cards_payable-wc_sales_tax_payable-wc_deferred_revenue' },
    ],
  },

  // ──────────────────────────────────────────────────────────────────
  // COGS VENDORS
  // ──────────────────────────────────────────────────────────────────
  {
    key: 'cogs-vendors',
    displayName: 'COGS Vendors',
    ragQuery:
      'top vendors by spend supplier payments accounts payable vendor report purchases by supplier cost of goods sold vendor breakdown vendor spend analysis largest vendor payments check disbursements',
    ragQueries: [
      'top vendors by spend purchases by supplier accounts payable vendor report COGS vendor breakdown',
      'largest vendors check disbursements ACH wire transfers vendor concentration accounts payable aging',
    ],
    requiredRows: [
      { rowKey: 'vendor_1', label: 'Vendor 1 (largest spend)', valueType: 'currency', isCalculated: false },
      { rowKey: 'vendor_2', label: 'Vendor 2', valueType: 'currency', isCalculated: false },
      { rowKey: 'vendor_3', label: 'Vendor 3', valueType: 'currency', isCalculated: false },
      { rowKey: 'vendor_4', label: 'Vendor 4', valueType: 'currency', isCalculated: false },
      { rowKey: 'vendor_5', label: 'Vendor 5', valueType: 'currency', isCalculated: false },
      { rowKey: 'vendor_6', label: 'Vendor 6', valueType: 'currency', isCalculated: false },
      { rowKey: 'vendor_7', label: 'Vendor 7', valueType: 'currency', isCalculated: false },
      { rowKey: 'vendor_8', label: 'Vendor 8', valueType: 'currency', isCalculated: false },
      { rowKey: 'vendor_all_other', label: 'All Other Vendors', valueType: 'currency', isCalculated: false },
      { rowKey: 'total_cogs', label: 'Total COGS', valueType: 'currency', isCalculated: true },
    ],
  },

  // ──────────────────────────────────────────────────────────────────
  // AP/ACCRUAL TESTING
  // ──────────────────────────────────────────────────────────────────
  {
    key: 'testing',
    displayName: 'AP/Accrual Testing',
    ragQuery:
      'accounts payable check payments disbursements cutoff testing accrual testing subsequent disbursements post period payments check register ACH wire transfers',
    ragQueries: [
      'accounts payable disbursements tested cutoff testing subsequent check payments post period',
      'accrual testing period allocation future period not accrued total error AP testing',
    ],
    requiredRows: [
      { rowKey: 'total_disbursements', label: 'Total Disbursements Tested', valueType: 'currency', isCalculated: false },
      { rowKey: 'pct_total_ap', label: '% of Total AP', valueType: 'percent', isCalculated: false },
      { rowKey: 'period_amount', label: 'Allocated to Period', valueType: 'currency', isCalculated: false },
      { rowKey: 'future_period_amount', label: 'Future Period', valueType: 'currency', isCalculated: false },
      { rowKey: 'not_accrued_amount', label: 'Not Accrued', valueType: 'currency', isCalculated: false },
      { rowKey: 'total_error', label: 'Total Error', valueType: 'currency', isCalculated: true, formula: 'future_period_amount+not_accrued_amount' },
    ],
  },

  // ──────────────────────────────────────────────────────────────────
  // NET DEBT & DEBT-LIKE ITEMS
  // ──────────────────────────────────────────────────────────────────
  {
    key: 'net-debt',
    displayName: 'Net Debt & Debt-Like Items',
    ragQuery:
      'net debt debt-like items capital leases accrued vacation aged accounts payable deferred revenue customer deposits cash and equivalents closing balance sheet debt schedule',
    ragQueries: [
      'total debt term loan revolving credit capital leases long-term debt schedule closing balance sheet',
      'debt-like items accrued vacation aged accounts payable deferred revenue customer deposits net debt cash',
    ],
    requiredRows: [
      { rowKey: 'reported_debt', label: 'Reported Debt (Term Loan + Revolver)', valueType: 'currency', isCalculated: false },
      { rowKey: 'capital_leases', label: 'Capital Leases', valueType: 'currency', isCalculated: false },
      { rowKey: 'unpaid_accrued_vacation', label: 'Unpaid Accrued Vacation', valueType: 'currency', isCalculated: false },
      { rowKey: 'aged_accounts_payable', label: 'Aged Accounts Payable (>90 days)', valueType: 'currency', isCalculated: false },
      { rowKey: 'deferred_revenue', label: 'Deferred Revenue', valueType: 'currency', isCalculated: false },
      { rowKey: 'customer_deposits', label: 'Customer Deposits', valueType: 'currency', isCalculated: false },
      { rowKey: 'total_debt_like_items', label: 'Total Debt & Debt-Like Items', valueType: 'currency', isCalculated: true, formula: 'reported_debt+capital_leases+unpaid_accrued_vacation+aged_accounts_payable+deferred_revenue+customer_deposits' },
      { rowKey: 'cash_and_equivalents', label: 'Less: Cash & Cash Equivalents', valueType: 'currency', isCalculated: false },
      { rowKey: 'net_debt', label: 'Net Debt (Cash)', valueType: 'currency', isCalculated: true, formula: 'total_debt_like_items-cash_and_equivalents' },
    ],
  },

  // ──────────────────────────────────────────────────────────────────
  // REVENUE & CUSTOMER CONCENTRATION
  // ──────────────────────────────────────────────────────────────────
  {
    key: 'customer-concentration',
    displayName: 'Revenue & Customer Concentration',
    ragQuery:
      'customer concentration revenue by customer top customers accounts receivable aging customer list revenue breakdown by client',
    ragQueries: [
      'revenue by customer top customers client concentration sales breakdown percentage of revenue',
      'largest customers accounts receivable aging customer list revenue by account',
    ],
    requiredRows: [
      { rowKey: 'customer_1', label: 'Customer 1', valueType: 'currency', isCalculated: false },
      { rowKey: 'customer_2', label: 'Customer 2', valueType: 'currency', isCalculated: false },
      { rowKey: 'customer_3', label: 'Customer 3', valueType: 'currency', isCalculated: false },
      { rowKey: 'customer_4', label: 'Customer 4', valueType: 'currency', isCalculated: false },
      { rowKey: 'customer_5', label: 'Customer 5', valueType: 'currency', isCalculated: false },
      { rowKey: 'customer_6', label: 'Customer 6', valueType: 'currency', isCalculated: false },
      { rowKey: 'customer_7', label: 'Customer 7', valueType: 'currency', isCalculated: false },
      { rowKey: 'customer_8', label: 'Customer 8', valueType: 'currency', isCalculated: false },
      { rowKey: 'customer_9', label: 'Customer 9', valueType: 'currency', isCalculated: false },
      { rowKey: 'customer_10', label: 'Customer 10', valueType: 'currency', isCalculated: false },
      { rowKey: 'all_other_customers', label: 'All Other Customers', valueType: 'currency', isCalculated: false },
      { rowKey: 'total_revenue', label: 'Total Revenue', valueType: 'currency', isCalculated: true },
    ],
  },

  // ──────────────────────────────────────────────────────────────────
  // CASH CONVERSION ANALYSIS
  // ──────────────────────────────────────────────────────────────────
  {
    key: 'cash-conversion',
    displayName: 'Cash Conversion Analysis',
    ragQuery:
      'free cash flow capex capital expenditures working capital change taxes paid interest paid cash flow from operations EBITDA to FCF conversion cash conversion cycle',
    ragQueries: [
      'free cash flow EBITDA capex capital expenditures cash conversion FCF',
      'taxes paid cash interest paid working capital change operating cash flow cash flow from operations',
    ],
    requiredRows: [
      { rowKey: 'ebitda', label: 'Adjusted EBITDA', valueType: 'currency', isCalculated: false },
      { rowKey: 'capex', label: 'Less: Capital Expenditures', valueType: 'currency', isCalculated: false },
      { rowKey: 'change_in_working_capital', label: 'Less: Change in Working Capital', valueType: 'currency', isCalculated: false },
      { rowKey: 'taxes_paid_cash', label: 'Less: Taxes Paid (Cash)', valueType: 'currency', isCalculated: false },
      { rowKey: 'interest_paid_cash', label: 'Less: Interest Paid (Cash)', valueType: 'currency', isCalculated: false },
      { rowKey: 'free_cash_flow', label: 'Free Cash Flow', valueType: 'currency', isCalculated: true, formula: 'ebitda-capex-change_in_working_capital-taxes_paid_cash-interest_paid_cash' },
      { rowKey: 'fcf_to_ebitda_ratio', label: 'FCF / EBITDA Conversion', valueType: 'percent', isCalculated: true },
    ],
  },

  // ──────────────────────────────────────────────────────────────────
  // PROOF OF REVENUE (THREE-WAY MATCH)
  // ──────────────────────────────────────────────────────────────────
  {
    key: 'proof-revenue',
    displayName: 'Proof of Revenue (Three-Way Match)',
    ragQuery:
      'revenue reconciliation GL revenue tax return revenue bank deposits three-way match proof of revenue variance general ledger vs tax return vs bank statements',
    ragQueries: [
      'general ledger GL revenue total revenue reported income statement reconciliation three-way match',
      'tax return gross receipts reported revenue bank deposit revenue variance proof of revenue',
    ],
    requiredRows: [
      { rowKey: 'gl_revenue', label: 'GL Revenue', valueType: 'currency', isCalculated: false },
      { rowKey: 'tax_return_revenue', label: 'Tax Return Revenue', valueType: 'currency', isCalculated: false },
      { rowKey: 'bank_deposit_revenue', label: 'Bank Deposit Revenue', valueType: 'currency', isCalculated: false },
      { rowKey: 'variance_gl_vs_tax', label: 'Variance: GL vs. Tax Return', valueType: 'currency', isCalculated: true, formula: 'gl_revenue-tax_return_revenue' },
      { rowKey: 'variance_gl_vs_bank', label: 'Variance: GL vs. Bank Deposits', valueType: 'currency', isCalculated: true, formula: 'gl_revenue-bank_deposit_revenue' },
      { rowKey: 'pct_variance_gl_tax', label: '% Variance GL vs. Tax', valueType: 'percent', isCalculated: true },
    ],
  },

  // ──────────────────────────────────────────────────────────────────
  // RUN-RATE & PRO-FORMA ADJUSTMENTS
  // ──────────────────────────────────────────────────────────────────
  {
    key: 'run-rate',
    displayName: 'Run-Rate & Pro-Forma Adjustments',
    ragQuery:
      'run rate pro forma adjustments annualized revenue new contracts full year effect price increases lost revenue pro forma EBITDA forward projections normalized earnings',
    ragQueries: [
      'pro forma revenue annualized run rate new contracts full year price increases adjustments forward',
      'pro forma EBITDA lost revenue churn normalized earnings forward projections TTM annualized',
    ],
    requiredRows: [
      { rowKey: 'ttm_revenue_base', label: 'TTM Revenue (Base)', valueType: 'currency', isCalculated: false },
      { rowKey: 'new_contract_value', label: 'New Contract Value (Annualized)', valueType: 'currency', isCalculated: false },
      { rowKey: 'full_year_hire_effect', label: 'Full-Year Hire Effect', valueType: 'currency', isCalculated: false },
      { rowKey: 'price_increase_effect', label: 'Price Increase Effect', valueType: 'currency', isCalculated: false },
      { rowKey: 'lost_revenue_adjustment', label: 'Less: Lost/Churned Revenue', valueType: 'currency', isCalculated: false },
      { rowKey: 'pro_forma_revenue', label: 'Pro Forma Revenue', valueType: 'currency', isCalculated: true, formula: 'ttm_revenue_base+new_contract_value+full_year_hire_effect+price_increase_effect+lost_revenue_adjustment' },
      { rowKey: 'pro_forma_ebitda', label: 'Pro Forma EBITDA', valueType: 'currency', isCalculated: true },
    ],
  },
]

export function getSectionConfig(key: string): SectionConfig | undefined {
  return SECTION_CONFIGS.find((s) => s.key === key)
}
