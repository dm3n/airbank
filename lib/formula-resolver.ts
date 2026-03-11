/**
 * Formula resolver for workbook cell calculations.
 *
 * Formulas use simple arithmetic (+, -, *, /) over row keys.
 * Dependency order is determined by topological sort so calculated rows
 * that depend on other calculated rows resolve correctly.
 *
 * Example formula: "net_income+interest_expense+tax_provision+depreciation+amortization"
 */

import { SECTION_CONFIGS, RowDef } from './section-prompts'

export interface CellValue {
  rowKey: string
  period: string
  rawValue: number | null
}

/**
 * Evaluate a formula string given a map of rowKey → value.
 * Returns null if any operand is null.
 */
function evalFormula(formula: string, values: Map<string, number | null>): number | null {
  // Tokenise: split into [operand, operator, operand, ...] using +/-/* / preserving sign
  // We support: a+b-c*d/e  (left-to-right, no parens needed for current formulas)
  // Strategy: replace row keys with their numeric values then eval via simple parser.

  // Build a regex that matches any known rowKey
  const tokens: Array<{ type: 'key' | 'op'; value: string }> = []
  let remaining = formula.trim()

  while (remaining.length > 0) {
    // Try operator first if we have prior tokens
    const opMatch = remaining.match(/^([+\-*/])(.*)/)
    if (opMatch && tokens.length > 0) {
      tokens.push({ type: 'op', value: opMatch[1] })
      remaining = opMatch[2]
      continue
    }
    // Try identifier (row key)
    const idMatch = remaining.match(/^([a-zA-Z_][a-zA-Z0-9_]*)(.*)/)
    if (idMatch) {
      tokens.push({ type: 'key', value: idMatch[1] })
      remaining = idMatch[2]
      continue
    }
    // Try numeric literal
    const numMatch = remaining.match(/^([0-9]*\.?[0-9]+)(.*)/)
    if (numMatch) {
      tokens.push({ type: 'key', value: numMatch[1] })
      remaining = numMatch[2]
      continue
    }
    // Skip whitespace
    remaining = remaining.slice(1)
  }

  if (tokens.length === 0) return null

  // Evaluate left-to-right
  let result: number | null = null
  let pendingOp: string = '+'

  for (const token of tokens) {
    if (token.type === 'op') {
      pendingOp = token.value
      continue
    }

    // Resolve value
    let val: number | null
    const asNum = parseFloat(token.value)
    if (!isNaN(asNum) && /^[0-9]/.test(token.value)) {
      val = asNum
    } else {
      const mapVal = values.get(token.value)
      val = mapVal === undefined ? null : mapVal
    }

    if (val === null) return null

    if (result === null) {
      result = val
    } else {
      switch (pendingOp) {
        case '+': result += val; break
        case '-': result -= val; break
        case '*': result *= val; break
        case '/': result = val !== 0 ? result / val : null; break
      }
    }
  }

  return result
}

/**
 * Topological sort of RowDefs so that calculated rows are resolved after
 * their dependencies.
 */
function topoSort(rows: RowDef[]): RowDef[] {
  const keySet = new Set(rows.map(r => r.rowKey))
  const sorted: RowDef[] = []
  const visited = new Set<string>()

  function visit(row: RowDef) {
    if (visited.has(row.rowKey)) return
    visited.add(row.rowKey)

    if (row.formula) {
      // Extract dependency keys from formula
      const deps = row.formula.match(/[a-zA-Z_][a-zA-Z0-9_]*/g) ?? []
      for (const dep of deps) {
        if (keySet.has(dep)) {
          const depRow = rows.find(r => r.rowKey === dep)
          if (depRow) visit(depRow)
        }
      }
    }

    sorted.push(row)
  }

  for (const row of rows) visit(row)
  return sorted
}

/**
 * Given a section key and a flat list of current cell values, resolve all
 * formula cells and return the recalculated values.
 *
 * Only returns cells whose value changed (or became computable for the first time).
 */
export function resolveFormulas(
  sectionKey: string,
  cells: CellValue[]
): CellValue[] {
  const config = SECTION_CONFIGS.find(s => s.key === sectionKey)
  if (!config) return []

  const formulaRows = config.requiredRows.filter(r => r.formula)
  if (formulaRows.length === 0) return []

  // Collect all distinct periods present in the cell data
  const periods = [...new Set(cells.map(c => c.period))]

  // Build lookup: rowKey → period → value
  const lookup = new Map<string, number | null>()
  for (const cell of cells) {
    lookup.set(`${cell.rowKey}::${cell.period}`, cell.rawValue)
  }

  const sorted = topoSort(config.requiredRows)
  const updated: CellValue[] = []

  for (const period of periods) {
    // Materialise a working map for this period
    const vals = new Map<string, number | null>()
    for (const row of config.requiredRows) {
      const v = lookup.get(`${row.rowKey}::${period}`)
      vals.set(row.rowKey, v === undefined ? null : v)
    }

    // Resolve in topo order
    for (const row of sorted) {
      if (!row.formula) continue
      const newVal = evalFormula(row.formula, vals)
      const oldVal = vals.get(row.rowKey) ?? null

      // Update working map for subsequent deps
      vals.set(row.rowKey, newVal)

      // Only emit if value changed
      if (newVal !== oldVal) {
        updated.push({ rowKey: row.rowKey, period, rawValue: newVal })
      }
    }
  }

  return updated
}
