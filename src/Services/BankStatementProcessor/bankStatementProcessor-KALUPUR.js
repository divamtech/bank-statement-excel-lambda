/**
 * Processes Kalupur Bank Statement Excel
 * Layout:
 *   Row(s) 1..N : meta key/value (Bank Name, Customer Name, Account Number, Account Type, From Date, To Date)
 *   Header row  : DATE | VALUE DATE | NARRATION | CHEQ/REF NO. | WITHDRAWAL | DEPOSIT | BALANCE
 *   Then        : transaction rows, last row is "TOTAL" summary (skipped)
 *
 * @param {Array} rawData - Array of objects from bank statement Excel
 * @returns {Object} Processed bank statement data
 */
function processBankStatement(rawData) {
  const bank_details = {
    bank_name: null,
    opening_balance: 0,
    ifsc: null,
    address: null,
    city: null,
    account_no: null,
    account_holder_name: null,
    branch_name: null,
    branch_code: null,
  }

  const headerKeywords = ['date', 'narration', 'cheq', 'ref', 'withdrawal', 'deposit', 'balance', 'value date']

  let headerRowIndex = -1
  const keyMapping = {}
  const transactions = []
  let voucher_number = 1

  function parseAndFormatDate(dateString) {
    if (dateString === null || dateString === undefined) return null
    if (dateString instanceof Date && !isNaN(dateString.getTime())) {
      const y = dateString.getFullYear()
      const m = String(dateString.getMonth() + 1).padStart(2, '0')
      const d = String(dateString.getDate()).padStart(2, '0')
      return `${y}-${m}-${d}`
    }
    const str = String(dateString).trim()
    if (!str) return null

    if (/^\d+(\.\d+)?$/.test(str)) {
      const serial = parseInt(str, 10)
      const epochDays = serial - 25569
      const date = new Date(epochDays * 86400000)
      const y = date.getUTCFullYear()
      const m = String(date.getUTCMonth() + 1).padStart(2, '0')
      const d = String(date.getUTCDate()).padStart(2, '0')
      return `${y}-${m}-${d}`
    }

    const monthNames = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec']
    const formats = [
      { regex: /^(\d{2})-(\d{2})-(\d{4})$/, fn: (m) => `${m[3]}-${m[2]}-${m[1]}` },
      { regex: /^(\d{2})\/(\d{2})\/(\d{4})$/, fn: (m) => `${m[3]}-${m[2]}-${m[1]}` },
      { regex: /^(\d{2})\.(\d{2})\.(\d{4})$/, fn: (m) => `${m[3]}-${m[2]}-${m[1]}` },
      { regex: /^(\d{4})-(\d{2})-(\d{2})$/, fn: (m) => `${m[1]}-${m[2]}-${m[3]}` },
      {
        regex: /^(\d{2})-([A-Za-z]{3})-(\d{4})$/,
        fn: (m) => {
          const idx = monthNames.indexOf(m[2].toLowerCase())
          if (idx === -1) return null
          return `${m[3]}-${String(idx + 1).padStart(2, '0')}-${m[1]}`
        },
      },
    ]

    for (const f of formats) {
      const m = str.match(f.regex)
      if (m) {
        const out = f.fn(m)
        if (out) return out
      }
    }
    return null
  }

  function extractNumber(value) {
    if (value === null || value === undefined || value === '') return null
    if (typeof value === 'number') return isNaN(value) ? null : value
    const cleaned = String(value).replace(/[^\d.-]/g, '')
    if (!cleaned) return null
    const n = parseFloat(cleaned)
    return isNaN(n) ? null : n
  }

  function rowValues(row) {
    return Object.keys(row).map((k) => ({ key: k, value: row[k] }))
  }

  function findNextNonNullValue(cells, fromIndex) {
    for (let i = fromIndex + 1; i < cells.length; i++) {
      const v = cells[i].value
      if (v !== null && v !== undefined && String(v).trim() !== '') return v
    }
    return null
  }

  for (let i = 0; i < Math.min(25, rawData.length); i++) {
    const row = rawData[i]
    if (!row) continue
    const cells = rowValues(row)
    for (let c = 0; c < cells.length; c++) {
      const raw = cells[c].value
      if (raw === null || raw === undefined) continue
      const label = String(raw).trim().toLowerCase()
      if (!label) continue

      if (label === 'bank name' || label.startsWith('bank name')) {
        const v = findNextNonNullValue(cells, c)
        if (v !== null && !bank_details.bank_name) bank_details.bank_name = String(v).trim()
      } else if (
        label === 'customer name' ||
        label.startsWith('customer name') ||
        label === 'account holder name' ||
        label.startsWith('holder name') ||
        label === 'name'
      ) {
        const v = findNextNonNullValue(cells, c)
        if (v !== null && !bank_details.account_holder_name) bank_details.account_holder_name = String(v).trim()
      } else if (label === 'account number' || label.startsWith('account number') || label.startsWith('account no') || label === 'a/c no') {
        const v = findNextNonNullValue(cells, c)
        if (v !== null && !bank_details.account_no) bank_details.account_no = String(v).trim()
      } else if (label.startsWith('branch')) {
        const v = findNextNonNullValue(cells, c)
        if (v !== null && !bank_details.branch_name) bank_details.branch_name = String(v).trim()
      } else if (label.startsWith('ifsc') || label.startsWith('ifs')) {
        const v = findNextNonNullValue(cells, c)
        if (v !== null && !bank_details.ifsc) bank_details.ifsc = String(v).replace(/^[-:]+\s*/, '').trim()
      } else if (label.startsWith('address')) {
        const v = findNextNonNullValue(cells, c)
        if (v !== null && !bank_details.address) bank_details.address = String(v).trim()
      } else if (label === 'city' || label.startsWith('city')) {
        const v = findNextNonNullValue(cells, c)
        if (v !== null && !bank_details.city) bank_details.city = String(v).trim()
      } else if (label.startsWith('opening balance')) {
        const v = findNextNonNullValue(cells, c)
        const num = extractNumber(v)
        if (num !== null) bank_details.opening_balance = num
      }
    }
  }

  for (let i = 0; i < rawData.length; i++) {
    const row = rawData[i]
    if (!row) continue
    const values = Object.values(row).map((v) => (v === null || v === undefined ? '' : String(v).trim().toLowerCase()))

    let matchCount = 0
    let hasDate = false
    let hasAmountCol = false
    let hasNarration = false
    for (const v of values) {
      if (!v) continue
      if (headerKeywords.some((k) => v === k || v.includes(k))) matchCount++
      if (v === 'date' || v.includes('date')) hasDate = true
      if (v.includes('withdrawal') || v.includes('deposit') || v.includes('debit') || v.includes('credit')) hasAmountCol = true
      if (v.includes('narration') || v.includes('description') || v.includes('particulars') || v.includes('details')) hasNarration = true
    }

    if (matchCount >= 3 && hasDate && hasAmountCol && hasNarration) {
      headerRowIndex = i
      for (const key of Object.keys(row)) {
        const raw = row[key]
        if (raw === null || raw === undefined) continue
        const v = String(raw).trim().toLowerCase()
        if (!v) continue

        if (!keyMapping.dateKey && v === 'date') {
          keyMapping.dateKey = key
        } else if (!keyMapping.valueDateKey && v.includes('value date')) {
          keyMapping.valueDateKey = key
        } else if (!keyMapping.dateKey && v.includes('date')) {
          keyMapping.dateKey = key
        } else if (!keyMapping.narrationKey && (v.includes('narration') || v.includes('description') || v.includes('particulars') || v.includes('details'))) {
          keyMapping.narrationKey = key
        } else if (!keyMapping.refKey && (v.includes('cheq') || v.includes('ref'))) {
          keyMapping.refKey = key
        } else if (!keyMapping.debitKey && (v.includes('withdrawal') || v.includes('debit') || v === 'dr')) {
          keyMapping.debitKey = key
        } else if (!keyMapping.creditKey && (v.includes('deposit') || v.includes('credit') || v === 'cr')) {
          keyMapping.creditKey = key
        } else if (!keyMapping.balanceKey && (v.includes('balance') || v === 'bal')) {
          keyMapping.balanceKey = key
        }
      }
      break
    }
  }

  if (headerRowIndex === -1) {
    return { bank_details, transactions }
  }

  for (let i = headerRowIndex + 1; i < rawData.length; i++) {
    const row = rawData[i]
    if (!row) continue

    const hasData = Object.values(row).some((v) => v !== null && v !== undefined && String(v).trim() !== '')
    if (!hasData) continue

    const narrationRaw = keyMapping.narrationKey ? row[keyMapping.narrationKey] : null
    const narration = narrationRaw !== null && narrationRaw !== undefined ? String(narrationRaw).trim() : null

    if (narration && narration.toLowerCase() === 'total') continue

    const dateRaw = keyMapping.dateKey ? row[keyMapping.dateKey] : null
    const date = parseAndFormatDate(dateRaw)
    if (!date) continue

    const debit = keyMapping.debitKey ? extractNumber(row[keyMapping.debitKey]) : null
    const credit = keyMapping.creditKey ? extractNumber(row[keyMapping.creditKey]) : null
    const balance = keyMapping.balanceKey ? extractNumber(row[keyMapping.balanceKey]) : null

    let type = null
    let amount = null
    if (debit !== null && debit !== 0) {
      type = 'withdrawal'
      amount = Math.abs(debit)
    } else if (credit !== null && credit !== 0) {
      type = 'deposit'
      amount = Math.abs(credit)
    } else {
      continue
    }

    transactions.push({
      date,
      voucher_number: voucher_number++,
      amount,
      desc: narration,
      from: null,
      to: null,
      type,
      balance: balance !== null ? balance : null,
    })
  }

  return { bank_details, transactions }
}

module.exports = processBankStatement
