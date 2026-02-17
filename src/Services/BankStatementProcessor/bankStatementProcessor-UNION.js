/**
 * Union Bank of India Statement Processor
 * Processes bank statement data from Excel
 * @param {Array} rawData - Array of objects from bank statement Excel
 * @returns {Object} Processed bank statement data
 */
function processBankStatement(rawData) {
  const bank_details = {
    bank_name: 'Union Bank of India',
    opening_balance: 0,
    ifsc: null,
    address: null,
    city: null,
    account_no: null,
    account_holder_name: null,
    branch_name: null,
    branch_code: null,
  }
  const transactions = []
  let headerRowIndex = -1
  let keyMap = {}
  let voucher_number = 1

  // Helper function to parse various date formats including Union Bank's DD-MM-YYYY HH:MM:SS format
  function parseDate(dateString) {
    if (!dateString) return null

    // If it's already a Date object, format it directly
    if (dateString instanceof Date && !isNaN(dateString)) {
      const year = dateString.getFullYear()
      const month = String(dateString.getMonth() + 1).padStart(2, '0')
      const day = String(dateString.getDate()).padStart(2, '0')
      return `${year}-${month}-${day}`
    }

    // Convert to string for pattern matching
    const dateStr = String(dateString).trim()

    // Union Bank specific format: DD-MM-YYYY HH:MM:SS
    const unionBankFormat = /^(\d{2})-(\d{2})-(\d{4})\s+\d{2}:\d{2}:\d{2}$/
    if (unionBankFormat.test(dateStr)) {
      const match = dateStr.match(/^(\d{2})-(\d{2})-(\d{4})/)
      if (match) {
        const [, day, month, year] = match
        return `${year}-${month}-${day}`
      }
    }

    const formats = [
      /^\d{2}\/\d{2}\/\d{4}$/, // DD/MM/YYYY
      /^\d{2}-\d{2}-\d{4}$/, // DD-MM-YYYY
      /^\d{2}-\w{3}-\d{4}$/, // DD-Mon-YYYY
      /^\d{2}\/\d{2}\/\d{2}$/, // DD/MM/YY
      /^\d{4}-\d{2}-\d{2}$/, // YYYY-MM-DD
    ]

    const parsers = [
      (str) => {
        // DD/MM/YYYY
        const [day, month, year] = str.split('/').map(Number)
        return new Date(year, month - 1, day)
      },
      (str) => {
        // DD-MM-YYYY
        const [day, month, year] = str.split('-').map(Number)
        return new Date(year, month - 1, day)
      },
      (str) => {
        // DD-Mon-YYYY
        const [day, monthStr, year] = str.split('-')
        const month = new Date(Date.parse(monthStr + ' 1, 2000')).getMonth()
        return new Date(year, month, day)
      },
      (str) => {
        // DD/MM/YY
        const [day, month, year] = str.split('/').map(Number)
        const fullYear = year < 70 ? 2000 + year : 1900 + year
        return new Date(fullYear, month - 1, day)
      },
      (str) => {
        // YYYY-MM-DD
        const [year, month, day] = str.split('-').map(Number)
        return new Date(year, month - 1, day)
      },
    ]

    for (let i = 0; i < formats.length; i++) {
      if (formats[i].test(dateStr)) {
        try {
          const date = parsers[i](dateStr)
          if (!isNaN(date)) {
            const year = date.getFullYear()
            const month = String(date.getMonth() + 1).padStart(2, '0')
            const day = String(date.getDate()).padStart(2, '0')
            return `${year}-${month}-${day}`
          }
        } catch (e) {
          // Ignore parsing errors, try next format
        }
      }
    }

    // Try Excel date serial number
    const excelSerial = parseInt(dateStr, 10)
    if (!isNaN(excelSerial) && excelSerial > 0 && excelSerial < 100000) {
      try {
        const excelEpoch = new Date(Date.UTC(1899, 11, 31))
        const date = new Date(excelEpoch.getTime() + excelSerial * 24 * 60 * 60 * 1000)
        const year = date.getFullYear()
        const month = String(date.getMonth() + 1).padStart(2, '0')
        const day = String(date.getDate()).padStart(2, '0')
        return `${year}-${month}-${day}`
      } catch (e) {
        // Ignore Excel Serial Date Parsing
      }
    }

    return null // No valid format found
  }

  function extractFromDescription(description) {
    if (!description) return { from: null, to: null }

    const upiPattern = /[a-zA-Z0-9.\-_]{2,}@[a-zA-Z0-9.\-_]{2,}/
    const match = description.match(upiPattern)
    if (match) {
      return { to: match[0], from: null }
    }

    return { from: null, to: null }
  }

  // 1. Bank Details Extraction from Union Bank format
  for (let i = 0; i < Math.min(rawData.length, 25); i++) {
    const row = rawData[i]
    for (const key in row) {
      if (row.hasOwnProperty(key)) {
        const value = row[key]
        if (value) {
          const nextKey = Object.keys(row)[Object.keys(row).indexOf(key) + 1]
          const nextValue = row[nextKey]

          if (typeof value === 'string') {
            const lowerValue = value.toLowerCase()

            // Helper to extract value after colon or from next cell
            function extractValue(str, next) {
              const idx = str.indexOf(':')
              if (idx !== -1) {
                return str.slice(idx + 1).trim()
              }
              return next || null
            }

            // Look for account holder name (Union Bank puts it near the top)
            if (
              !lowerValue.includes('bank') &&
              !lowerValue.includes('branch') &&
              !lowerValue.includes('address') &&
              !lowerValue.includes('account') &&
              !lowerValue.includes('customer') &&
              !lowerValue.includes('ckyc') &&
              !lowerValue.includes('city') &&
              !lowerValue.includes('state') &&
              !lowerValue.includes('country') &&
              !lowerValue.includes('zip') &&
              !lowerValue.includes('mobile') &&
              !lowerValue.includes('email') &&
              !lowerValue.includes('statement') &&
              !lowerValue.includes('records') &&
              value.length > 5 &&
              value.length < 100 &&
              /^[A-Z\s]+$/.test(value) &&
              !bank_details.account_holder_name &&
              i < 5
            ) {
              bank_details.account_holder_name = value.trim()
            }

            if (lowerValue.includes('branch') && !lowerValue.includes('code')) {
              bank_details.branch_name = extractValue(value, nextValue)
            } else if (lowerValue.includes('customer id')) {
              // Customer ID row, skip it
            } else if (lowerValue.includes('a/c number') || lowerValue.includes('account number')) {
              bank_details.account_no = extractValue(value, nextValue)
            } else if (lowerValue.includes('city') && !lowerValue.includes('account')) {
              bank_details.city = extractValue(value, nextValue)
            } else if (lowerValue.includes('address') && !lowerValue.includes('email')) {
              bank_details.address = extractValue(value, nextValue)
            } else if (lowerValue.includes('ifsc')) {
              const ifscValue = extractValue(value, nextValue)
              if (ifscValue) {
                const ifscMatch = ifscValue.match(/[A-Z]{4}0[A-Z0-9]{6}/i)
                if (ifscMatch) {
                  bank_details.ifsc = ifscMatch[0].toUpperCase()
                }
              }
            } else if (lowerValue.includes('opening balance')) {
              let balValue = null
              const extractedVal = extractValue(value, nextValue)
              if (extractedVal) {
                balValue = parseFloat(extractedVal.replace(/[^0-9.-]+/g, ''))
                if (!isNaN(balValue)) {
                  bank_details.opening_balance = balValue
                }
              }
            }
          }
        }
      }
    }
  }

  // Universal fallback for IFSC extraction
  if (!bank_details.ifsc) {
    for (let i = 0; i < Math.min(rawData.length, 25); i++) {
      const row = rawData[i]
      for (const key in row) {
        if (row.hasOwnProperty(key)) {
          const value = row[key]
          if (typeof value === 'string') {
            const ifscMatch = value.match(/[A-Z]{4}0[A-Z0-9]{6}/i)
            if (ifscMatch) {
              bank_details.ifsc = ifscMatch[0].toUpperCase()
              break
            }
          }
        }
      }
      if (bank_details.ifsc) break
    }
  }

  // 2. Header Row Identification & Key Mapping for Union Bank
  // Union Bank headers: Date, Tran Id, Remarks, UTR Number, Instr. Id, Withdrawals, Deposits, Balance
  for (let i = 0; i < rawData.length; i++) {
    const row = rawData[i]
    let dateFound = false
    let remarksFound = false
    let withdrawalsFound = false
    let depositsFound = false
    let balanceFound = false

    for (const key in row) {
      if (row.hasOwnProperty(key)) {
        const value = row[key]
        if (typeof value === 'string') {
          const lowerValue = value.toLowerCase().trim()

          if (lowerValue === 'date' || lowerValue === 'tran date') {
            keyMap.dateKey = key
            dateFound = true
          } else if (lowerValue === 'tran id' || lowerValue.includes('transaction id')) {
            keyMap.tranIdKey = key
          } else if (lowerValue === 'remarks' || lowerValue.includes('narration') || lowerValue.includes('description')) {
            keyMap.narrationKey = key
            remarksFound = true
          } else if (lowerValue === 'utr number' || lowerValue.includes('utr')) {
            keyMap.utrKey = key
          } else if (lowerValue.includes('instr') && lowerValue.includes('id')) {
            keyMap.instrKey = key
          } else if (lowerValue === 'withdrawals' || lowerValue.includes('withdrawal')) {
            keyMap.debitKey = key
            withdrawalsFound = true
          } else if (lowerValue === 'deposits' || lowerValue.includes('deposit')) {
            keyMap.creditKey = key
            depositsFound = true
          } else if (lowerValue === 'balance' || lowerValue.includes('closing balance')) {
            keyMap.balanceKey = key
            balanceFound = true
          }
        }
      }
    }

    if (dateFound && remarksFound && (withdrawalsFound || depositsFound) && balanceFound) {
      headerRowIndex = i
      break
    }
  }

  // 3. Transaction Processing
  if (headerRowIndex !== -1) {
    for (let i = headerRowIndex + 1; i < rawData.length; i++) {
      const row = rawData[i]
      if (!row || Object.keys(row).length === 0) {
        continue // Skip empty rows
      }

      const dateString = row[keyMap.dateKey]
      const desc = row[keyMap.narrationKey] ? String(row[keyMap.narrationKey]).trim() : null
      const debitValue = row[keyMap.debitKey]
      const creditValue = row[keyMap.creditKey]
      const balanceValue = keyMap.balanceKey ? row[keyMap.balanceKey] : null

      const date = parseDate(dateString)

      // Skip if no valid date
      if (!date) {
        continue
      }

      // Skip dates with year less than 2000 (invalid/incorrectly parsed dates)
      const year = parseInt(date.split('-')[0])
      if (year < 2000) {
        continue
      }

      // Skip summary rows, footer rows, etc.
      if (
        desc &&
        (desc.includes('***') ||
          desc.includes('End Of Statement') ||
          desc.includes('STATEMENT SUMMARY') ||
          desc.includes('Generated') ||
          desc.includes('No more records') ||
          desc.toLowerCase().includes('records from'))
      ) {
        continue
      }

      // Parse amounts
      let amount = null
      let type = null

      if (debitValue !== null && debitValue !== undefined && debitValue !== '') {
        const parsedDebit = typeof debitValue === 'number' ? debitValue : parseFloat(String(debitValue).replace(/[^0-9.-]+/g, ''))
        if (!isNaN(parsedDebit) && parsedDebit > 0) {
          amount = Math.abs(parsedDebit)
          type = 'withdrawal'
        }
      }

      if (creditValue !== null && creditValue !== undefined && creditValue !== '') {
        const parsedCredit = typeof creditValue === 'number' ? creditValue : parseFloat(String(creditValue).replace(/[^0-9.-]+/g, ''))
        if (!isNaN(parsedCredit) && parsedCredit > 0) {
          amount = Math.abs(parsedCredit)
          type = 'deposit'
        }
      }

      // Skip if no amount
      if (amount === null || type === null) {
        continue
      }

      // Balance is optional
      let balValue = null
      if (balanceValue !== null && balanceValue !== undefined) {
        balValue = typeof balanceValue === 'number' ? balanceValue : parseFloat(String(balanceValue).replace(/[^0-9.-]+/g, ''))
        if (isNaN(balValue)) balValue = null
      }

      const { from, to } = extractFromDescription(desc)

      transactions.push({
        date: date,
        voucher_number: voucher_number++,
        amount: amount,
        desc: desc,
        from: from,
        to: to,
        type: type,
        balance: balValue,
      })
    }
  }

  return { bank_details: bank_details, transactions: transactions }
}

module.exports = processBankStatement
