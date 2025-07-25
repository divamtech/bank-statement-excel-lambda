/**
 * Generated by Gemini AI - Attempt 1
 * Processes bank statement data from Excel
 * @param {Array} rawData - Array of objects from bank statement Excel
 * @returns {Object} Processed bank statement data
 */
function processBankStatement(rawData) {
  // Use the new extraction function for array-of-arrays input
  const bank_details = extractBankDetailsFromRows(rawData)
  const transactions = []
  let headerRowIndex = -1
  const maxBankDetailsRows = Math.min(20, rawData.length)
  let keyMap = {}

  // Helper function to parse dates
  function parseDate(dateString) {
    if (!dateString) return null

    // Attempt to parse as Excel date serial number
    if (typeof dateString === 'number') {
      const excelEpoch = new Date(Date.UTC(1899, 11, 31))
      const javascriptDate = new Date(excelEpoch.getTime() + (dateString - 1) * 24 * 60 * 60 * 1000)
      const year = javascriptDate.getFullYear()
      const month = String(javascriptDate.getMonth() + 1).padStart(2, '0')
      const day = String(javascriptDate.getDate()).padStart(2, '0')
      return `${year}-${month}-${day}`
    }

    // Attempt to parse various date formats
    const dateFormats = [
      { format: /(\d{2})\/(\d{2})\/(\d{4})/, order: [2, 1, 3] }, // DD/MM/YYYY
      { format: /(\d{2})-(\w+)-(\d{4})/, order: [1, 2, 3] }, // DD-Mon-YYYY
      { format: /(\d{2})-(\d{2})-(\d{4})/, order: [2, 1, 3] }, // DD-MM-YYYY
      { format: /(\d{2})\/(\d{2})\/(\d{2,4})/, order: [2, 1, 3] }, // DD/MM/YY or DD/MM/YYYY
      { format: /(\d{4})-(\d{2})-(\d{2})/, order: [3, 2, 1] }, //YYYY-MM-DD
      { format: /(\d{2})-(\w{3})-(\d{2})/, order: [1, 2, 3] }, //DD-MON-YY
    ]

    for (const dateFormat of dateFormats) {
      const match = dateString?.match(dateFormat.format)
      if (match) {
        let year = match[dateFormat.order[2]]
        if (year.length === 2) {
          year = `20${year}` // Assuming 21st century for 2-digit years
        }
        const month = match[dateFormat.order[1]]
        const day = match[dateFormat.order[0]]

        //Convert month abbreviations to number
        const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
        let monthNumber = month
        if (isNaN(month)) {
          const monthIndex = months.findIndex((m) => m.toLowerCase() == month.toLowerCase())
          monthNumber = String(monthIndex + 1).padStart(2, '0')
        }

        return `${year}-${String(monthNumber).padStart(2, '0')}-${String(day).padStart(2, '0')}`
      }
    }
    return null
  }

  // Helper function to clean amount
  function cleanAmount(amountString) {
    if (!amountString) return null
    const cleanedString = String(amountString).replace(/[^0-9.-]+/g, '')
    const parsedAmount = parseFloat(cleanedString)
    return isNaN(parsedAmount) ? null : parsedAmount
  }

  // 1. Extract Bank Details
  for (let i = 0; i < maxBankDetailsRows; i++) {
    const row = rawData[i]
    if (!row) continue

    for (const key in row) {
      if (row.hasOwnProperty(key) && row[key]) {
        const value = String(row[key]).trim()

        if (value.toLowerCase().includes('bank name')) {
          const bankNameKey = Object.keys(row).find((k) => k !== key)
          if (bankNameKey && row[bankNameKey]) {
            bank_details.bank_name = String(row[bankNameKey]).trim() || bank_details.bank_name
          }
        } else if (value.toLowerCase().includes('account holder name')) {
          const accountHolderNameKey = Object.keys(row).find((k) => k !== key)
          if (accountHolderNameKey && row[accountHolderNameKey]) {
            bank_details.account_holder_name = String(row[accountHolderNameKey]).trim() || bank_details.account_holder_name
          }
        } else if (value.toLowerCase().includes('account no')) {
          const accountNoKey = Object.keys(row).find((k) => k !== key)
          if (accountNoKey && row[accountNoKey]) {
            bank_details.account_no = String(row[accountNoKey]).trim() || bank_details.account_no
          }
        } else if (value.toLowerCase().includes('ifsc') || value.toLowerCase().includes('ifs')) {
          const ifscKey = Object.keys(row).find((k) => k !== key)
          if (ifscKey && row[ifscKey]) {
            bank_details.ifsc = String(row[ifscKey]).trim() || bank_details.ifsc
          }
        } else if (value.toLowerCase().includes('branch name')) {
          const branchNameKey = Object.keys(row).find((k) => k !== key)
          if (branchNameKey && row[branchNameKey]) {
            bank_details.branch_name = String(row[branchNameKey]).trim() || bank_details.branch_name
          }
        } else if (value.toLowerCase().includes('branch code')) {
          const branchCodeKey = Object.keys(row).find((k) => k !== key)
          if (branchCodeKey && row[branchCodeKey]) {
            bank_details.branch_code = String(row[branchCodeKey]).trim() || bank_details.branch_code
          }
        } else if (value.toLowerCase().includes('address')) {
          const addressKey = Object.keys(row).find((k) => k !== key)
          if (addressKey && row[addressKey]) {
            bank_details.address = String(row[addressKey]).trim() || bank_details.address
          }
        }
      }
    }
  }

  if (!bank_details.ifsc) {
    for (let i = 0; i < maxBankDetailsRows; i++) {
      const row = rawData[i]
      if (!row) continue

      for (const key in row) {
        if (row.hasOwnProperty(key) && row[key]) {
          const value = String(row[key]).trim()
          if (key.toLowerCase().includes('ifsc') || key.toLowerCase().includes('ifs')) {
            bank_details.ifsc = value || bank_details.ifsc
          }
        }
      }
    }
  }

  // 2. Identify Header Row and Create Key Mapping
  for (let i = 0; i < rawData.length; i++) {
    const row = rawData[i]
    if (!row) continue

    let isHeaderRow = false
    for (const key in row) {
      if (row.hasOwnProperty(key) && row[key]) {
        const value = String(row[key]).trim().toLowerCase()
        if (
          value.includes('date') ||
          value.includes('narration') ||
          value.includes('description') ||
          value.includes('details') ||
          value.includes('withdrawal') ||
          value.includes('debit') ||
          value.includes('deposit') ||
          value.includes('credit') ||
          value.includes('balance') ||
          value.includes('closing balance')
        ) {
          isHeaderRow = true
          break
        }
      }
    }

    if (isHeaderRow) {
      headerRowIndex = i
      for (const key in row) {
        if (row.hasOwnProperty(key) && row[key]) {
          const value = String(row[key]).trim().toLowerCase()
          if (value.includes('date')) {
            keyMap.dateKey = key
          } else if (value.includes('narration') || value.includes('description') || value.includes('details')) {
            keyMap.narrationKey = key
          } else if (value.includes('withdrawal') || value.includes('debit')) {
            keyMap.debitKey = key
          } else if (value.includes('deposit') || value.includes('credit')) {
            keyMap.creditKey = key
          } else if (value.includes('balance')) {
            keyMap.balanceKey = key
          }
        }
      }
      break
    }
  }

  // 3. Process Transactions
  let voucherNumber = 1
  for (let i = headerRowIndex + 1; i < rawData.length; i++) {
    const row = rawData[i]
    if (!row) continue

    // Check if date is null, then set row index to 1 for setting date
    const dateString = row[keyMap.dateKey] ? row[keyMap.dateKey] : row['1']
    const desc = row[keyMap.narrationKey] ? String(row[keyMap.narrationKey]).trim() : null
    const debitString = row[keyMap.debitKey]
    const creditString = row[keyMap.creditKey]
    const balanceString = row[keyMap.balanceKey]

    const date = dateString
    const debit = cleanAmount(debitString)
    const credit = cleanAmount(creditString)
    const balance = cleanAmount(balanceString)

    if (!date && !desc && !debit && !credit && !balance) {
      continue // Skip empty rows
    }

    let type = null
    let amount = null

    if (debit !== null) {
      type = 'withdrawal'
      amount = Math.abs(debit) // Ensure amount is positive
    } else if (credit !== null) {
      type = 'deposit'
      amount = Math.abs(credit) // Ensure amount is positive
    }

    if (type && amount !== null) {
      const transaction = {
        date: date,
        voucher_number: voucherNumber++,
        amount: amount,
        desc: desc,
        from: null,
        to: null,
        type: type,
        balance: balance !== null ? balance : null,
      }

      transactions.push(transaction)
    }
  }

  // 4. Return Result
  return {
    bank_details: bank_details,
    transactions: transactions,
  }
}

// Helper function to extract bank details from array-of-arrays
function extractBankDetailsFromRows(rows) {
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

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i]
    for (let j = 0; j < row.length; j++) {
      const cell = row[j]
      if (!cell) continue
      const value = String(cell).trim()

      if (value.toLowerCase().includes('account holder name')) {
        const afterColon = value.split(':')[1]
        if (afterColon && afterColon.trim()) {
          bank_details.account_holder_name = afterColon.trim()
        } else {
          for (let k = j + 1; k < row.length; k++) {
            if (row[k]) {
              bank_details.account_holder_name = String(row[k]).trim()
              break
            }
          }
        }
      } else if (value.toLowerCase().includes('branch name')) {
        const afterColon = value.split(':')[1]
        if (afterColon && afterColon.trim()) {
          bank_details.branch_name = afterColon.trim()
        } else {
          for (let k = j + 1; k < row.length; k++) {
            if (row[k]) {
              bank_details.branch_name = String(row[k]).trim()
              break
            }
          }
        }
      } else if (value.toLowerCase().includes('ifsc')) {
        const afterColon = value.split(':')[1]
        if (afterColon && afterColon.trim()) {
          bank_details.ifsc = afterColon.trim()
        } else {
          for (let k = j + 1; k < row.length; k++) {
            if (row[k]) {
              bank_details.ifsc = String(row[k]).trim()
              break
            }
          }
        }
      } else if (value.toLowerCase().includes('address')) {
        const afterColon = value.split(':')[1]
        if (afterColon && afterColon.trim()) {
          bank_details.address = afterColon.trim()
        } else {
          for (let k = j + 1; k < row.length; k++) {
            if (row[k]) {
              bank_details.address = String(row[k]).trim()
              break
            }
          }
        }
      }
      // Add more fields as needed...
    }
  }
  return bank_details
}

module.exports = processBankStatement
