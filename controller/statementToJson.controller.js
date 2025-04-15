const XLSX = require("xlsx");
const path = require("path");

const StatementToJsonController = {
  async convert(req, res) {
    try {
      if (!req.file) {
        return res.status(400).json({ flag: 0, message: "No file uploaded" });
      }

      const workbook = XLSX.readFile(req.file.path);
      const worksheet = workbook.Sheets[workbook.SheetNames[0]];
      const rawData = XLSX.utils.sheet_to_json(worksheet, {
        header: 1,
        defval: null,
        blankrows: true,
      });

      let headerRowIndex = rawData.findIndex((row) =>
        row?.some((cell) => {
          const cellText = String(cell || "").toLowerCase();
          return ["txn", "debit", "credit", "balance", "narration"].some(
            (keyword) => cellText.includes(keyword)
          );
        })
      );

      if (headerRowIndex === -1) {
        for (const rowIndex of [0, 1, 2, 3]) {
          if (
            getNormalizedHeaderMap(rawData[rowIndex] || []).date !== undefined
          ) {
            headerRowIndex = rowIndex;
            break;
          }
        }
        if (headerRowIndex === -1) {
          throw new Error(`Header row not found.`);
        }
      }

      const cleanedData = rawData.map((row) =>
        Array.isArray(row)
          ? row.map((cell) =>
              cell == null || cell === ""
                ? null
                : typeof cell === "string"
                ? cell.trim()
                : cell
            )
          : []
      );

      const result = parseHDFCExcel(cleanedData, headerRowIndex);
      return res.status(200).json({
        success: true,
        message: "File processed successfully",
        processed_data: {
          meta: result.meta,
          data: result.data,
        },
      });
    } catch (error) {
      console.error("Error:", error.message);
      return res.status(500).json({
        flag: 0,
        message: "Failed to process Excel file",
        error: error.message,
      });
    }
  },
};

// Helper functions (you already defined these, just include them outside of controller object)
const headerAliasMap = {
  date: [
    "date",
    "txn date",
    "transaction date",
    "value date",
    "posting date",
    "book date",
  ],
  narration: [
    "narration",
    "description",
    "transaction details",
    "particulars",
    "remarks",
    "details",
  ],
  ref_no: [
    "chq/ref number",
    "reference number",
    "ref no",
    "cheque no",
    "ref",
    "chq no",
  ],
  transaction_number: ["txn no.", "transaction no", "transaction number"],
  value_date: ["value date", "posting date", "effective date"],
  withdrawal: [
    "withdrawal",
    "dr amount",
    "debit",
    "withdrawal amount",
    "debit amount",
  ],
  deposit: [
    "deposit",
    "cr amount",
    "credit",
    "deposit amount",
    "credit amount",
  ],
  balance: [
    "balance",
    "closing balance",
    "available balance",
    "running balance",
    "current balance",
  ],
};

function getNormalizedHeaderMap(headers) {
  const normalized = {};
  headers.forEach((header, index) => {
    if (!header) return;
    const cleanHeader = String(header).toLowerCase().trim();

    if (cleanHeader.includes("date") || cleanHeader.includes("dt")) {
      if (!normalized.date) normalized.date = index;
    }

    for (const [standardKey, aliases] of Object.entries(headerAliasMap)) {
      if (
        aliases.some((alias) => {
          const cleanAlias = alias.toLowerCase();
          return cleanHeader === cleanAlias || cleanHeader.includes(cleanAlias);
        })
      ) {
        if (!normalized[standardKey]) {
          normalized[standardKey] = index;
          break;
        }
      }
    }
  });
  return normalized;
}

function formatDate(dateStr) {
  if (!dateStr) return null;
  try {
    const dateParts = String(dateStr).trim().split(/[/-]/);
    if (dateParts.length !== 3) return null;
    let [day, month, year] = dateParts.map((part) => part.padStart(2, "0"));
    if (year.length === 2) year = `20${year}`;
    const dateObj = new Date(`${year}-${month}-${day}`);
    if (isNaN(dateObj.getTime())) return null;
    return `${year}-${month}-${day}`;
  } catch {
    return null;
  }
}

function parseMoney(value) {
  if (value === null || value === undefined || value === "") return 0;
  const strValue = String(value)
    .replace(/,/g, "")
    .replace(/[^\d.-]/g, "");
  const num = parseFloat(strValue);
  return isNaN(num) ? 0 : Math.abs(num);
}

function isValidTransaction(tx) {
  return (
    tx.date && tx.desc && tx.desc.trim() && tx.amount > 0 && !isNaN(tx.balance)
  );
}

function parseHDFCExcel(sheetData, headerRowIndex) {
  const meta = {
    account_number: "Not Available",
    account_holder_name: "Not Available",
    bank_name: "HDFC BANK Ltd.",
    ifsc: "Not Available",
    branch_name: "Not Available",
    branch_city: "Not Available",
    branch_state: "Not Available",
    statement_from: "",
    statement_to: "",
  };

  for (let i = 0; i < headerRowIndex; i++) {
    const row = sheetData[i] || [];
    for (const cell of row) {
      if (!cell) continue;
      const cellText = String(cell).trim();

      if (/^(MR|MS|MRS)/i.test(cellText)) {
        meta.account_holder_name = cellText.replace(/^(MR|MS|MRS)/i, "").trim();
      } else if (cellText.includes("Account No :")) {
        meta.account_number = cellText
          .split("Account No :")[1]
          .split(/\s+/)[0]
          .trim();
      } else if (cellText.includes("IFSC :")) {
        meta.ifsc = cellText.split("IFSC :")[1].split(/\s+/)[0].trim();
      } else if (cellText.includes("Account Branch :")) {
        meta.branch_name = cellText.split("Account Branch :")[1].trim();
      } else if (cellText.startsWith("Statement From")) {
        const dates = cellText.match(/(\d{2}\/\d{2}\/\d{4})/g);
        if (dates?.length === 2) {
          meta.statement_from = formatDate(dates[0]) || "";
          meta.statement_to = formatDate(dates[1]) || "";
        }
      } else if (cellText.startsWith("City :")) {
        meta.branch_city = cellText.split("City :")[1].trim();
      } else if (cellText.startsWith("State :")) {
        meta.branch_state = cellText.split("State :")[1].trim();
      }
    }
  }

  const headers = sheetData[headerRowIndex] || [];
  const headerMap = getNormalizedHeaderMap(headers);
  if (!headerMap.date) throw new Error("Date column not found");

  const transactions = [];
  for (let i = headerRowIndex + 1; i < sheetData.length; i++) {
    const row = sheetData[i] || [];
    if (row.length === 0) continue;

    const formattedDate = formatDate(row[headerMap.date]);
    if (!formattedDate) continue;

    const withdrawal = parseMoney(row[headerMap.withdrawal]);
    const deposit = parseMoney(row[headerMap.deposit]);
    const amount = withdrawal > 0 ? withdrawal : deposit;
    if (amount <= 0) continue;

    const transaction = {
      date: formattedDate,
      value_date: formatDate(row[headerMap.value_date]) || formattedDate,
      transaction_no:
        String(row[headerMap.transaction_number] || "").trim() || "N/A",
      desc: String(row[headerMap.narration] || "").trim(),
      ref_no: String(row[headerMap.ref_no] || "").trim() || "N/A",
      amount: amount,
      type: withdrawal > 0 ? "withdrawal" : "deposit",
      balance: parseMoney(row[headerMap.balance]),
    };

    if (isValidTransaction(transaction)) {
      transactions.push(transaction);
    }
  }

  return { meta, data: transactions };
}

module.exports = StatementToJsonController;
