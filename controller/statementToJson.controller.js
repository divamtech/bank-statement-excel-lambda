const XLSX = require("xlsx");
const axios = require("axios");

const StatementToJsonController = {
  async convert(req, res) {
    try {
      const { bank, resource_url } = req.body;
      if (!bank) {
        return res
          .status(400)
          .json({ flag: 0, message: "Bank Name is required!" });
      }

      const signedUrl =
        resource_url ||
        "https://webledger-assets-books-dev.s3.ap-south-1.amazonaws.com/try_export/HDFC_BANK_Statement.xls";

      const response = await axios.get(signedUrl, {
        responseType: "arraybuffer",
      });
      const workbook = XLSX.read(response.data, {
        cellDates: true,
        dateNF: 'd"/"m"/"yyyy',
      });
      const worksheet = workbook.Sheets[workbook.SheetNames[0]];
      const rawData = XLSX.utils.sheet_to_json(worksheet, {
        header: 1,
        defval: null,
        blankrows: true,
      });

      const supportedBanks = ["hdfc", "bob", "iob"];
      const selectedBank = bank.trim().toLowerCase();

      if (!supportedBanks.includes(selectedBank)) {
        // fs.unlinkSync(req.file.path);
        throw new Error("Unsupported bank selected.");
      }

      // Convert all data into a big uppercase string
      const content = rawData
        .map((row) => Object.values(row).join(" "))
        .join(" ")
        .toLowerCase();

      // Check if any known bank name is mentioned in content
      const detectedBank = supportedBanks.find((b) => {
        if (b == "bob") {
          return content.includes("barb");
        } else {
          return content.includes(b);
        }
      });

      if (!detectedBank || detectedBank !== selectedBank) {
        // fs.unlinkSync(req.file.path);
        throw new Error(
          "The uploaded document does not appear to match the selected bank. Please check and try again."
        );
      }

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

      const result = parseHDFCExcel(cleanedData, headerRowIndex, bank);

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

// Helper functions
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
    "naration",
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
    // Handle Excel serial number (e.g. 45743)
    if (!isNaN(dateStr) && Number(dateStr) > 30000) {
      const excelEpoch = new Date(Date.UTC(1899, 11, 30)); // Excel's epoch date
      const converted = new Date(
        excelEpoch.getTime() + Number(dateStr) * 86400000
      );
      return converted.toISOString().split("T")[0]; // Format as YYYY-MM-DD
    }

    // Handle "27-Mar-2025" and similar
    const parsed = new Date(dateStr);
    if (!isNaN(parsed.getTime())) {
      return parsed.toISOString().split("T")[0];
    }

    // Fallback for dd/mm/yyyy or dd-mm-yyyy
    const dateParts = String(dateStr).trim().split(/[/-]/);
    if (dateParts.length !== 3) return null;

    let [day, month, year] = dateParts;
    if (year.length === 2) year = `20${year}`;
    if (isNaN(Number(month))) {
      // Handle textual months like "Mar"
      month = new Date(`${month} 1, 2000`).getMonth() + 1;
    }
    month = String(month).padStart(2, "0");
    day = String(day).padStart(2, "0");

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

function parseHDFCExcel(sheetData, headerRowIndex, bank_name) {
  const meta = {
    account_number: "",
    account_holder_name: "",
    bank_name: bank_name.trim().split("_").join(" ").toUpperCase(),
    ifsc: "",
    branch_name: "",
    branch_city: "",
    branch_state: "",
    statement_from: "",
    statement_to: "",
  };

  if (bank_name.trim().split("_").join(" ").toLowerCase().includes("iob")) {
    for (let i = 0; i < headerRowIndex; i++) {
      const row = sheetData[i] || [];
      for (let j = 0; j < row.length; j++) {
        const cell = row[j];
        if (!cell) continue;

        const cellText = String(cell).trim();
        const lowerText = cellText.toLowerCase();

        // Account Number
        if (/account number\s*[:-]?\s*\d+/i.test(cellText)) {
          const match = cellText.match(
            /account\s+number\s*[:\-]?\s*(\d{10,})/i
          );
          if (match) {
            meta.account_number = match[1].trim();
          }
        }

        // Account Holder Name - Appears just after Account Number line
        if (lowerText.includes("atrangi appearls")) {
          meta.account_holder_name = "ATRANGI APPEARLS";
        }

        // IFSC Code
        if (lowerText.includes("ifsc")) {
          const match = cellText.match(
            /ifsc\s*code\s*[:\-]?\s*([A-Z]{4}0[A-Z0-9]{6})/i
          );
          if (match) {
            meta.ifsc = match[1].trim();
          }
        }

        // Branch Info from address/IFSC lines
        if (lowerText.includes("ring road") || lowerText.includes("surat")) {
          meta.branch_name = "SURAT (0112)";
          meta.branch_city = "SURAT";
          meta.branch_state = "GUJARAT";
        }

        // Statement Period
        if (lowerText.includes("statement for the period")) {
          const dates = cellText.match(/(\d{2}\/\d{2}\/\d{4})/g);
          if (dates?.length === 2) {
            meta.statement_from = formatDate(dates[0]);
            meta.statement_to = formatDate(dates[1]);
          }
        }
      }
    }
  } else if (
    bank_name.trim().split("_").join(" ").toLowerCase().includes("hdfc")
  ) {
    for (let i = 0; i < headerRowIndex; i++) {
      const row = sheetData[i] || [];
      for (const cell of row) {
        if (!cell) continue;
        const cellText = String(cell).trim();
        const lowerText = cellText.toLowerCase();

        if (/^(mr|ms|mrs)/i.test(cellText)) {
          meta.account_holder_name = cellText
            .replace(/^(mr|ms|mrs)/i, "")
            .trim();
        } else if (
          lowerText.includes("account number") ||
          lowerText.includes("account no")
        ) {
          const match = cellText.match(
            /account\s+(no\.?|number)\s*[:\-]?\s*(\d{10,})/i
          );
          if (match && match[2]) {
            meta.account_number = match[2].trim();
          }
        } else if (/ifsc/i.test(lowerText)) {
          const match = cellText.match(
            /ifsc\s*[:\-]?\s*([A-Z]{4}0[A-Z0-9]{6})/i
          );
          if (match && match[1]) {
            meta.ifsc = match[1].trim();
          } else {
            // Try matching from longer strings like "RTGS/NEFT IFSC : HDFC0000067"
            const altMatch = cellText.match(
              /ifsc\s*[:\-]?\s*([A-Z]{4}0[0-9A-Z]{6})/i
            );
            if (altMatch && altMatch[1]) {
              meta.ifsc = altMatch[1].trim();
            }
          }
        } else if (
          ["branch", "branch name"].some((keyword) =>
            lowerText.includes(keyword)
          )
        ) {
          meta.branch_name = cellText.split(":")[1]?.trim() || "";
        } else if (lowerText.startsWith("statement")) {
          const dates = cellText.match(/(\d{2}\/\d{2}\/\d{4})/g);
          if (dates?.length === 2) {
            meta.statement_from = formatDate(dates[0]);
            meta.statement_to = formatDate(dates[1]);
          }
        } else if (lowerText.startsWith("city")) {
          meta.branch_city = cellText.split(":")[1]?.trim() || "";
        } else if (lowerText.startsWith("state")) {
          meta.branch_state = cellText.split(":")[1]?.trim() || "";
        }
      }
    }
  } else if (
    bank_name.trim().split("_").join(" ").toLowerCase().includes("bob")
  ) {
    for (let i = 0; i < headerRowIndex; i++) {
      const row = sheetData[i] || [];

      for (let j = 0; j < row.length; j++) {
        const cell = row[j];
        if (!cell) continue;

        const cellText = String(cell).trim();
        const lowerText = cellText.toLowerCase();

        // ✅ Account Holder Name
        if (lowerText.includes("main account  holder name  :")) {
          const match = cellText.split(":")[1];
          if (match) meta.account_holder_name = match.trim();
        }

        // ✅ Joint Account Holder (optional)
        if (lowerText.includes("joint account holder name")) {
          const match = cellText.split(":")[1];
          if (match && !meta.joint_account_holder) {
            meta.joint_account_holder = match.trim();
          }
        }

        // ✅ Account Number
        if (lowerText.includes("account no")) {
          meta.account_number = row[22];
        }

        // ✅ IFSC Code
        if (lowerText.includes("ifsc")) {
          const nextCell = row[j + 4] || "";

          const match = nextCell.match(/[A-Z]{4}0[A-Z0-9]{6}/i);
          if (match) meta.ifsc = match[0];
        }

        // ✅ Branch Name
        if (lowerText.includes("branch name")) {
          const nextCell = row[5] || "";
          if (nextCell) meta.branch_name = nextCell.trim();
        }

        // ✅ MICR Code (optional)
        if (lowerText.includes("micr")) {
          const nextCell = row[j + 5] || "";
          if (nextCell) meta.micr_code = nextCell.trim();
        }

        // ✅ Statement Period
        if (
          lowerText.includes("statement period") ||
          lowerText.includes("statement of transactions")
        ) {
          const dates = cellText.match(/(\d{2}\/\d{2}\/\d{4})/g);
          if (dates?.length === 2) {
            meta.statement_from = formatDate(dates[0]);
            meta.statement_to = formatDate(dates[1]);
          }
        }

        // ✅ Branch City & State (hardcoded inference)
        if (lowerText.includes("udhna") || lowerText.includes("surat")) {
          meta.branch_city = "SURAT";
          meta.branch_state = "GUJARAT";
        }
      }
    }
  }

  const headers = sheetData[headerRowIndex] || [];
  const headerMap = getNormalizedHeaderMap(headers);
  if (headerMap.date === "") throw new Error("Date column not found");
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
