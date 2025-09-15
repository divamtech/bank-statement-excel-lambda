const XLSX = require('xlsx')
const axios = require('axios')
const BankENUM = require('../BANK')
const { SQSClient, SendMessageCommand } = require('@aws-sdk/client-sqs')
const fs = require('fs')
const path = require('path')

const sqsClient = new SQSClient({
  region: process.env.AWS_REGION,
  credentials: {
    accessKeyId: process.env.AWS_ACCESS_KEY_ID,
    secretAccessKey: process.env.AWS_SECRET_ACCESS_KEY,
  },
})

async function pushSQS(bank, resource_url) {
  const params = {
    QueueUrl: process.env.SQS_URL,
    MessageBody: JSON.stringify({
      bank: bank,
      resource_url: resource_url,
      timestamp: Date.now(),
    }),
  }

  return await sqsClient.send(new SendMessageCommand(params))
}

const StatementToJsonController = {
  async convert(req, res) {
    try {
      const { bank, resource_url } = req.body

      if (!bank && !resource_url) {
        return res.status(400).json({ flag: 0, message: 'Bank Name and Resource URL are required!' })
      }

      // Check if bank exists in BankENUM (case-insensitive)
      const bankKey = Object.keys(BankENUM).find((key) => BankENUM[key] === bank.toLowerCase())
      if (!bankKey) {
        await pushSQS(bank, resource_url)
        console.log('bankKeyE::', bankKey)
        return res.status(400).json({ flag: 0, message: 'BANK_NOT_FOUND' })
      }
      if (!bank && !resource_url) {
        return res.status(400).json({ flag: 0, message: 'Bank Name and Resource URL is required!' })
      }
      const validatedBankName = BankENUM[bankKey] // Use the value from ENUM e.g. "hdfc"

      // Dynamically load bank processor
      const processorModuleName = `bankStatementProcessor-${validatedBankName.toUpperCase()}.js`
      const processBankStatement = require(`../Services/BankStatementProcessor/bankStatementProcessor-${validatedBankName.toUpperCase()}.js`)
      // if (!fs.existsSync(processBankStatement)) {
      //   console.error(`Processor file not found: ${processorModuleName}`)
      //   return res.status(400).json({ flag: 0, message: `BANK_PROCESSOR_NOT_FOUND_FOR_${validatedBankName.toUpperCase()}` })
      // }

      if (typeof processBankStatement !== 'function') {
        console.error(`processBankStatement function not found in ${processorModuleName}`)
        return res.status(500).json({ flag: 0, message: `INVALID_BANK_PROCESSOR_${validatedBankName.toUpperCase()}` })
      }

      const signedUrl = resource_url

      const response = await axios.get(signedUrl, {
        responseType: 'arraybuffer',
      })
      const workbook = XLSX.read(response.data, {
        cellDates: true,
        dateNF: 'd"/"m"/"yyyy', // Keep original date parsing for XLSX
      })
      const worksheet = workbook.Sheets[workbook.SheetNames[0]]
      const rawData = XLSX.utils.sheet_to_json(worksheet, {
        header: 1, // Keep header:1 to get array of arrays
        defval: null,
        blankrows: true,
      })

      // Call the bank-specific processor
      // Pass validatedBankName in case the processor needs it, though HDFC example doesn't use a second param.
      const result = processBankStatement(rawData)

      return res.status(200).json({
        success: true,
        message: 'File processed successfully',
        processed_data: result,
      })
    } catch (error) {
      console.error('Error:', error.message, error.stack)
      return res.status(500).json({
        flag: 0,
        message: 'Failed to process file', // More generic message
        error: error.message,
      })
    }
  },
}

module.exports = StatementToJsonController
