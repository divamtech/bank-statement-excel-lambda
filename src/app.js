require("dotenv").config();
const express = require("express");
const StatementToJsonController = require('./controller/statementToJson.controller.js');
const app = express();

// Middleware
app.use(express.json());

// Route
app.post(
  "/api/bank_statement_processor",
  StatementToJsonController.convert
);

// 404 Handler - Alternative syntax
app.use((req, res, next) => {
  const error = new Error(`Not Found - ${req.originalUrl}`);
  error.status = 404;
  next(error);
});

// Error handler
app.use((err, req, res, next) => {
  res.status(err.status || 500).json({
    success: false,
    message: err.message
  });
});

app.listen(process.env.PORT || 3000, () => {
  console.log(`Server running on port ${process.env.PORT || 3000}`);
});