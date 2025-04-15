const multer = require("multer");
const path = require("path");



// Custom storage configuration to save with original file name
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, "uploads/");
  },
  filename: (req, file, cb) => {
    const ext = path.extname(file.originalname).toLowerCase();
    const baseName = path.basename(file.originalname, ext);
    const timestamp = Date.now();
    cb(null, `${baseName}-${timestamp}${ext}`); // adds timestamp to avoid overwriting
  },
});

// Configure multer for file uploads
const Upload = multer({
  storage: storage,
  fileFilter: (req, file, cb) => {
    const ext = path.extname(file.originalname).toLowerCase();
    if (ext === ".xls" || ext === ".xlsx") {
      cb(null, true);
    } else {
      cb(new Error("Only Excel files (.xls, .xlsx) are allowed!"), false);
    }
  },
  limits: {
    fileSize: 5 * 1024 * 1024, // 5MB limit
  },
});

module.exports = Upload;
