const multer = require("multer");
const path = require("path");

const ALLOWED_TYPES = [".pdf", ".txt", ".md", ".doc", ".docx"];
const MAX_FILE_SIZE_MB = 10;

const storage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, path.join(__dirname, "..", "..", "uploads")),
  filename: (req, file, cb) => {
    const unique = Date.now() + "-" + Math.round(Math.random() * 1e9);
    cb(null, unique + path.extname(file.originalname));
  },
});

function fileFilter(req, file, cb) {
  const ext = path.extname(file.originalname).toLowerCase();
  if (!ALLOWED_TYPES.includes(ext)) {
    return cb(new Error(`Unsupported file type "${ext}". Allowed: ${ALLOWED_TYPES.join(", ")}`));
  }
  cb(null, true);
}

const upload = multer({
  storage,
  fileFilter,
  limits: { fileSize: MAX_FILE_SIZE_MB * 1024 * 1024 },
});

module.exports = upload;
