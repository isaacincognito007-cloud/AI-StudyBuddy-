// Centralized error handler — every controller can just throw or call next(err)
// and requests never crash the process or leak stack traces to the client.
function errorHandler(err, req, res, next) {
  console.error(err.stack || err.message);

  if (err.name === "ValidationError") {
    return res.status(400).json({ message: "Validation failed", details: err.message });
  }
  if (err.code === 11000) {
    return res.status(409).json({ message: "Duplicate value", field: Object.keys(err.keyValue || {})[0] });
  }
  if (err instanceof multerErrorTypeCheck()) {
    return res.status(400).json({ message: err.message });
  }

  const status = err.statusCode || 500;
  res.status(status).json({ message: err.message || "Internal server error" });
}

function multerErrorTypeCheck() {
  return require("multer").MulterError;
}

function notFound(req, res) {
  res.status(404).json({ message: `Route not found: ${req.method} ${req.originalUrl}` });
}

module.exports = { errorHandler, notFound };
