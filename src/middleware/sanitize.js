// express-mongo-sanitize (the "standard" package for this) is broken under
// Express 5: it tries to reassign req.query, which Express 5 made a
// read-only getter — confirmed live, it 500s on every single request.
// This is a small hand-rolled replacement: instead of trying to mutate
// input (which Express 5's req.query doesn't support — it's a derived
// getter, not a stored object), it REJECTS any request containing a
// MongoDB-operator-shaped key. Rejecting suspicious input outright is
// arguably better practice than silently rewriting it anyway.

function hasInjectionAttempt(obj, depth = 0) {
  if (depth > 8 || obj === null || typeof obj !== "object") return false;
  for (const key of Object.keys(obj)) {
    if (key.startsWith("$") || key.includes(".")) return true;
    if (hasInjectionAttempt(obj[key], depth + 1)) return true;
  }
  return false;
}

function sanitizeInput(req, res, next) {
  if (
    hasInjectionAttempt(req.body) ||
    hasInjectionAttempt(req.query) ||
    hasInjectionAttempt(req.params)
  ) {
    return res.status(400).json({ message: "Request rejected: invalid characters in input keys" });
  }
  next();
}

module.exports = sanitizeInput;
