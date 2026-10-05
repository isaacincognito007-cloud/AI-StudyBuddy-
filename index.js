require("dotenv").config();

const express = require("express");
const cors = require("cors");
const helmet = require("helmet");
const morgan = require("morgan");
const rateLimit = require("express-rate-limit");

const connectDB = require("./src/utils/db");
const sanitizeInput = require("./src/middleware/sanitize");
const { errorHandler, notFound } = require("./src/middleware/errorHandler");

const authRoutes = require("./src/routes/authRoutes");
const materialRoutes = require("./src/routes/materialRoutes");
const aiRoutes = require("./src/routes/aiRoutes");
const adminRoutes = require("./src/routes/adminRoutes");

const app = express();

// --- core middleware ---
app.use(helmet());
app.use(cors());
app.use(morgan(process.env.NODE_ENV === "production" ? "combined" : "dev"));
app.use(express.json({ limit: "2mb" }));
app.use(sanitizeInput);

// A global limiter for the whole API; aiRoutes additionally applies its
// own tighter limit on top of this, since AI calls cost real quota.
app.use(
  rateLimit({
    windowMs: 15 * 60 * 1000,
    max: 200,
    message: { message: "Too many requests — please try again later." },
  })
);

app.get("/health", (req, res) => res.json({ status: "ok", time: new Date().toISOString() }));

// --- routes ---
app.use("/api/auth", authRoutes);
app.use("/api/materials", materialRoutes);
app.use("/api/ai", aiRoutes);
app.use("/api/admin", adminRoutes);

app.use(notFound);
app.use(errorHandler);

const PORT = process.env.PORT || 5000;

connectDB().then(() => {
  app.listen(PORT, () => console.log(`AI StudyBuddy backend running on port ${PORT}`));
});

module.exports = app;
