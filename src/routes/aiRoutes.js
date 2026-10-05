const express = require("express");
const rateLimit = require("express-rate-limit");
const { body } = require("express-validator");
const validate = require("../middleware/validate");
const { protect } = require("../middleware/auth");
const { summarize, flashcards, quiz, studyPlan, ask } = require("../controllers/aiController");

const router = express.Router();

// AI calls cost real API quota (the free tier is small — see README), so
// these get a tighter limit than the rest of the API, not just the global one.
const aiLimiter = rateLimit({
  windowMs: 60 * 1000,
  max: 15,
  message: { message: "Too many AI requests — wait a moment and try again." },
});

router.use(protect, aiLimiter);

router.post("/materials/:id/summarize", summarize);
router.post("/flashcards", [body("materialId").notEmpty()], validate, flashcards);
router.post("/quiz", [body("materialId").notEmpty()], validate, quiz);
router.post(
  "/study-plan",
  [body("subjects").notEmpty(), body("examDate").isISO8601()],
  validate,
  studyPlan
);
router.post("/ask", [body("question").trim().notEmpty()], validate, ask);

module.exports = router;
