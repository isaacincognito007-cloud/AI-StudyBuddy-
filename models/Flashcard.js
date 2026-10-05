const mongoose = require("mongoose");

const flashcardSchema = new mongoose.Schema(
  {
    userId: { type: mongoose.Schema.Types.ObjectId, ref: "User", required: true },
    materialId: { type: mongoose.Schema.Types.ObjectId, ref: "StudyMaterial", required: true },
    question: { type: String, required: true },
    answer: { type: String, required: true },
  },
  { timestamps: true }
);

module.exports = mongoose.model("Flashcard", flashcardSchema);
