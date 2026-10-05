const mongoose = require("mongoose");

const quizQuestionSchema = new mongoose.Schema(
  {
    question: { type: String, required: true },
    options: { type: [String], required: true },
    correctIndex: { type: Number, required: true },
  },
  { _id: false }
);

const quizSchema = new mongoose.Schema(
  {
    userId: { type: mongoose.Schema.Types.ObjectId, ref: "User", required: true },
    materialId: { type: mongoose.Schema.Types.ObjectId, ref: "StudyMaterial", required: true },
    questions: { type: [quizQuestionSchema], required: true },
  },
  { timestamps: true }
);

module.exports = mongoose.model("Quiz", quizSchema);
