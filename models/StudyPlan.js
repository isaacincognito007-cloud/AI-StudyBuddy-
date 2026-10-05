const mongoose = require("mongoose");

const studyPlanSchema = new mongoose.Schema(
  {
    userId: { type: mongoose.Schema.Types.ObjectId, ref: "User", required: true },
    studyPlan: { type: String, required: true },
    examDate: { type: Date },
  },
  { timestamps: true }
);

module.exports = mongoose.model("StudyPlan", studyPlanSchema);
