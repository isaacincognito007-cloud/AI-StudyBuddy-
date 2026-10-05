const mongoose = require("mongoose");

const summarySchema = new mongoose.Schema(
  {
    userId: { type: mongoose.Schema.Types.ObjectId, ref: "User", required: true },
    materialId: { type: mongoose.Schema.Types.ObjectId, ref: "StudyMaterial", required: true },
    summary: { type: String, required: true },
  },
  { timestamps: true }
);

module.exports = mongoose.model("Summary", summarySchema);
