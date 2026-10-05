const mongoose = require("mongoose");

const studyMaterialSchema = new mongoose.Schema(
  {
    userId: { type: mongoose.Schema.Types.ObjectId, ref: "User", required: true },
    title: { type: String, required: true, trim: true },
    subject: { type: String, trim: true },
    content: { type: String, required: true }, // extracted/raw text used as AI input
    fileName: { type: String }, // original uploaded file name, if any
  },
  { timestamps: true }
);

// Database Indexing (per the spec's Optimization checklist): userId is
// queried on every request, subject supports filtering, and a text index
// on title+content backs the search endpoint.
studyMaterialSchema.index({ userId: 1, createdAt: -1 });
studyMaterialSchema.index({ subject: 1 });
studyMaterialSchema.index({ title: "text", content: "text" });

module.exports = mongoose.model("StudyMaterial", studyMaterialSchema);
