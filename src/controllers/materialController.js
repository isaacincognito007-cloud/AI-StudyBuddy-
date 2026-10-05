const fs = require("fs/promises");
const path = require("path");
const StudyMaterial = require("../../models/StudyMaterial");

// Only .txt/.md are read directly as text here to keep the reference
// implementation dependency-free; PDF/DOC text extraction is a natural
// next step (e.g. pdf-parse / mammoth) — see README "What's next".
async function extractText(file) {
  const ext = path.extname(file.originalname).toLowerCase();
  if (ext === ".txt" || ext === ".md") {
    return fs.readFile(file.path, "utf-8");
  }
  return null; // caller falls back to req.body.content for unsupported types
}

async function uploadMaterial(req, res, next) {
  try {
    const { title, subject, content } = req.body;
    let materialContent = content || "";

    if (req.file) {
      const extracted = await extractText(req.file);
      if (extracted !== null) materialContent = extracted;
    }

    if (!title || !materialContent.trim()) {
      return res.status(400).json({
        message: "A title and some content are required (paste text, or upload a .txt/.md file)",
      });
    }

    const material = await StudyMaterial.create({
      userId: req.user._id,
      title,
      subject,
      content: materialContent,
      fileName: req.file ? req.file.originalname : undefined,
    });

    res.status(201).json({ material });
  } catch (err) {
    next(err);
  }
}

async function listMaterials(req, res, next) {
  try {
    const { q, subject, page = 1, limit = 10 } = req.query;
    const filter = { userId: req.user._id };
    if (subject) filter.subject = subject;
    if (q) filter.$text = { $search: q };

    const skip = (Number(page) - 1) * Number(limit);
    const [materials, total] = await Promise.all([
      StudyMaterial.find(filter).sort({ createdAt: -1 }).skip(skip).limit(Number(limit)),
      StudyMaterial.countDocuments(filter),
    ]);

    res.json({
      materials,
      pagination: { page: Number(page), limit: Number(limit), total, pages: Math.ceil(total / limit) },
    });
  } catch (err) {
    next(err);
  }
}

async function downloadMaterial(req, res, next) {
  try {
    const material = await StudyMaterial.findOne({ _id: req.params.id, userId: req.user._id });
    if (!material) return res.status(404).json({ message: "Material not found" });
    res.setHeader("Content-Disposition", `attachment; filename="${material.title}.txt"`);
    res.type("text/plain").send(material.content);
  } catch (err) {
    next(err);
  }
}

async function getMaterial(req, res, next) {
  try {
    const material = await StudyMaterial.findOne({ _id: req.params.id, userId: req.user._id });
    if (!material) return res.status(404).json({ message: "Material not found" });
    res.json({ material });
  } catch (err) {
    next(err);
  }
}

async function updateMaterial(req, res, next) {
  try {
    const { title, subject, content } = req.body;
    const material = await StudyMaterial.findOneAndUpdate(
      { _id: req.params.id, userId: req.user._id },
      { $set: { title, subject, content } },
      { new: true, runValidators: true }
    );
    if (!material) return res.status(404).json({ message: "Material not found" });
    res.json({ material });
  } catch (err) {
    next(err);
  }
}

async function deleteMaterial(req, res, next) {
  try {
    const material = await StudyMaterial.findOneAndDelete({ _id: req.params.id, userId: req.user._id });
    if (!material) return res.status(404).json({ message: "Material not found" });
    res.json({ message: "Material deleted" });
  } catch (err) {
    next(err);
  }
}

module.exports = {
  uploadMaterial,
  listMaterials,
  getMaterial,
  updateMaterial,
  deleteMaterial,
  downloadMaterial,
};
