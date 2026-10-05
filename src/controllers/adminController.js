const User = require("../../models/User");
const StudyMaterial = require("../../models/StudyMaterial");
const Summary = require("../../models/Summary");
const Flashcard = require("../../models/Flashcard");
const Quiz = require("../../models/Quiz");
const StudyPlan = require("../../models/StudyPlan");

async function listUsers(req, res, next) {
  try {
    const users = await User.find().sort({ createdAt: -1 });
    res.json({ users });
  } catch (err) {
    next(err);
  }
}

async function updateUserRole(req, res, next) {
  try {
    const { role } = req.body;
    if (!["student", "admin"].includes(role)) {
      return res.status(400).json({ message: "role must be 'student' or 'admin'" });
    }
    const user = await User.findByIdAndUpdate(req.params.id, { role }, { new: true });
    if (!user) return res.status(404).json({ message: "User not found" });
    res.json({ user });
  } catch (err) {
    next(err);
  }
}

async function deleteUser(req, res, next) {
  try {
    const user = await User.findByIdAndDelete(req.params.id);
    if (!user) return res.status(404).json({ message: "User not found" });
    res.json({ message: "User deleted" });
  } catch (err) {
    next(err);
  }
}

async function listAllMaterials(req, res, next) {
  try {
    const materials = await StudyMaterial.find().populate("userId", "name email").sort({ createdAt: -1 });
    res.json({ materials });
  } catch (err) {
    next(err);
  }
}

// Covers "Monitor AI Usage" + "View Learning Analytics" + "View Reports" —
// a single analytics snapshot rather than three separate, vaguely-scoped
// endpoints with no defined shape in the spec.
async function analytics(req, res, next) {
  try {
    const [userCount, materialCount, summaryCount, flashcardCount, quizCount, planCount] =
      await Promise.all([
        User.countDocuments(),
        StudyMaterial.countDocuments(),
        Summary.countDocuments(),
        Flashcard.countDocuments(),
        Quiz.countDocuments(),
        StudyPlan.countDocuments(),
      ]);

    res.json({
      totalUsers: userCount,
      totalMaterials: materialCount,
      aiUsage: {
        summaries: summaryCount,
        flashcards: flashcardCount,
        quizzes: quizCount,
        studyPlans: planCount,
        total: summaryCount + flashcardCount + quizCount + planCount,
      },
    });
  } catch (err) {
    next(err);
  }
}

module.exports = { listUsers, updateUserRole, deleteUser, listAllMaterials, analytics };
