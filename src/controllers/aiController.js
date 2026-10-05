const StudyMaterial = require("../../models/StudyMaterial");
const Summary = require("../../models/Summary");
const Flashcard = require("../../models/Flashcard");
const Quiz = require("../../models/Quiz");
const StudyPlan = require("../../models/StudyPlan");
const gemini = require("../utils/gemini");

async function loadOwnedMaterial(materialId, userId) {
  return StudyMaterial.findOne({ _id: materialId, userId });
}

async function summarize(req, res, next) {
  try {
    const material = await loadOwnedMaterial(req.params.id, req.user._id);
    if (!material) return res.status(404).json({ message: "Material not found" });

    const summaryText = await gemini.generateSummary(material.content);
    const summary = await Summary.create({
      userId: req.user._id,
      materialId: material._id,
      summary: summaryText,
    });
    res.status(201).json({ summary });
  } catch (err) {
    next(err);
  }
}

async function flashcards(req, res, next) {
  try {
    const { materialId, count } = req.body;
    const material = await loadOwnedMaterial(materialId, req.user._id);
    if (!material) return res.status(404).json({ message: "Material not found" });

    const cards = await gemini.generateFlashcards(material.content, count || 8);
    const saved = await Flashcard.insertMany(
      cards.map((c) => ({
        userId: req.user._id,
        materialId: material._id,
        question: c.question,
        answer: c.answer,
      }))
    );
    res.status(201).json({ flashcards: saved });
  } catch (err) {
    next(err);
  }
}

async function quiz(req, res, next) {
  try {
    const { materialId, count } = req.body;
    const material = await loadOwnedMaterial(materialId, req.user._id);
    if (!material) return res.status(404).json({ message: "Material not found" });

    const questions = await gemini.generateQuiz(material.content, count || 5);
    const savedQuiz = await Quiz.create({
      userId: req.user._id,
      materialId: material._id,
      questions,
    });
    res.status(201).json({ quiz: savedQuiz });
  } catch (err) {
    next(err);
  }
}

async function studyPlan(req, res, next) {
  try {
    const { subjects, examDate, hoursPerDay } = req.body;
    if (!subjects || !examDate) {
      return res.status(400).json({ message: "subjects and examDate are required" });
    }

    const planText = await gemini.generateStudyPlan(subjects, examDate, hoursPerDay || "2");
    const plan = await StudyPlan.create({
      userId: req.user._id,
      studyPlan: planText,
      examDate,
    });
    res.status(201).json({ studyPlan: plan });
  } catch (err) {
    next(err);
  }
}

async function ask(req, res, next) {
  try {
    const { question, materialId } = req.body;
    if (!question || !question.trim()) {
      return res.status(400).json({ message: "question is required" });
    }

    let context = null;
    if (materialId) {
      const material = await loadOwnedMaterial(materialId, req.user._id);
      if (!material) return res.status(404).json({ message: "Material not found" });
      context = material.content;
    }

    const answer = await gemini.askQuestion(question, context);
    res.json({ question, answer });
  } catch (err) {
    next(err);
  }
}

module.exports = { summarize, flashcards, quiz, studyPlan, ask };
