const { GoogleGenAI, Type } = require("@google/genai");

const client = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY });
const MODEL = "gemini-2.5-flash";

/**
 * Generates a concise study summary from raw material text.
 * @param {string} content
 * @returns {Promise<string>}
 */
async function generateSummary(content) {
  const response = await client.models.generateContent({
    model: MODEL,
    contents:
      "Summarize the following study material into clear, concise revision notes. " +
      "Use short paragraphs or bullet points. Cover only what's actually in the " +
      "material below — do not add outside facts.\n\n---\n" +
      content,
  });
  return response.text?.trim() || "";
}

/**
 * Generates flashcards (question/answer pairs) from study material.
 * @param {string} content
 * @param {number} count
 * @returns {Promise<{question: string, answer: string}[]>}
 */
async function generateFlashcards(content, count = 8) {
  const response = await client.models.generateContent({
    model: MODEL,
    contents:
      `Create exactly ${count} flashcards for active-recall study, grounded ` +
      "strictly in the study material below. Each flashcard is a short " +
      "question and a concise answer.\n\n---\n" +
      content,
    config: {
      responseMimeType: "application/json",
      responseSchema: {
        type: Type.ARRAY,
        items: {
          type: Type.OBJECT,
          properties: {
            question: { type: Type.STRING },
            answer: { type: Type.STRING },
          },
          required: ["question", "answer"],
        },
      },
    },
  });
  return JSON.parse(response.text || "[]");
}

/**
 * Generates a multiple-choice quiz from study material.
 * @param {string} content
 * @param {number} count
 * @returns {Promise<{question: string, options: string[], correctIndex: number}[]>}
 */
async function generateQuiz(content, count = 5) {
  const response = await client.models.generateContent({
    model: MODEL,
    contents:
      `Create exactly ${count} multiple-choice quiz questions grounded strictly ` +
      "in the study material below. Each question has 4 options and exactly " +
      "one correct answer. Don't introduce facts the material doesn't cover.\n\n---\n" +
      content,
    config: {
      responseMimeType: "application/json",
      responseSchema: {
        type: Type.ARRAY,
        items: {
          type: Type.OBJECT,
          properties: {
            question: { type: Type.STRING },
            options: { type: Type.ARRAY, items: { type: Type.STRING } },
            correctIndex: { type: Type.INTEGER },
          },
          required: ["question", "options", "correctIndex"],
        },
      },
    },
  });
  return JSON.parse(response.text || "[]");
}

/**
 * Generates a personalized study plan.
 * @param {string} subjects - subjects/topics the user needs to cover
 * @param {string} examDate - ISO date string
 * @param {string} hoursPerDay - rough available study time
 * @returns {Promise<string>}
 */
async function generateStudyPlan(subjects, examDate, hoursPerDay) {
  const response = await client.models.generateContent({
    model: MODEL,
    contents:
      `Create a day-by-day study plan for a student preparing for an exam on ${examDate}. ` +
      `They have about ${hoursPerDay} hours available per day. Subjects/topics to cover: ` +
      `${subjects}. Be specific about what to study each day, and build in periodic ` +
      "review days rather than only covering new material once. Format as a clear, " +
      "readable day-by-day breakdown.",
  });
  return response.text?.trim() || "";
}

/**
 * Answers a free-form study question, optionally grounded in a specific
 * material. Covers both "answer learning queries" and "topic explanations"
 * from the spec — asking a question and getting a topic explained are the
 * same underlying capability.
 * @param {string} question
 * @param {string|null} context - optional material content to ground the answer in
 * @returns {Promise<string>}
 */
async function askQuestion(question, context) {
  const prompt = context
    ? `Using ONLY the study material below, answer the student's question. ` +
      `If the material doesn't cover it, say so rather than guessing.\n\n` +
      `Material:\n${context}\n\nQuestion: ${question}`
    : `Answer this study question clearly, at a level appropriate for a student. ` +
      `Include a brief real-world example if it helps understanding.\n\nQuestion: ${question}`;

  const response = await client.models.generateContent({ model: MODEL, contents: prompt });
  return response.text?.trim() || "";
}

module.exports = { generateSummary, generateFlashcards, generateQuiz, generateStudyPlan, askQuestion };
