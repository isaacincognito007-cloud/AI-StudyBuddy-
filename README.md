# LearnMate — AI StudyBuddy

**An AI-powered learning assistance platform that turns a student's own study material into summaries, flashcards, quizzes, and personalized study plans — generated on demand by Google Gemini, not pulled from a template.**

---

## Why this exists

Rahul is a second-year engineering student juggling multiple semester exams while also trying to pick up new skills for internship season. His material is scattered — lecture notes, PDFs, textbooks, a few online articles, some handwritten pages — and with limited time, he struggles to figure out what actually matters, turn it into something revisable, and check whether he's actually understood it before the exam hits.

That's the gap LearnMate closes: upload what you have, and get back a summary to revise from, flashcards for active recall, a quiz to test yourself against, and a day-by-day plan for the time you actually have left — all grounded in *your* material, not generic content, and all sitting behind a real account so it's there next time you log in.

---

## Description

LearnMate (AI StudyBuddy) is built on **Node.js** and **Express.js** as a RESTful API, with **MongoDB** (via **Mongoose**) as the data layer. Every account is protected by **JWT authentication** and **Role-Based Access Control** — students can only ever touch their own data, and the handful of administrative actions (user management, usage analytics) are locked to an `admin` role that is never self-assignable at registration. Passwords are hashed with **bcryptjs** before they ever touch the database.

The core of the platform is its integration with the **Google Gemini API**: instead of filling in a fixed template, every summary, flashcard set, quiz, and study plan is generated fresh by sending the user's actual material to Gemini and asking it to work from that content specifically.

---

## Software Requirements

- **OS:** Windows 10/11, macOS, or Linux
- **Node.js** v18 or above
- **npm** v8 or above
- **MongoDB** — via MongoDB Atlas (free tier, no local install needed — see §11)
- **Postman** — for exercising and verifying the API
- **A code editor** — VS Code or similar

## Hardware Requirements

- **Processor:** Intel Core i5 (8th gen or newer) / AMD Ryzen 5 or equivalent
- **RAM:** 8 GB minimum
- **Storage:** ~1 GB free for dependencies and uploaded files

---

## 1. Technical Architecture

```
Client (Postman / future React app)
        │  HTTPS + JWT Bearer token
        ▼
┌─────────────────────────────────────────────┐
│ Express Server (index.js)                    │
│  helmet → cors → morgan → json parser →      │
│  sanitizeInput → global rate limiter          │
└───────────────────┬───────────────────────────┘
                     ▼
        ┌────────────┴────────────┐
        │         Routes           │
        │  /api/auth  /api/materials │
        │  /api/ai    /api/admin   │
        └────────────┬────────────┘
                      ▼
     ┌────────────────┴─────────────────┐
     │   protect (JWT)  +  authorize      │
     │        (role check)                │
     └────────────────┬─────────────────┘
                      ▼
              ┌───────┴────────┐
              │  Controllers    │  (business logic)
              └───────┬────────┘
              ┌────────┴────────┐
              ▼                 ▼
      ┌───────────────┐  ┌─────────────────┐
      │ Mongoose Models│  │  Gemini AI utility│
      │   → MongoDB    │  │  (@google/genai)  │
      └───────────────┘  └─────────────────┘
```

Every request passes through the same middleware chain before it reaches a route: security headers, CORS, request logging, JSON parsing, a NoSQL-injection guard, and rate limiting. AI routes sit behind a *second*, tighter rate limiter on top of the global one, since Gemini calls cost real API quota.

---

## 2. ER Diagram

| Entity | Key fields | Relationship |
|---|---|---|
| **User** | name, email (unique), password (hashed), role | 1 → N on everything below |
| **StudyMaterial** | userId (FK), title, subject, content, fileName | 1 → N Summary, Flashcard; 1 → 1 Quiz |
| **Summary** | userId (FK), materialId (FK), summary | N → 1 StudyMaterial |
| **Flashcard** | userId (FK), materialId (FK), question, answer | N → 1 StudyMaterial |
| **Quiz** | userId (FK), materialId (FK), questions[] (question, options, correctIndex) | N → 1 StudyMaterial |
| **StudyPlan** | userId (FK), studyPlan, examDate | N → 1 User |

**Known gap, documented rather than guessed at:** earlier drafts of the spec also name **Progress** and **Notifications** as collections, but no fields were ever defined for either, in this document or the ER diagram. Both are left out until their structure is actually specified.

---

## 3. Features

**Auth & account** — register / login (JWT, 7-day expiry), get/update profile, bcrypt-hashed passwords.

**Study materials** — upload as text or a file, list (paginated, searchable, filterable by subject), get one, update, delete, download as text.

**AI-powered generation**, all grounded in the uploaded material:
- Summary
- Flashcards (structured question/answer pairs)
- Multiple-choice quiz (structured, with a correct answer)
- Personalized study plan (subjects + exam date + hours/day → a day-by-day plan)
- Direct Q&A — ask a question, optionally grounded in one material

**Admin** — list/delete users, change roles, view all materials across every user, pull a usage analytics snapshot (user count, material count, AI generations by type).

**Security** — JWT + RBAC, bcrypt hashing, input validation on every write, a custom NoSQL-injection guard (the standard package for this is broken on Express 5 — see §10), two-tier rate limiting, centralized error handling, secrets only ever from environment variables.

---

## 4. Roles and Responsibilities

**Student** (the default — every self-registration lands here) — uploads and manages their own materials, generates and reviews their own AI content. Cannot see or affect any other user's data.

**Administrator** — everything a student can do, plus managing users and pulling system-wide analytics. There is deliberately no way to register directly as admin; `role` is forced to `"student"` server-side regardless of what's sent, and promotion only happens by an existing admin (or, for the very first admin, by editing the database directly — see §14).

---

## 5. User Flow

Register or log in → receive a JWT → upload a study material (paste text or attach a file) → pick an AI feature (summary, flashcards, quiz, study plan, or ask a direct question) → review the generated resource, which is saved for later, not thrown away → manage materials over time (search, update, delete, download) → (admin only) oversee users and pull analytics.

---

## 6. MVC Pattern

- **Model** (`/models`) — six Mongoose schemas; all reads/writes go through these.
- **Controller** (`/src/controllers`) — business logic: validates, calls the Gemini utility when AI generation is needed, shapes the response.
- **View** — no traditional view layer; this is a REST API, so the JSON response *is* the view, consumed by whatever client calls it.

---

## 7. Creating the Project Folder

```bash
mkdir ai-studybuddy-backend && cd ai-studybuddy-backend
npm init -y
npm install express mongoose bcryptjs jsonwebtoken cors dotenv multer @google/genai express-validator helmet morgan express-rate-limit
npm install --save-dev nodemon
```

(Already done in the delivered project — `npm install` is the only step you need to actually run.)

---

## 8. Server Setup

`index.js` is the single entry point: loads `.env`, builds the Express app, applies the middleware chain, mounts all four route groups, adds 404/error handling, and — only after `connectDB()` succeeds — starts listening on `process.env.PORT` (default 5000). If MongoDB can't be reached, the server never starts accepting requests rather than running half-broken.

`npm run dev` (nodemon, auto-restarts) for development, `npm start` for a plain run.

---

## 9. Backend Structure

```
ai-studybuddy-backend/
├── index.js
├── package.json
├── .env.example
├── .gitignore
├── models/
│   ├── User.js
│   ├── StudyMaterial.js
│   ├── Summary.js
│   ├── Flashcard.js
│   ├── Quiz.js
│   └── StudyPlan.js
├── src/
│   ├── utils/
│   │   ├── db.js
│   │   └── gemini.js
│   ├── middleware/
│   │   ├── auth.js
│   │   ├── upload.js
│   │   ├── validate.js
│   │   ├── sanitize.js
│   │   └── errorHandler.js
│   ├── controllers/
│   │   ├── authController.js
│   │   ├── materialController.js
│   │   ├── aiController.js
│   │   └── adminController.js
│   └── routes/
│       ├── authRoutes.js
│       ├── materialRoutes.js
│       ├── aiRoutes.js
│       └── adminRoutes.js
└── uploads/
```

---

## 10. Development and Explanation

- **`src/utils/db.js`** — `connectDB()` wraps `mongoose.connect()`; on failure it logs and exits rather than limping along without a database.
- **`src/middleware/auth.js`** — `protect` verifies the JWT and attaches `req.user`; `authorize(...roles)` is a factory, so `authorize("admin")` on a route means exactly that.
- **`src/middleware/sanitize.js`** — **not the standard approach, deliberately.** `express-mongo-sanitize` tries to reassign `req.query`, which Express 5 made a read-only getter — confirmed live, it 500s on every request. This hand-rolled version instead rejects (400) any request with a `$`-prefixed or dot-containing key, rather than trying to silently mutate input Express 5 won't let it mutate.
- **`src/utils/gemini.js`** — one shared client, model `gemini-2.5-flash`. Summaries and study plans return prose; flashcards and quizzes use Gemini's structured-output mode so the response is guaranteed-parseable JSON, not regex-matched free text.
- **Controllers** — every material query is scoped by `userId: req.user._id`, so a student can only ever touch their own data, enforced at the query level, not just the route level.

---

## 11. Configure MongoDB

1. Create a free account at MongoDB Atlas.
2. **Create a cluster** → free **M0** tier → any region.
3. **Database Access** → add a database user (username + password — save these).
4. **Network Access** → **Allow Access From Anywhere** (0.0.0.0/0) for local development.
5. **Connect → Drivers** → copy the connection string, replace `<username>`/`<password>`, add a database name before the `?`.

---

## 12. Create Database Connection

Paste the connection string from §11 into `.env` as `MONGO_URI`. `connectDB()` picks it up automatically on startup and logs `MongoDB connected: <host>` on success.

---

## 13. Create Schema and Models

Six schemas in `/models`, each with `{ timestamps: true }`:

- **User** — `name`, `email` (unique, lowercased), `password` (hashed, `select: false` by default), `role` (`student`/`admin`).
- **StudyMaterial** — `userId`, `title`, `subject`, `content`, `fileName`. Indexed on `userId`+`createdAt`, `subject`, and a text index on `title`+`content` for search.
- **Summary / Flashcard / Quiz / StudyPlan** — reference `userId` and `materialId` (StudyPlan is user-level, not material-specific), matching the ER diagram.

---

## 14. Steps for Project Execution

**What you need:** Node.js v18+, a MongoDB Atlas connection string, a free Gemini API key, and Postman.

**Setup:**
1. `npm install`
2. Copy `.env.example` to `.env`, fill in `MONGO_URI`, `GEMINI_API_KEY`, and any long random string for `JWT_SECRET`.
3. `npm run dev`
4. Confirm the console shows both `MongoDB connected: ...` and the server listening — if only one appears, something's wrong before Postman even enters the picture.

**Where to test:** import `AI-StudyBuddy.postman_collection.json` (included alongside this README) into Postman. Login/Register auto-save your JWT; uploading a material auto-saves its ID — nothing needs copy-pasting by hand.

**What to test, in order:** Health → Auth (register, login, profile, and one request with no token to confirm it 401s) → Study Materials (upload, list, search, get, download, update) → AI Features (summary, flashcards, quiz, study plan, ask — plus one deliberate injection attempt to confirm it 400s) → Admin (will 403 until you manually set a user's `role` to `"admin"` directly in Atlas's collection browser, then re-login so the new token reflects it — there's no API path to self-promote, by design).

**What's next:** PDF/DOC text extraction on upload (currently only .txt/.md are read automatically), the Progress/Notifications collections once their fields are defined, and the React frontend.

---

## Demo Video https://youtu.be/N5lNC06rn7w

## Code Files https://drive.google.com/drive/folders/1mfJrGr63VDaF4zhXO7j9DIS_EWGJWjK0?usp=sharing
