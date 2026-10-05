const jwt = require("jsonwebtoken");
const User = require("../../models/User");

function signToken(user) {
  return jwt.sign({ id: user._id, role: user.role }, process.env.JWT_SECRET, {
    expiresIn: process.env.JWT_EXPIRES_IN || "7d",
  });
}

function toPublicUser(user) {
  return { id: user._id, name: user.name, email: user.email, role: user.role };
}

async function register(req, res, next) {
  try {
    const { name, email, password, role } = req.body;
    if (!name || !email || !password) {
      return res.status(400).json({ message: "name, email, and password are required" });
    }

    const existing = await User.findOne({ email: email.toLowerCase() });
    if (existing) {
      return res.status(409).json({ message: "An account with this email already exists" });
    }

    // Only allow self-registering as a student; admins are promoted separately,
    // never chosen by the person signing up.
    const user = await User.create({ name, email, password, role: "student" });
    void role; // intentionally ignored — see comment above

    res.status(201).json({ token: signToken(user), user: toPublicUser(user) });
  } catch (err) {
    next(err);
  }
}

async function login(req, res, next) {
  try {
    const { email, password } = req.body;
    if (!email || !password) {
      return res.status(400).json({ message: "email and password are required" });
    }

    const user = await User.findOne({ email: email.toLowerCase() }).select("+password");
    if (!user || !(await user.comparePassword(password))) {
      return res.status(401).json({ message: "Invalid email or password" });
    }

    res.json({ token: signToken(user), user: toPublicUser(user) });
  } catch (err) {
    next(err);
  }
}

async function me(req, res) {
  res.json({ user: toPublicUser(req.user) });
}

async function updateProfile(req, res, next) {
  try {
    const { name } = req.body; // email/role/password intentionally not editable here
    if (!name || !name.trim()) {
      return res.status(400).json({ message: "name is required" });
    }
    req.user.name = name.trim();
    await req.user.save();
    res.json({ user: toPublicUser(req.user) });
  } catch (err) {
    next(err);
  }
}

module.exports = { register, login, me, updateProfile };
