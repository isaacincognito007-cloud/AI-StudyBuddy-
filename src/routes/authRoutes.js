const express = require("express");
const { body } = require("express-validator");
const validate = require("../middleware/validate");
const { protect } = require("../middleware/auth");
const { register, login, me, updateProfile } = require("../controllers/authController");

const router = express.Router();

router.post(
  "/register",
  [
    body("name").trim().notEmpty().withMessage("name is required"),
    body("email").isEmail().withMessage("a valid email is required"),
    body("password").isLength({ min: 6 }).withMessage("password must be at least 6 characters"),
  ],
  validate,
  register
);

router.post(
  "/login",
  [body("email").isEmail(), body("password").notEmpty()],
  validate,
  login
);

router.get("/me", protect, me);
router.put("/me", protect, [body("name").trim().notEmpty()], validate, updateProfile);

module.exports = router;
