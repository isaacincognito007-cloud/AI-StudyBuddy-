const express = require("express");
const { body } = require("express-validator");
const validate = require("../middleware/validate");
const { protect } = require("../middleware/auth");
const upload = require("../middleware/upload");
const {
  uploadMaterial,
  listMaterials,
  getMaterial,
  updateMaterial,
  deleteMaterial,
  downloadMaterial,
} = require("../controllers/materialController");

const router = express.Router();

router.use(protect); // every route below requires a logged-in user

router.post("/upload", upload.single("file"), [body("title").trim().notEmpty()], validate, uploadMaterial);
router.get("/", listMaterials); // supports ?q=&subject=&page=&limit=
router.get("/:id", getMaterial);
router.get("/:id/download", downloadMaterial);
router.put("/:id", updateMaterial);
router.delete("/:id", deleteMaterial);

module.exports = router;
