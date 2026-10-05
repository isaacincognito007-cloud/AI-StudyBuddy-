const express = require("express");
const { protect, authorize } = require("../middleware/auth");
const {
  listUsers,
  updateUserRole,
  deleteUser,
  listAllMaterials,
  analytics,
} = require("../controllers/adminController");

const router = express.Router();

router.use(protect, authorize("admin")); // everything below is admin-only

router.get("/users", listUsers);
router.put("/users/:id/role", updateUserRole);
router.delete("/users/:id", deleteUser);
router.get("/materials", listAllMaterials);
router.get("/analytics", analytics);

module.exports = router;
