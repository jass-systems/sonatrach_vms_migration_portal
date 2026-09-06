const express = require('express');
const router = express.Router();
const pool = require('../db');

// POST /api/auth/login
router.post('/login', async (req, res) => {
  const { username, password } = req.body;
  try {
    const result = await pool.query('SELECT * FROM users WHERE username = $1 AND password = $2', [username, password]);
    if (result.rows.length > 0) {
      res.json({
        success: true,
        user: {
          id: result.rows[0].id,
          username: result.rows[0].username,
          role: result.rows[0].role
        }
      });
    } else {
      res.status(401).json({ success: false, message: 'Identifiants invalides' });
    }
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

module.exports = router;