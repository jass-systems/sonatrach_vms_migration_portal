const express = require('express');
const router = express.Router();
const pool = require('../db');

// GET /api/stats - Statistiques pour le Tableau de bord
router.get('/', async (req, res) => {
  try {
    const totalRequests = await pool.query('SELECT COUNT(*) FROM migration_requests');
    const totalVMs = await pool.query('SELECT COUNT(*) FROM vms');
    const osDist = await pool.query('SELECT os_server, COUNT(*) as count FROM vms GROUP BY os_server');
    const stackDist = await pool.query('SELECT software_name, COUNT(*) as count FROM vm_software_stack WHERE exists = true GROUP BY software_name');

    res.json({
      totalRequests: parseInt(totalRequests.rows[0].count, 10),
      totalVMs: parseInt(totalVMs.rows[0].count, 10),
      osDistribution: osDist.rows,
      stackDistribution: stackDist.rows
    });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

module.exports = router;