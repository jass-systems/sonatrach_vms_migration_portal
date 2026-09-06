const express = require('express');
const router = express.Router();
const ExcelJS = require('exceljs');
const pool = require('../db');

// Helper interne pour générer la structure Excel à 4 onglets
async function buildExcelWorkbook(vmId) {
  const vmQuery = await pool.query(`
    SELECT v.*, m.pole, m.structure, m.responsable_structure, m.responsable_service, m.contact
    FROM vms v 
    JOIN migration_requests m ON v.migration_request_id = m.id 
    WHERE v.id = $1
  `, [vmId]);

  if (vmQuery.rows.length === 0) return null;
  const vm = vmQuery.rows[0];

  const stackQuery = await pool.query('SELECT * FROM vm_software_stack WHERE vm_id = $1', [vmId]);
  const flowsQuery = await pool.query('SELECT * FROM network_flows WHERE vm_id = $1', [vmId]);
  const secQuery = await pool.query('SELECT * FROM security_compliance WHERE vm_id = $1', [vmId]);

  const workbook = new ExcelJS.Workbook();

  // Onglet 1: 1. Principale
  const sheet1 = workbook.addWorksheet('1. Principale');
  sheet1.columns = [{ header: 'Champ', key: 'key', width: 32 }, { header: 'Valeur', key: 'val', width: 45 }];
  sheet1.addRows([
    { key: 'Pôle', val: vm.pole },
    { key: 'Structure', val: vm.structure },
    { key: 'Responsable de la Structure', val: vm.responsable_structure },
    { key: 'Responsable du Service à publier', val: vm.responsable_service },
    { key: 'Contact / Demandeur', val: vm.contact }
  ]);

  // Onglet 2: 2. Publication VM
  const sheet2 = workbook.addWorksheet('2. Publication VM');
  sheet2.columns = [
    { header: 'Nom Application', key: 'app_name', width: 25 },
    { header: 'Type Publication', key: 'publication_type', width: 20 },
    { header: 'Population Cible', key: 'target_population', width: 25 },
    { header: 'Entrée DNS', key: 'dns_entry', width: 25 },
    { header: 'Adresse IP', key: 'ip_address', width: 18 },
    { header: 'Port', key: 'port', width: 10 },
    { header: 'OS Serveur', key: 'os_server', width: 20 }
  ];
  sheet2.addRow({
    app_name: vm.app_name,
    publication_type: vm.publication_type,
    target_population: vm.target_population,
    dns_entry: vm.dns_entry || '',
    ip_address: vm.ip_address,
    port: vm.port,
    os_server: vm.os_server
  });

  sheet2.addRow([]);
  sheet2.addRow(['Logiciel / Composant', 'Installé', 'Version']);
  stackQuery.rows.forEach(s => {
    sheet2.addRow([s.software_name, s.exists ? 'OUI' : 'NON', s.version || '/']);
  });

  // Onglet 3: 3. Informations liées au service
  const sheet3 = workbook.addWorksheet('3. Informations liées au service');
  sheet3.addRow(['Description de l\'architecture:', vm.architecture_desc || 'N/A']);
  sheet3.addRow([]);
  sheet3.addRow(['Source', 'Destination', 'Service', 'Port', 'Type de Flux', 'Description']);
  flowsQuery.rows.forEach(f => {
    sheet3.addRow([f.source, f.destination, f.service, f.port, f.flow_type, f.description || '']);
  });

  // Onglet 4: 4. Suivi des Non-conformités
  const sheet4 = workbook.addWorksheet('4. Suivi des Non-conformités');
  sheet4.columns = [
    { header: 'Contrôle Sécurité', key: 'control_name', width: 45 },
    { header: 'Statut', key: 'status', width: 20 },
    { header: 'Commentaires', key: 'comments', width: 35 }
  ];
  secQuery.rows.forEach(s => {
    sheet4.addRow([s.control_name, s.status, s.comments || '']);
  });

  const sanitizedName = (vm.app_name || 'VM').replace(/[^a-zA-Z0-9_-]/g, '_');
  return { workbook, fileName: `Fiche_VM_${sanitizedName}.xlsx` };
}

// GET /api/vms/:id/export
router.get('/vms/:id/export', async (req, res) => {
  try {
    const data = await buildExcelWorkbook(req.params.id);
    if (!data) return res.status(404).send('VM introuvable');

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename=${data.fileName}`);

    await data.workbook.xlsx.write(res);
    res.end();
  } catch (err) {
    res.status(500).send(err.message);
  }
});

// GET /api/migrations/:id/export
router.get('/migrations/:id/export', async (req, res) => {
  try {
    const vmRes = await pool.query('SELECT id FROM vms WHERE migration_request_id = $1 ORDER BY id ASC LIMIT 1', [req.params.id]);
    
    if (vmRes.rows.length === 0) {
      return res.status(404).send('Aucune VM associée à cette demande pour générer l\'export Excel');
    }

    const data = await buildExcelWorkbook(vmRes.rows[0].id);
    if (!data) return res.status(404).send('Données introuvables');

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename=Demande_${req.params.id}_${data.fileName}`);

    await data.workbook.xlsx.write(res);
    res.end();
  } catch (err) {
    res.status(500).send(err.message);
  }
});

module.exports = router;