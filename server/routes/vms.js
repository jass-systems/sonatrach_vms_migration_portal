const express = require('express');
const router = express.Router();
const db = require('../db'); // Votre pool de connexion pg

// =========================================================================
// SUPPRIMER UNE VM : DELETE /api/vms/:id
// Grâce à "ON DELETE CASCADE", les tables filles sont nettoyées automatiquement
// =========================================================================
router.delete('/:id', async (req, res) => {
  try {
    const { id } = req.params;

    const result = await db.query('DELETE FROM vms WHERE id = $1 RETURNING *', [id]);

    if (result.rowCount === 0) {
      return res.status(404).json({ message: 'Machine virtuelle introuvable.' });
    }

    res.status(200).json({ 
      message: 'Machine virtuelle et ses dépendances supprimées avec succès',
      deletedVm: result.rows[0] 
    });
  } catch (err) {
    console.error('Erreur lors de la suppression de la VM :', err);
    res.status(500).json({ message: 'Erreur serveur lors de la suppression de la VM', error: err.message });
  }
});

// =========================================================================
// MODIFIER UNE VM SPÉCIFIQUE : PUT /api/vms/:id (avec Transaction PostgreSQL)
// =========================================================================
router.put('/:id', async (req, res) => {
  const client = await db.pool.connect();
  try {
    const { id } = req.params;
    const {
      formPublication = {},
      softwareStack = [],
      networkFlows = [],
      securityCompliance = [],
      securityParams = {},
      architecture_desc = ''
    } = req.body;

    await client.query('BEGIN'); // Début de transaction SQL

    // 1. Mise à jour de la table principale 'vms'
    const updateVmQuery = `
      UPDATE vms SET
        app_name = $1,
        publication_type = $2,
        target_population = $3,
        dns_entry = $4,
        ip_address = $5,
        port = $6,
        os_server = $7,
        architecture_desc = $8,
        public_ip = $9,
        f5_virtual_ip = $10
      WHERE id = $11
      RETURNING *;
    `;

    const vmValues = [
      formPublication.app_name || 'VM',
      formPublication.publication_type || 'Intranet',
      formPublication.target_population || '',
      formPublication.dns_entry || 'N/A',
      formPublication.ip_address || '',
      formPublication.port || '443',
      formPublication.os_server || 'Windows Server 2022',
      architecture_desc,
      securityParams.ip_publique || 'N/A',
      securityParams.ip_virtuelle_f5 || 'N/A',
      id
    ];

    const vmResult = await client.query(updateVmQuery, vmValues);
    if (vmResult.rowCount === 0) {
      await client.query('ROLLBACK');
      return res.status(404).json({ message: 'Machine virtuelle introuvable.' });
    }

    // 2. Mettre à jour la Stack Logicielle (Suppression + Réinsertion)
    await client.query('DELETE FROM vm_software_stack WHERE vm_id = $1', [id]);
    for (const sw of softwareStack) {
      await client.query(
        'INSERT INTO vm_software_stack (vm_id, software_name, is_present, version) VALUES ($1, $2, $3, $4)',
        [id, sw.software_name, sw.exists || sw.is_present || false, sw.version || '']
      );
    }

    // 3. Mettre à jour la Matrice des Flux (Suppression + Réinsertion)
    await client.query('DELETE FROM network_flows WHERE vm_id = $1', [id]);
    for (const flow of networkFlows) {
      await client.query(
        'INSERT INTO network_flows (vm_id, source, destination, service, port, flow_type, description) VALUES ($1, $2, $3, $4, $5, $6, $7)',
        [id, flow.source, flow.destination, flow.service, flow.port, flow.flow_type, flow.description || '']
      );
    }

    // 4. Mettre à jour les Contrôles Sécurité (Suppression + Réinsertion)
    await client.query('DELETE FROM security_compliance WHERE vm_id = $1', [id]);
    for (const ctrl of securityCompliance) {
      await client.query(
        'INSERT INTO security_compliance (vm_id, control_name, status, comments) VALUES ($1, $2, $3, $4)',
        [id, ctrl.control_name, ctrl.status, ctrl.comments || '']
      );
    }

    await client.query('COMMIT'); // Validation des modifications SQL

    res.status(200).json({
      message: 'VM mise à jour avec succès',
      vm: vmResult.rows[0]
    });
  } catch (err) {
    await client.query('ROLLBACK'); // Annulation en cas d'erreur
    console.error('Erreur modification VM :', err);
    res.status(500).json({ message: 'Erreur serveur lors de la modification de la VM', error: err.message });
  } finally {
    client.release();
  }
});

module.exports = router;