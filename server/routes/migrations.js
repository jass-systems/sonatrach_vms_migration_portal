const express = require('express');
const router = express.Router();
const pool = require('../db');

// =========================================================================
// GET /api/migrations - Liste des demandes avec le nombre de VMs rattachées
// =========================================================================
router.get('/', async (req, res) => {
  try {
    const result = await pool.query(`
      SELECT m.*, COUNT(v.id)::int as vm_count 
      FROM migration_requests m 
      LEFT JOIN vms v ON v.migration_request_id = m.id 
      GROUP BY m.id 
      ORDER BY m.created_at DESC
    `);
    res.json(result.rows);
  } catch (err) {
    console.error('Erreur GET /api/migrations :', err);
    res.status(500).json({ error: err.message });
  }
});

// =========================================================================
// POST /api/migrations - Initialiser une demande globale
// =========================================================================
router.post('/', async (req, res) => {
  const { title, pole, structure, responsable_structure, responsable_service, contact } = req.body;
  try {
    const result = await pool.query(
      `INSERT INTO migration_requests (title, pole, structure, responsable_structure, responsable_service, contact)
       VALUES ($1, $2, $3, $4, $5, $6) RETURNING *`,
      [
        title,
        pole || 'ALGER',
        structure || 'TRC Siège / EXP',
        responsable_structure || '/',
        responsable_service || 'INTRANET',
        contact || 'Berkat Siham'
      ]
    );
    res.status(201).json(result.rows[0]);
  } catch (err) {
    console.error('Erreur POST /api/migrations :', err);
    res.status(500).json({ error: err.message });
  }
});

// =========================================================================
// DELETE /api/migrations/:id - Supprimer une demande complète
// (Les VMs et sous-tables sont supprimées grâce à ON DELETE CASCADE)
// =========================================================================
router.delete('/:id', async (req, res) => {
  const { id } = req.params;
  try {
    const result = await pool.query('DELETE FROM migration_requests WHERE id = $1 RETURNING *', [id]);
    
    if (result.rowCount === 0) {
      return res.status(404).json({ error: 'Demande introuvable.' });
    }

    res.json({ success: true, message: 'Demande et toutes ses VMs supprimées avec succès' });
  } catch (err) {
    console.error('Erreur DELETE /api/migrations/:id :', err);
    res.status(500).json({ error: err.message });
  }
});

// =========================================================================
// GET /api/migrations/:id/vms - Obtenir la liste COMPLÈTE des VMs d'une demande
// (Inclut la stack logicielle, les flux réseau et la conformité sécurité)
// =========================================================================
router.get('/:id/vms', async (req, res) => {
  try {
    const migrationId = req.params.id;

    // 1. Récupérer la demande parente pour les infos de structure
    const reqResult = await pool.query('SELECT * FROM migration_requests WHERE id = $1', [migrationId]);
    const migrationReq = reqResult.rows[0] || {};

    // 2. Récupérer toutes les VMs de cette demande
    const vmsResult = await pool.query(
      'SELECT * FROM vms WHERE migration_request_id = $1 ORDER BY id ASC',
      [migrationId]
    );

    const vms = vmsResult.rows;

    // 3. Charger les sous-tables pour chaque VM
    for (let vm of vms) {
      // Stack logicielle
      const swRes = await pool.query(
        'SELECT id, software_name, exists, version FROM vm_software_stack WHERE vm_id = $1 ORDER BY id ASC',
        [vm.id]
      );
      vm.softwareStack = swRes.rows;
      vm.software_stack = swRes.rows;

      // Matrice des flux réseau
      const flowsRes = await pool.query(
        'SELECT id, source, destination, service, port, flow_type, description FROM network_flows WHERE vm_id = $1 ORDER BY id ASC',
        [vm.id]
      );
      vm.networkFlows = flowsRes.rows;
      vm.network_flows = flowsRes.rows;

      // Suivi de la sécurité SI
      const secRes = await pool.query(
        'SELECT id, control_name, status, comments FROM security_compliance WHERE vm_id = $1 ORDER BY id ASC',
        [vm.id]
      );
      vm.securityCompliance = secRes.rows;
      vm.security_compliance = secRes.rows;

      // Structuration des objets pour le formulaire React & l'Export Excel
      vm.structureInfo = {
        pole: migrationReq.pole || 'ALGER',
        structure: migrationReq.structure || 'DTI',
        responsable_structure: migrationReq.responsable_structure || '/',
        responsable_service: migrationReq.responsable_service || 'INTRANET',
        contact: migrationReq.contact || ''
      };

      vm.formPublication = {
        publication_type: vm.publication_type,
        target_population: vm.target_population,
        app_name: vm.app_name,
        dns_entry: vm.dns_entry,
        ip_address: vm.ip_address,
        port: vm.port,
        os_server: vm.os_server
      };

      vm.securityParams = {
        dns_site_web: vm.dns_entry || 'N/A',
        ip_publique: vm.public_ip || 'N/A',
        ip_interne: vm.ip_address || '',
        ip_virtuelle_f5: vm.f5_virtual_ip || 'N/A',
        publication: vm.publication_type === 'Internet' ? 'PROD' : 'DEV',
        date_derniere_maj: vm.created_at ? new Date(vm.created_at).toISOString().slice(0, 16) : ''
      };
    }

    res.json(vms);
  } catch (err) {
    console.error('Erreur GET /api/migrations/:id/vms :', err);
    res.status(500).json({ error: err.message });
  }
});

// =========================================================================
// POST /api/migrations/:id/vms - Saisir une nouvelle VM avec ses 4 volets
// =========================================================================
router.post('/:id/vms', async (req, res) => {
  const migration_request_id = req.params.id;
  const {
    formPublication = {},
    softwareStack = [],
    networkFlows = [],
    securityCompliance = [],
    securityParams = {},
    architecture_desc = ''
  } = req.body;

  const client = await pool.connect();
  try {
    await client.query('BEGIN'); // Début de transaction SQL

    // 1. Insertion dans la table principale 'vms'
    const vmRes = await client.query(
      `INSERT INTO vms (
        migration_request_id, app_name, publication_type, target_population, 
        dns_entry, ip_address, port, os_server, architecture_desc, public_ip, f5_virtual_ip
      )
      VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11) RETURNING *`,
      [
        migration_request_id,
        formPublication.app_name || 'VM',
        formPublication.publication_type || 'Intranet',
        formPublication.target_population || '',
        formPublication.dns_entry || 'N/A',
        formPublication.ip_address || '',
        formPublication.port || '443',
        formPublication.os_server || 'Windows Server 2022',
        architecture_desc || '',
        securityParams.ip_publique || 'N/A',
        securityParams.ip_virtuelle_f5 || 'N/A'
      ]
    );

    const newVm = vmRes.rows[0];
    const vmId = newVm.id;

    // 2. Insertion de la Stack Logicielle (vm_software_stack)
    for (const sw of softwareStack) {
      const isPresent = sw.exists !== undefined ? sw.exists : (sw.is_present !== undefined ? sw.is_present : false);
      if (isPresent) {
        await client.query(
          `INSERT INTO vm_software_stack (vm_id, software_name, exists, version) VALUES ($1, $2, $3, $4)`,
          [vmId, sw.software_name, true, sw.version || '']
        );
      }
    }

    // 3. Insertion de la Matrice des Flux (network_flows)
    for (const flow of networkFlows) {
      if (flow.source || flow.destination) {
        await client.query(
          `INSERT INTO network_flows (vm_id, source, destination, service, port, flow_type, description) 
           VALUES ($1, $2, $3, $4, $5, $6, $7)`,
          [
            vmId, 
            flow.source || '', 
            flow.destination || '', 
            flow.service || 'TCP/443', 
            flow.port || '443', 
            flow.flow_type || 'Flux applicatif Web', 
            flow.description || ''
          ]
        );
      }
    }

    // 4. Insertion des Contrôles Sécurité (security_compliance)
    for (const sec of securityCompliance) {
      await client.query(
        `INSERT INTO security_compliance (vm_id, control_name, status, comments) VALUES ($1, $2, $3, $4)`,
        [vmId, sec.control_name, sec.status || 'En attente', sec.comments || '']
      );
    }

    await client.query('COMMIT'); // Validation finale
    res.status(201).json({ success: true, vm: newVm });
  } catch (err) {
    await client.query('ROLLBACK'); // Annulation en cas d'erreur
    console.error('Erreur POST /api/migrations/:id/vms :', err);
    res.status(500).json({ error: err.message });
  } finally {
    client.release();
  }
});

// =========================================================================
// DELETE /api/migrations/vms/:id - Supprimer une VM individuelle
// =========================================================================
const deleteVM = async (req, res) => {
  const vmId = req.params.id;
  try {
    // Grâce à ON DELETE CASCADE dans PostgreSQL schema.sql,
    // la suppression dans vms supprime automatiquement les tables filles.
    const result = await pool.query('DELETE FROM vms WHERE id = $1 RETURNING *', [vmId]);

    if (result.rowCount === 0) {
      return res.status(404).json({ error: 'Machine virtuelle introuvable.' });
    }

    res.json({ success: true, message: 'Machine virtuelle supprimée avec succès', deletedVm: result.rows[0] });
  } catch (err) {
    console.error('Erreur DELETE VM :', err);
    res.status(500).json({ error: err.message });
  }
};

router.delete('/vms/:id', deleteVM);

module.exports = router;