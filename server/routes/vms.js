const express = require('express');
const router = express.Router();
const db = require('../db');

const getClient = async () => {
  if (typeof db.connect === 'function') return await db.connect();
  if (db.pool && typeof db.pool.connect === 'function') return await db.pool.connect();
  throw new Error('Database connection failed. Check db.js export.');
};

const getTableColumns = async (client, tableName) => {
  const res = await client.query(
    `SELECT column_name FROM information_schema.columns WHERE table_schema = 'public' AND table_name = $1`,
    [tableName]
  );
  return res.rows.map(r => r.column_name);
};

const formatVmResponse = (vm, stack = [], flows = [], compliance = []) => {
  return {
    ...vm,
    id: vm.id,
    structureInfo: {
      pole: vm.pole || 'ALGER',
      structure: vm.structure || 'DTI',
      responsable_structure: vm.responsable_structure || '/',
      responsable_service: vm.responsable_service || 'INTRANET',
      contact: vm.contact || ''
    },
    formPublication: {
      publication_type: vm.publication_type || 'Intranet',
      target_population: vm.target_population || '',
      app_name: vm.app_name || '',
      dns_entry: vm.dns_entry || 'N/A',
      ip_address: vm.ip_address || '',
      port: vm.port ? String(vm.port) : '443',
      os_server: vm.os_server || 'Windows Server'
    },
    securityParams: {
      dns_site_web: vm.dns_site_web || vm.dns_entry || 'N/A',
      ip_publique: vm.public_ip || vm.ip_publique || 'N/A',
      ip_interne: vm.ip_interne || vm.ip_address || '',
      ip_virtuelle_f5: vm.f5_virtual_ip || vm.ip_virtuelle_f5 || 'N/A',
      publication: vm.publication || 'DEV',
      date_derniere_maj: vm.date_derniere_maj || new Date().toISOString().slice(0, 16)
    },
    architecture_desc: vm.architecture_desc || '',
    softwareStack: stack.map(s => ({
      software_name: s.software_name,
      exists: Boolean(s.exists ?? s.is_present),
      version: s.version || ''
    })),
    networkFlows: flows.map(f => ({
      id: f.id ? `flow-${f.id}` : `flow-${Date.now()}`,
      source: f.source || '',
      destination: f.destination || '',
      service: f.service || '',
      port: f.port ? String(f.port) : '',
      flow_type: f.flow_type || '',
      description: f.description || ''
    })),
    securityCompliance: compliance.map(c => ({
      id: c.id,
      control_name: c.control_name,
      status: c.status || 'En attente',
      comments: c.comments || ''
    }))
  };
};

// GET /api/vms/:id - Récupération complète avec jointures
router.get('/:id', async (req, res) => {
  let client;
  try {
    const { id } = req.params;
    client = await getClient();

    const vmRes = await client.query('SELECT * FROM vms WHERE id = $1', [id]);
    if (vmRes.rows.length === 0) {
      return res.status(404).json({ message: 'Machine virtuelle introuvable.' });
    }

    const vm = vmRes.rows[0];
    const stackRes = await client.query('SELECT * FROM vm_software_stack WHERE vm_id = $1 ORDER BY id ASC', [id]);
    const flowsRes = await client.query('SELECT * FROM network_flows WHERE vm_id = $1 ORDER BY id ASC', [id]);
    const compRes = await client.query('SELECT * FROM security_compliance WHERE vm_id = $1 ORDER BY id ASC', [id]);

    return res.status(200).json(formatVmResponse(vm, stackRes.rows, flowsRes.rows, compRes.rows));
  } catch (err) {
    console.error('Erreur GET /api/vms/:id :', err);
    return res.status(500).json({ message: 'Erreur serveur', error: err.message });
  } finally {
    if (client) client.release();
  }
});

// DELETE /api/vms/:id
router.delete('/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const queryFn = typeof db.query === 'function' ? db.query.bind(db) : db.pool.query.bind(db.pool);
    const result = await queryFn('DELETE FROM vms WHERE id = $1 RETURNING *', [id]);

    if (result.rowCount === 0) {
      return res.status(404).json({ message: 'Machine virtuelle introuvable.' });
    }

    return res.status(200).json({ message: 'Machine virtuelle supprimée avec succès', deletedVm: result.rows[0] });
  } catch (err) {
    console.error('Erreur lors de la suppression :', err);
    return res.status(500).json({ message: 'Erreur serveur', error: err.message });
  }
});

// PUT /api/vms/:id - Mise à jour intégrale des 4 pages
router.put('/:id', async (req, res) => {
  let client;

  try {
    const vmId = req.params.id;
    client = await getClient();
    await client.query('BEGIN');

    const body = req.body || {};
    const p1 = body.structureInfo || {};
    const p2 = body.formPublication || {};
    const p3 = body.formService || {};
    const p4 = body.securityParams || {};

    const softwareStackArray = body.softwareStack || body.software_stack || [];
    const networkFlowsArray = body.networkFlows || body.network_flows || [];
    const securityComplianceArray = body.securityCompliance || body.security_compliance || [];

    const flatData = {
      pole: p1.pole || body.pole || 'ALGER',
      structure: p1.structure || body.structure || 'DTI',
      responsable_structure: p1.responsable_structure || body.responsable_structure || '/',
      responsable_service: p1.responsable_service || body.responsable_service || 'INTRANET',
      contact: p1.contact || body.contact || '',
      demandeur: p1.contact || body.demandeur || '',

      app_name: p2.app_name || body.app_name || 'VM',
      publication_type: p2.publication_type || body.publication_type || 'Intranet',
      target_population: p2.target_population || body.target_population || '',
      dns_entry: p2.dns_entry || body.dns_entry || 'N/A',
      ip_address: p2.ip_address || body.ip_address || '',
      port: p2.port ? String(p2.port) : (body.port ? String(body.port) : '443'),
      os_server: p2.os_server || body.os_server || 'Windows Server 2022',

      architecture_desc: body.architecture_desc || p3.architecture_desc || '',

      dns_site_web: p4.dns_site_web || body.dns_site_web || 'N/A',
      public_ip: p4.ip_publique || p4.public_ip || body.public_ip || 'N/A',
      ip_interne: p4.ip_interne || p2.ip_address || body.ip_address || '',
      f5_virtual_ip: p4.ip_virtuelle_f5 || p4.f5_virtual_ip || body.f5_virtual_ip || 'N/A',
      publication: p4.publication || body.publication || 'DEV',
      date_derniere_maj: p4.date_derniere_maj || body.date_derniere_maj || null
    };

    const vmsCols = await getTableColumns(client, 'vms');
    const setClauses = [];
    const updateValues = [];
    let paramIndex = 1;

    for (const col of vmsCols) {
      if (col === 'id') continue;
      if (flatData[col] !== undefined) {
        setClauses.push(`"${col}" = $${paramIndex}`);
        updateValues.push(flatData[col]);
        paramIndex++;
      }
    }

    let vmResult;
    if (setClauses.length > 0) {
      updateValues.push(vmId);
      const updateSql = `UPDATE vms SET ${setClauses.join(', ')} WHERE id = $${paramIndex} RETURNING *;`;
      vmResult = await client.query(updateSql, updateValues);
    } else {
      vmResult = await client.query('SELECT * FROM vms WHERE id = $1', [vmId]);
    }

    if (vmResult.rowCount === 0) {
      await client.query('ROLLBACK');
      return res.status(404).json({ message: `Machine virtuelle ID ${vmId} introuvable.` });
    }

    // 2. vm_software_stack
    if (Array.isArray(softwareStackArray)) {
      await client.query('DELETE FROM vm_software_stack WHERE vm_id = $1', [vmId]);
      const swCols = await getTableColumns(client, 'vm_software_stack');

      for (const item of softwareStackArray) {
        const rowData = {
          vm_id: vmId,
          software_name: item.software_name || item.name || 'Logiciel',
          version: item.version || '',
          exists: Boolean(item.exists ?? item.is_present),
          is_present: Boolean(item.exists ?? item.is_present)
        };

        const colsToInsert = swCols.filter(c => c !== 'id' && rowData[c] !== undefined);
        if (colsToInsert.length > 0) {
          const placeholders = colsToInsert.map((_, i) => `$${i + 1}`).join(', ');
          const values = colsToInsert.map(c => rowData[c]);
          const colNames = colsToInsert.map(c => `"${c}"`).join(', ');
          await client.query(`INSERT INTO vm_software_stack (${colNames}) VALUES (${placeholders})`, values);
        }
      }
    }

    // 3. network_flows
    if (Array.isArray(networkFlowsArray)) {
      await client.query('DELETE FROM network_flows WHERE vm_id = $1', [vmId]);
      const flowCols = await getTableColumns(client, 'network_flows');

      for (const item of networkFlowsArray) {
        const rawPort = String(item.port || '');
        const parsedPort = parseInt(rawPort.replace(/\D/g, ''), 10);
        const rowData = {
          vm_id: vmId,
          source: item.source || '',
          destination: item.destination || '',
          service: item.service || '',
          port: isNaN(parsedPort) ? null : parsedPort,
          flow_type: item.flow_type || '',
          description: item.description || ''
        };

        const colsToInsert = flowCols.filter(c => c !== 'id' && rowData[c] !== undefined);
        if (colsToInsert.length > 0) {
          const placeholders = colsToInsert.map((_, i) => `$${i + 1}`).join(', ');
          const values = colsToInsert.map(c => rowData[c]);
          const colNames = colsToInsert.map(c => `"${c}"`).join(', ');
          await client.query(`INSERT INTO network_flows (${colNames}) VALUES (${placeholders})`, values);
        }
      }
    }

    // 4. security_compliance
    if (Array.isArray(securityComplianceArray)) {
      await client.query('DELETE FROM security_compliance WHERE vm_id = $1', [vmId]);
      const secCols = await getTableColumns(client, 'security_compliance');

      for (const item of securityComplianceArray) {
        const rowData = {
          vm_id: vmId,
          control_name: item.control_name || item.label || item.title || item.name || 'Contrôle Sécurité',
          status: item.status || 'En attente',
          comments: item.comments || item.commentaires || ''
        };

        const colsToInsert = secCols.filter(c => c !== 'id' && rowData[c] !== undefined);
        if (colsToInsert.length > 0) {
          const placeholders = colsToInsert.map((_, i) => `$${i + 1}`).join(', ');
          const values = colsToInsert.map(c => rowData[c]);
          const colNames = colsToInsert.map(c => `"${c}"`).join(', ');
          await client.query(`INSERT INTO security_compliance (${colNames}) VALUES (${placeholders})`, values);
        }
      }
    }

    await client.query('COMMIT');

    const stackRes = await client.query('SELECT * FROM vm_software_stack WHERE vm_id = $1 ORDER BY id ASC', [vmId]);
    const flowsRes = await client.query('SELECT * FROM network_flows WHERE vm_id = $1 ORDER BY id ASC', [vmId]);
    const compRes = await client.query('SELECT * FROM security_compliance WHERE vm_id = $1 ORDER BY id ASC', [vmId]);

    const fullUpdatedVm = formatVmResponse(vmResult.rows[0], stackRes.rows, flowsRes.rows, compRes.rows);

    return res.status(200).json({
      message: 'Les 4 pages de la VM ont été enregistrées avec succès',
      vm: fullUpdatedVm
    });

  } catch (err) {
    if (client) await client.query('ROLLBACK');
    console.error('Erreur lors de la sauvegarde :', err);
    return res.status(500).json({ message: 'Erreur serveur lors de la sauvegarde', error: err.message });
  } finally {
    if (client) client.release();
  }
});

module.exports = router;