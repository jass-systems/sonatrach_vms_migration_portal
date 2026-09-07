const express = require('express');
const router = express.Router();
const ExcelJS = require('exceljs');
const path = require('path');

router.post('/export-formulaire', async (req, res) => {
  try {
    const data = req.body || {};
    const templatePath = path.join(__dirname, '../formulaire_création_VM.xlsx');

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(templatePath);

    // Data extraction & fallback normalization
    const pole = data.pole || data.structureInfo?.pole || 'ALGER';
    const structure = data.structure || data.structureInfo?.structure || 'DTI';
    const responsable_structure = data.responsable_structure || data.structureInfo?.responsable_structure || '/';
    const responsable_service = data.responsable_service || data.structureInfo?.responsable_service || '/';
    const contact = data.contact || data.structureInfo?.contact || '';

    const population = data.population_exploitante || data.target_population || data.formPublication?.target_population || '';
    const appName = data.nom_application || data.app_name || data.formPublication?.app_name || '';
    const ipAddress = data.ip_address || data.formPublication?.ip_address || '';
    const port = data.port || data.formPublication?.port || '443';
    const os = data.os || data.os_server || data.formPublication?.os_server || '';

    const softwareStack = data.softwareStack || data.software_stack || [];
    const architectureDesc = data.architecture_desc || data.architectureDesc || '';
    const networkFlows = data.flux || data.networkFlows || data.network_flows || [];

    const dnsSiteWeb = data.dns_site_web || data.securityParams?.dns_site_web || data.dns_entry || 'N/A';
    const ipPublique = data.ip_publique || data.public_ip || data.securityParams?.ip_publique || 'N/A';
    const ipInterne = data.ip_interne || data.securityParams?.ip_interne || ipAddress || '';
    const ipVirtuelleF5 = data.ip_virtuelle_f5 || data.f5_virtual_ip || data.securityParams?.ip_virtuelle_f5 || 'N/A';
    const publication = data.publication || data.securityParams?.publication || 'DEV';
    const securityCompliance = data.securityCompliance || data.security_compliance || [];

    // 1. ONGLET : Principale
    const sheet1 = workbook.getWorksheet('Principale') || workbook.getWorksheet(1);
    if (sheet1) {
      sheet1.getCell('B29').value = pole;
      sheet1.getCell('B30').value = structure;
      sheet1.getCell('B31').value = responsable_structure;
      sheet1.getCell('B32').value = responsable_service;
      sheet1.getCell('B33').value = contact;
    }

    // 2. ONGLET : Publication VM
    const sheet2 = workbook.getWorksheet('Publication VM') || workbook.getWorksheet(2);
    if (sheet2) {
      const dateStr = new Date().toLocaleDateString('fr-FR');
      sheet2.getCell('D8').value = `Date : ${dateStr}`;
      
      sheet2.getCell('E14').value = population;
      sheet2.getCell('E15').value = appName;
      sheet2.getCell('E17').value = ipAddress;
      sheet2.getCell('E18').value = port;
      sheet2.getCell('F20').value = os;

      if (Array.isArray(softwareStack)) {
        const stackMap = new Map(
          softwareStack.map(s => [(s.software_name || s.name || '').toLowerCase().trim(), s])
        );

        for (let r = 21; r <= 68; r++) {
          const softLabel = (sheet2.getCell(`D${r}`).value || '').toString().toLowerCase().trim();
          if (stackMap.has(softLabel)) {
            const item = stackMap.get(softLabel);
            const exists = item.exists !== undefined ? Boolean(item.exists) : Boolean(item.is_present);
            sheet2.getCell(`E${r}`).value = exists ? 'Oui' : 'Non';
            sheet2.getCell(`F${r}`).value = exists ? (item.version || '—') : '—';
          }
        }
      }
    }

    // 3. ONGLET : Informations liées au service
    const sheet3 = workbook.getWorksheet('Informations liées au service') || workbook.getWorksheet(3);
    if (sheet3) {
      if (architectureDesc) {
        sheet3.getCell('C23').value = architectureDesc;
      }

      if (Array.isArray(networkFlows)) {
        const startRow = 73;
        const templateRow = sheet3.getRow(startRow);

        networkFlows.forEach((f, idx) => {
          const rowNum = startRow + idx;
          const row = sheet3.getRow(rowNum);

          ['C', 'D', 'E', 'F', 'G', 'H'].forEach((col) => {
            const templateCell = templateRow.getCell(col);
            const cell = row.getCell(col);
            if (templateCell.style) {
              cell.style = JSON.parse(JSON.stringify(templateCell.style));
            }
          });

          row.getCell('C').value = f.source || '';
          row.getCell('D').value = f.destination || '';
          row.getCell('E').value = f.service || '';
          row.getCell('F').value = f.port ? String(f.port) : '';
          row.getCell('G').value = f.flow_type || f.type_flux || '';
          row.getCell('H').value = f.description || '/';

          row.commit();
        });
      }
    }

    // 4. ONGLET : Suivie des Non conformités
    const sheet4 = workbook.getWorksheet('Suivie des Non conformités') || 
                   workbook.getWorksheet('Suivi des Non conformités') || 
                   workbook.getWorksheet(4);
    if (sheet4) {
      sheet4.getCell('C6').value = dnsSiteWeb;
      sheet4.getCell('C7').value = ipPublique;
      sheet4.getCell('C8').value = ipInterne;
      sheet4.getCell('C9').value = ipVirtuelleF5;
      sheet4.getCell('C10').value = publication;
      sheet4.getCell('C11').value = new Date().toLocaleString('fr-FR');

      if (Array.isArray(securityCompliance)) {
        securityCompliance.forEach((ctrl, idx) => {
          const rowNum = 13 + idx;
          sheet4.getCell(`C${rowNum}`).value = ctrl.status || 'En attente';
          sheet4.getCell(`D${rowNum}`).value = ctrl.comments || ctrl.commentaires || '/';
        });
      }
    }

    const buffer = await workbook.xlsx.writeBuffer();

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename=Formulaire_VM_${appName || 'Export'}.xlsx`);
    res.send(Buffer.from(buffer));

  } catch (error) {
    console.error('Erreur export Excel:', error);
    res.status(500).json({ error: 'Erreur lors de la génération du fichier Excel' });
  }
});

module.exports = router;