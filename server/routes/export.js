const express = require('express');
const router = express.Router();
const ExcelJS = require('exceljs');
const path = require('path');

router.post('/export-formulaire', async (req, res) => {
  try {
    const data = req.body;
    const templatePath = path.join(__dirname, '../formulaire_création_VM.xlsx');

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(templatePath);

    // ==========================================
    // 1. ONGLET : Principale
    // ==========================================
    const sheet1 = workbook.getWorksheet('Principale') || workbook.getWorksheet(1);
    if (sheet1) {
      // Écriture uniquement dans la colonne B (les libellés en A29:A33 ne sont pas touchés)
      sheet1.getCell('B29').value = data.pole || '';
      sheet1.getCell('B30').value = data.structure || '';
      sheet1.getCell('B31').value = data.responsable_structure || '/';
      sheet1.getCell('B32').value = data.responsable_service || '/';
      sheet1.getCell('B33').value = data.contact || '';
    }

    // ==========================================
    // 2. ONGLET : Publication VM
    // ==========================================
    const sheet2 = workbook.getWorksheet('Publication VM') || workbook.getWorksheet(2);
    if (sheet2) {
      const dateStr = new Date().toLocaleDateString('fr-FR');
      sheet2.getCell('D8').value = `Date : ${dateStr}`;
      
      // Champs de paramètres dans les colonnes E et F
      sheet2.getCell('E14').value = data.population_exploitante || '';
      sheet2.getCell('E15').value = data.nom_application || '';
      sheet2.getCell('E17').value = data.ip_address || '';
      sheet2.getCell('E18').value = data.port || '';
      sheet2.getCell('F20').value = data.os || '';

      // Software Stack : recherche dynamique par nom dans la colonne D
      if (Array.isArray(data.softwareStack)) {
        const stackMap = new Map(
          data.softwareStack.map(s => [ (s.software_name || s.name || '').toLowerCase().trim(), s ])
        );

        for (let r = 21; r <= 68; r++) {
          const softLabel = (sheet2.getCell(`D${r}`).value || '').toString().toLowerCase().trim();
          if (stackMap.has(softLabel)) {
            const item = stackMap.get(softLabel);
            const exists = item.exists !== undefined ? item.exists : Boolean(item.is_present);
            sheet2.getCell(`E${r}`).value = exists ? 'Oui' : 'Non';
            sheet2.getCell(`F${r}`).value = exists ? (item.version || '—') : '—';
          }
        }
      }
    }

    // ==========================================
    // 3. ONGLET : Informations liées au service
    // ==========================================
    const sheet3 = workbook.getWorksheet('Informations liées au service') || workbook.getWorksheet(3);
    if (sheet3) {
      if (data.architecture_desc) {
        sheet3.getCell('C23').value = data.architecture_desc;
      }

      // Matrice des flux (Début à la colonne C, ligne 73)
      if (Array.isArray(data.flux)) {
        const startRow = 73;
        const templateRow = sheet3.getRow(startRow);

        data.flux.forEach((f, idx) => {
          const rowNum = startRow + idx;
          const row = sheet3.getRow(rowNum);

          // Copie du style original pour conserver les bordures et couleurs
          ['C', 'D', 'E', 'F', 'G', 'H'].forEach((col) => {
            const templateCell = templateRow.getCell(col);
            const cell = row.getCell(col);
            if (templateCell.style) {
              cell.style = JSON.parse(JSON.stringify(templateCell.style));
            }
          });

          // Placement exact des données de la matrice
          row.getCell('C').value = f.source || '';
          row.getCell('D').value = f.destination || '';
          row.getCell('E').value = f.service || '';
          row.getCell('F').value = f.port || '';
          row.getCell('G').value = f.type_flux || '';
          row.getCell('H').value = f.description || '/';

          row.commit();
        });
      }
    }

    // ==========================================
    // 4. ONGLET : Suivie des Non conformités
    // ==========================================
    const sheet4 = workbook.getWorksheet('Suivie des Non conformités') || 
                   workbook.getWorksheet('Suivi des Non conformités') || 
                   workbook.getWorksheet(4);
    if (sheet4) {
      sheet4.getCell('C6').value = data.dns_site_web || 'N/A';
      sheet4.getCell('C7').value = data.ip_publique || 'N/A';
      sheet4.getCell('C8').value = data.ip_interne || data.ip_address || '';
      sheet4.getCell('C9').value = data.ip_virtuelle_f5 || 'N/A';
      sheet4.getCell('C10').value = data.publication || 'DEV';
      sheet4.getCell('C11').value = new Date().toLocaleString('fr-FR');

      if (Array.isArray(data.securityCompliance)) {
        data.securityCompliance.forEach((ctrl, idx) => {
          const rowNum = 13 + idx;
          sheet4.getCell(`C${rowNum}`).value = ctrl.status || 'En attente';
          sheet4.getCell(`D${rowNum}`).value = ctrl.comments || '/';
        });
      }
    }

    // Génération et envoi du buffer
    const buffer = await workbook.xlsx.writeBuffer();

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename=Formulaire_VM.xlsx');
    res.send(Buffer.from(buffer));

  } catch (error) {
    console.error('Erreur export Excel:', error);
    res.status(500).json({ error: 'Erreur lors de la génération du fichier Excel' });
  }
});

module.exports = router;