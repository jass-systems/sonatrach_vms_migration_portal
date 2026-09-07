import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';

const COLORS = {
  HEADER_BG: '1E293B',
  HEADER_TEXT: 'FFFFFF',
  ORANGE_TRC: 'D97706',
  GREEN_BG: 'D1FAE5',
  GREEN_TEXT: '047857',
  RED_BG: 'FEE2E2',
  RED_TEXT: 'B91C1C',
  GRAY_BG: 'F8FAFC'
};

const thinBorder = {
  top: { style: 'thin', color: { argb: 'CBD5E1' } },
  left: { style: 'thin', color: { argb: 'CBD5E1' } },
  bottom: { style: 'thin', color: { argb: 'CBD5E1' } },
  right: { style: 'thin', color: { argb: 'CBD5E1' } }
};

export const exportVMToExcel = async (rawVmData) => {
  try {
    if (!rawVmData) {
      alert("Données de la VM introuvables pour l'exportation.");
      return;
    }

    let vmData = rawVmData;
    const vmId = rawVmData?.id || rawVmData?._id || rawVmData?.vm_id;

    // Fetch full VM details from backend if sub-lists are missing
    if (vmId && (!rawVmData.networkFlows?.length || !rawVmData.softwareStack?.length || !rawVmData.securityCompliance?.length)) {
      try {
        const res = await fetch(`http://localhost:5000/api/vms/${vmId}`);
        if (res.ok) {
          const fetchedData = await res.json();
          vmData = { ...rawVmData, ...fetchedData };
        }
      } catch (err) {
        console.warn("Impossible de récupérer les détails complets de la VM depuis l'API, utilisation des données locales.", err);
      }
    }

    // Data Normalization
    const pole = vmData.structureInfo?.pole || vmData.pole || 'ALGER';
    const structure = vmData.structureInfo?.structure || vmData.structure || 'DTI';
    const respStructure = vmData.structureInfo?.responsable_structure || vmData.responsable_structure || '/';
    const respService = vmData.structureInfo?.responsable_service || vmData.responsable_service || 'INTRANET';
    const contact = vmData.structureInfo?.contact || vmData.contact || '';

    const pubType = vmData.formPublication?.publication_type || vmData.publication_type || 'Intranet';
    const targetPop = vmData.formPublication?.target_population || vmData.target_population || vmData.population_exploitante || '';
    const appName = vmData.formPublication?.app_name || vmData.app_name || vmData.nom_application || 'VM_APP';
    const dnsEntry = vmData.formPublication?.dns_entry || vmData.dns_entry || vmData.dns_site_web || 'N/A';
    const ipAddr = vmData.formPublication?.ip_address || vmData.ip_address || vmData.ip_interne || '10.0.0.1';
    const port = vmData.formPublication?.port || vmData.port || '443';
    const osServer = vmData.formPublication?.os_server || vmData.os_server || vmData.os || 'Windows Server 2022';

    const archDesc = vmData.architecture_desc || vmData.architectureDesc || '';

    const softwareStack = vmData.softwareStack || vmData.software_stack || [];
    const networkFlows = vmData.networkFlows || vmData.network_flows || [];
    const securityCompliance = vmData.securityCompliance || vmData.security_compliance || [];

    const secParams = vmData.securityParams || vmData.security_params || {};
    const dnsSiteWeb = secParams.dns_site_web || dnsEntry;
    const ipPublique = secParams.ip_publique || 'N/A';
    const ipInterne = secParams.ip_interne || ipAddr;
    const ipVirtuelleF5 = secParams.ip_virtuelle_f5 || 'N/A';
    const publication = secParams.publication || 'DEV';
    const dateDerniereMaj = secParams.date_derniere_maj || new Date().toISOString().slice(0, 16);

    const workbook = new ExcelJS.Workbook();
    let loadedFromTemplate = false;

    const urlsToTry = [
      '/formulaire_creation_VM.xlsx',
      '/formulaire_création_VM.xlsx',
      encodeURI('/formulaire_création_VM.xlsx')
    ];

    for (const url of urlsToTry) {
      try {
        const res = await fetch(url);
        const contentType = res.headers.get('content-type') || '';
        if (res.ok && !contentType.includes('text/html')) {
          const buffer = await res.arrayBuffer();
          await workbook.xlsx.load(buffer);
          loadedFromTemplate = true;
          break;
        }
      } catch (e) {
        // Continue trying next URL
      }
    }

    if (loadedFromTemplate) {
      // -----------------------------------------------------------------------
      // ONGLET 1 : Principale
      // -----------------------------------------------------------------------
      const sheet1 = workbook.getWorksheet('Principale') || workbook.getWorksheet(1);
      if (sheet1) {
        sheet1.getCell('B29').value = pole;
        sheet1.getCell('B30').value = structure;
        sheet1.getCell('B31').value = respStructure;
        sheet1.getCell('B32').value = respService;
        sheet1.getCell('B33').value = contact;
      }

      // -----------------------------------------------------------------------
      // ONGLET 2 : Publication VM
      // -----------------------------------------------------------------------
      const sheet2 = workbook.getWorksheet('Publication VM') || workbook.getWorksheet(2);
      if (sheet2) {
        sheet2.getCell('E13').value = pubType;
        sheet2.getCell('F13').value = targetPop;
        sheet2.getCell('F14').value = appName;
        sheet2.getCell('E15').value = dnsEntry;
        sheet2.getCell('E16').value = ipAddr;
        sheet2.getCell('E17').value = port;
        sheet2.getCell('F20').value = osServer;

        if (Array.isArray(softwareStack) && softwareStack.length > 0) {
          const stackMap = new Map();
          softwareStack.forEach(item => {
            const name = (item.software_name || item.name || '').trim().toLowerCase();
            if (name) stackMap.set(name, item);
          });

          for (let r = 20; r <= sheet2.rowCount; r++) {
            const cellD = sheet2.getCell(`D${r}`).value;
            const softName = cellD ? String(cellD).trim().toLowerCase() : '';

            if (stackMap.has(softName)) {
              const item = stackMap.get(softName);
              const isPresent = item.exists !== undefined ? item.exists : Boolean(item.is_present);
              sheet2.getCell(`E${r}`).value = isPresent ? 'Oui' : 'Non';
              sheet2.getCell(`F${r}`).value = isPresent ? (item.version || '—') : '—';
            }
          }
        }
      }

      // -----------------------------------------------------------------------
      // ONGLET 3 : Informations liées au service
      // -----------------------------------------------------------------------
      const sheet3 = workbook.getWorksheet('Informations liées au service') || workbook.getWorksheet(3);
      if (sheet3) {
        if (archDesc) sheet3.getCell('A2').value = archDesc;

        if (Array.isArray(networkFlows) && networkFlows.length > 0) {
          const startRow = 48; 
          networkFlows.forEach((flow, idx) => {
            const rowNum = startRow + idx;
            sheet3.getCell(`B${rowNum}`).value = flow.source || '';
            sheet3.getCell(`C${rowNum}`).value = flow.destination || '';
            sheet3.getCell(`D${rowNum}`).value = flow.service || '';
            sheet3.getCell(`E${rowNum}`).value = flow.port || '';
            sheet3.getCell(`F${rowNum}`).value = flow.flow_type || flow.type_flux || '';
            sheet3.getCell(`G${rowNum}`).value = flow.description || '/';
          });
        }
      }

      // -----------------------------------------------------------------------
      // ONGLET 4 : Suivi des Non conformités
      // -----------------------------------------------------------------------
      const sheet4 = workbook.getWorksheet('Suivie des Non conformités') || 
                     workbook.getWorksheet('Suivi des Non conformités') || 
                     workbook.getWorksheet(4);
      if (sheet4) {
        sheet4.getCell('C2').value = dnsSiteWeb;
        sheet4.getCell('C3').value = ipPublique;
        sheet4.getCell('C4').value = ipInterne;
        sheet4.getCell('C6').value = publication;
        sheet4.getCell('C7').value = dateDerniereMaj;

        if (Array.isArray(securityCompliance) && securityCompliance.length > 0) {
          securityCompliance.forEach((ctrl, idx) => {
            const rowNum = 10 + idx;
            const status = ctrl.status || 'En attente';
            const comments = ctrl.comments || ctrl.commentaires || '/';

            sheet4.getCell(`B${rowNum}`).value = status;
            sheet4.getCell(`C${rowNum}`).value = comments;

            const statusCell = sheet4.getCell(`B${rowNum}`);
            if (status.startsWith('Conforme')) {
              statusCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.GREEN_BG } };
              statusCell.font = { name: 'Calibri', size: 10, bold: true, color: { argb: COLORS.GREEN_TEXT } };
            } else if (status === 'Non validé' || status === 'Non Conforme') {
              statusCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.RED_BG } };
              statusCell.font = { name: 'Calibri', size: 10, bold: true, color: { argb: COLORS.RED_TEXT } };
            }
          });
        }
      }
    } else {
      // MODE 2: Dynamic workbook construction fallback
      workbook.creator = 'Sonatrach TRC';
      workbook.created = new Date();

      // ONGLET 1 : Principale
      const sheet1 = workbook.addWorksheet('Principale');
      sheet1.columns = [{ width: 35 }, { width: 50 }];

      const titleRow1 = sheet1.addRow(['Formulaire Technique de Publication et de Mise à Disposition des Ressources']);
      titleRow1.font = { bold: true, color: { argb: COLORS.HEADER_TEXT }, size: 11 };
      titleRow1.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.HEADER_BG } };
      sheet1.mergeCells('A1:B1');
      titleRow1.alignment = { vertical: 'middle', horizontal: 'center' };

      sheet1.addRow([]);

      const infoData1 = [
        ['Pôle (*)', pole],
        ['Structure (*)', structure],
        ['Responsable de la Structure (*)', respStructure],
        ['Responsable du service à publier (*)', respService],
        ['Contact (*)', contact]
      ];

      infoData1.forEach(([label, val]) => {
        const r = sheet1.addRow([label, val]);
        r.getCell(1).font = { bold: true };
        r.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.GRAY_BG } };
        r.getCell(1).border = thinBorder;
        r.getCell(2).border = thinBorder;
      });

      // ONGLET 2 : Publication VM
      const sheet2 = workbook.addWorksheet('Publication VM');
      sheet2.columns = [{ width: 38 }, { width: 22 }, { width: 25 }];

      const h2_1 = sheet2.addRow(['Formulaire Technique de Publication', '', '']);
      h2_1.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.HEADER_BG } };
      h2_1.font = { bold: true, color: { argb: COLORS.HEADER_TEXT } };
      sheet2.mergeCells('A1:C1');

      const pubData = [
        ['Type de publication', pubType],
        ['Population exploitant service', targetPop],
        ['Nom de l\'application (VM)', appName],
        ['Entrée DNS', dnsEntry],
        ['Adresse IP', ipAddr],
        ['Port', port],
        ['OS Serveur', osServer]
      ];

      pubData.forEach(([label, val]) => {
        const r = sheet2.addRow([label, val, '']);
        r.getCell(1).font = { bold: true };
        r.getCell(1).border = thinBorder;
        r.getCell(2).border = thinBorder;
      });

      sheet2.addRow([]);

      const stackHeader = sheet2.addRow(['Configuration Software', 'Existance', 'Version']);
      stackHeader.eachCell((cell) => {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.ORANGE_TRC } };
        cell.font = { bold: true, color: { argb: COLORS.HEADER_TEXT } };
        cell.border = thinBorder;
      });

      softwareStack.forEach((sw) => {
        const isPresent = sw.exists !== undefined ? sw.exists : sw.is_present;
        const existsText = isPresent ? 'Oui' : 'Non';
        const r = sheet2.addRow([sw.software_name || sw.name, existsText, sw.version || '—']);
        
        r.getCell(1).border = thinBorder;
        r.getCell(2).border = thinBorder;
        r.getCell(3).border = thinBorder;

        if (isPresent) {
          r.getCell(2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.GREEN_BG } };
          r.getCell(2).font = { bold: true, color: { argb: COLORS.GREEN_TEXT } };
        }
      });

      // ONGLET 3 : Informations liées au service
      const sheet3 = workbook.addWorksheet('Informations liées au service');
      sheet3.columns = [{ width: 25 }, { width: 20 }, { width: 15 }, { width: 12 }, { width: 25 }, { width: 35 }];

      const archHeader = sheet3.addRow(['Architecture du service', '', '', '', '', '']);
      archHeader.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.HEADER_BG } };
      archHeader.font = { bold: true, color: { argb: COLORS.HEADER_TEXT } };
      sheet3.mergeCells('A1:F1');

      const archDescRow = sheet3.addRow([archDesc]);
      sheet3.mergeCells('A2:F2');
      archDescRow.getCell(1).border = thinBorder;

      sheet3.addRow([]);

      const flowHeader = sheet3.addRow(['Source', 'Destination', 'Service', 'Ports', 'Type de flux', 'Description']);
      flowHeader.eachCell((cell) => {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.ORANGE_TRC } };
        cell.font = { bold: true, color: { argb: COLORS.HEADER_TEXT } };
        cell.border = thinBorder;
        cell.alignment = { horizontal: 'center' };
      });

      networkFlows.forEach((flow) => {
        const r = sheet3.addRow([
          flow.source,
          flow.destination,
          flow.service,
          flow.port,
          flow.flow_type || flow.type_flux,
          flow.description || '/'
        ]);
        r.eachCell((cell) => { cell.border = thinBorder; });
      });

      // ONGLET 4 : Suivi des Non conformités
      const sheet4 = workbook.addWorksheet('Suivie des Non conformités');
      sheet4.columns = [{ width: 60 }, { width: 30 }, { width: 35 }];

      const secHeader = sheet4.addRow(['Service Publié (Paramètres de Sécurité SI)', '', '']);
      secHeader.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.HEADER_BG } };
      secHeader.font = { bold: true, color: { argb: COLORS.HEADER_TEXT } };
      sheet4.mergeCells('A1:C1');

      const secParamsList = [
        ['Entrée DNS (site web)', dnsSiteWeb],
        ['Adresse IP Publique', ipPublique],
        ['Adresse IP Interne (*)', ipInterne],
        ['Adresse IP Virtuelle F5', ipVirtuelleF5],
        ['Publication (*)', publication],
        ['Date / Heure de la dernière Mise à jour', dateDerniereMaj]
      ];

      secParamsList.forEach(([k, v]) => {
        const r = sheet4.addRow([k, v, '']);
        r.getCell(1).font = { bold: true };
        r.getCell(1).border = thinBorder;
        r.getCell(2).border = thinBorder;
      });

      sheet4.addRow([]);

      const ctrlHeader = sheet4.addRow(['Contrôles Sécurité SI', 'État et Conformité', 'Commentaires']);
      ctrlHeader.eachCell((cell) => {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.ORANGE_TRC } };
        cell.font = { bold: true, color: { argb: COLORS.HEADER_TEXT } };
        cell.border = thinBorder;
      });

      securityCompliance.forEach((ctrl) => {
        const status = ctrl.status || 'En attente';
        const r = sheet4.addRow([ctrl.control_name, status, ctrl.comments || ctrl.commentaires || '/']);
        r.getCell(1).border = thinBorder;
        r.getCell(2).border = thinBorder;
        r.getCell(3).border = thinBorder;

        if (status.startsWith('Conforme')) {
          r.getCell(2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.GREEN_BG } };
          r.getCell(2).font = { bold: true, color: { argb: COLORS.GREEN_TEXT } };
        } else if (status === 'Non validé' || status === 'Non Conforme') {
          r.getCell(2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.RED_BG } };
          r.getCell(2).font = { bold: true, color: { argb: COLORS.RED_TEXT } };
        }
      });
    }

    // Generate output file
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const filename = `formulaire_création_VM_${appName || ipAddr || 'export'}.xlsx`;
    saveAs(blob, filename);

  } catch (error) {
    console.error("Erreur lors de l'exportation Excel :", error);
    alert(`Impossible d'exporter le fichier Excel : ${error.message}`);
  }
};