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
  YELLOW_BG: 'FEF3C7',
  YELLOW_TEXT: '92400E',
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

    if (vmId && (!rawVmData.networkFlows?.length || !rawVmData.softwareStack?.length || !rawVmData.securityCompliance?.length)) {
      try {
        const res = await fetch(`http://localhost:5000/api/vms/${vmId}`);
        if (res.ok) {
          const fetchedData = await res.json();
          vmData = { ...rawVmData, ...fetchedData };
        }
      } catch (err) {
        console.warn("Impossible de récupérer les détails complets depuis l'API, utilisation des données locales.", err);
      }
    }

    // Normalized Variables
    const pole = vmData.structureInfo?.pole || vmData.pole || 'ALGER';
    const structure = vmData.structureInfo?.structure || vmData.structure || 'TRC Siège / EXP';
    const respStructure = vmData.structureInfo?.responsable_structure || vmData.responsable_structure || '/';
    const respService = vmData.structureInfo?.responsable_service || vmData.responsable_service || 'INTRANET';
    const contact = vmData.structureInfo?.contact || vmData.contact || 'Berkat Siham';

    const pubType = vmData.formPublication?.publication_type || vmData.publication_type || 'Intranet';
    const targetPop = vmData.formPublication?.target_population || vmData.target_population || "Agents de la Direction EXP ainsi que les agents d'exploitation des 9 Directions Régionales.";
    const appName = vmData.formPublication?.app_name || vmData.app_name || vmData.nom_application || 'vm_app_dev';
    const dnsEntry = vmData.formPublication?.dns_entry || vmData.dns_entry || vmData.dns_site_web || 'N/A';
    const ipAddr = vmData.formPublication?.ip_address || vmData.ip_address || vmData.ip_interne || '10.118.100.64';
    const port = vmData.formPublication?.port || vmData.port || '443';
    const osServer = vmData.formPublication?.os_server || vmData.os_server || vmData.os || 'Windows Server 2022';

    const archDesc = vmData.architecture_desc || vmData.architectureDesc || vmData.service_architecture || 'Architecture Web / App / BDD : Reverse Proxy Nginx, API Python/Node.js et base de données Oracle.';

    const softwareStack = vmData.softwareStack || vmData.software_stack || [];
    const networkFlows = vmData.networkFlows || vmData.network_flows || [];
    const securityCompliance = vmData.securityCompliance || vmData.security_compliance || vmData.securityControls || [];

    const secParams = vmData.securityParams || vmData.security_params || {};
    const dnsSiteWeb = secParams.dns_site_web || dnsEntry;
    const ipPublique = secParams.ip_publique || 'N/A';
    const ipInterne = secParams.ip_interne || ipAddr;
    const ipVirtuelle = secParams.ip_virtuelle || '/';
    const publication = secParams.publication || 'DEV';
    
    // Clean Date Formatting
    const rawDate = secParams.date_derniere_maj ? new Date(secParams.date_derniere_maj) : new Date();
    const dateDerniereMaj = !isNaN(rawDate.getTime()) 
      ? rawDate.toISOString().replace('T', ' ').substring(0, 16) 
      : (secParams.date_derniere_maj || '');

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
        // Attempt next URL
      }
    }

    if (loadedFromTemplate) {
      // PAGE 1 : Principale
      const sheet1 = workbook.getWorksheet('Principale') || workbook.getWorksheet(1);
      if (sheet1) {
        sheet1.getCell('A29').value = 'Pôle (*)';
        sheet1.getCell('A30').value = 'Structure (*)';
        sheet1.getCell('A31').value = 'Responsable de la Structure (*)';
        sheet1.getCell('A32').value = 'Responsable du service à publier (*)';
        sheet1.getCell('A33').value = 'Contact (*)';

        sheet1.getCell('B29').value = pole;
        sheet1.getCell('B30').value = structure;
        sheet1.getCell('B31').value = respStructure;
        sheet1.getCell('B32').value = respService;
        sheet1.getCell('B33').value = contact;
      }

      // PAGE 2 : Publication VM
      const sheet2 = workbook.getWorksheet('Publication VM') || workbook.getWorksheet(2);
      if (sheet2) {
        sheet2.getCell('E13').value = pubType;
        sheet2.getCell('E14').value = targetPop;
        sheet2.getCell('E15').value = appName;
        sheet2.getCell('E16').value = dnsEntry;
        sheet2.getCell('E17').value = ipAddr;
        sheet2.getCell('E18').value = port;
        sheet2.getCell('F20').value = osServer;

        if (Array.isArray(softwareStack) && softwareStack.length > 0) {
          const stackMap = new Map();
          softwareStack.forEach(item => {
            const name = (item.software_name || item.name || '').trim().toLowerCase();
            if (name) stackMap.set(name, item);
          });

          for (let r = 21; r <= sheet2.rowCount; r++) {
            const cellD = sheet2.getCell(`D${r}`).value;
            const softName = cellD ? String(cellD).trim().toLowerCase() : '';

            if (stackMap.has(softName)) {
              const item = stackMap.get(softName);
              const isPresent = item.exists !== undefined ? item.exists : (item.is_present !== undefined ? item.is_present : Boolean(item.installed));
              
              const cellE = sheet2.getCell(`E${r}`);
              const cellF = sheet2.getCell(`F${r}`);
              cellE.value = isPresent ? 'Oui' : 'Non';
              cellF.value = isPresent ? (item.version || '—') : '—';

              [cellE, cellF].forEach(c => {
                c.border = thinBorder;
                c.alignment = { horizontal: 'center', vertical: 'middle' };
              });
            }
          }
        }
      }

      // PAGE 3 : Informations liées au service
      const sheet3 = workbook.getWorksheet('Informations liées au service') || workbook.getWorksheet(3);
      if (sheet3) {
        if (archDesc) sheet3.getCell('A1').value = archDesc;

        if (Array.isArray(networkFlows) && networkFlows.length > 0) {
          const startRow = 73; // Row 72 is header; data starts at row 73
          networkFlows.forEach((flow, idx) => {
            const rowNum = startRow + idx;
            sheet3.getCell(`B${rowNum}`).value = flow.source || '';
            sheet3.getCell(`C${rowNum}`).value = flow.destination || '';
            sheet3.getCell(`D${rowNum}`).value = flow.service || flow.port_protocol || '';
            sheet3.getCell(`E${rowNum}`).value = flow.port || '';
            sheet3.getCell(`F${rowNum}`).value = flow.flow_type || flow.type_flux || '';
            sheet3.getCell(`G${rowNum}`).value = flow.description || flow.rules || '/';

            ['B', 'C', 'D', 'E', 'F', 'G'].forEach(col => {
              const cell = sheet3.getCell(`${col}${rowNum}`);
              cell.border = thinBorder;
              cell.alignment = { vertical: 'middle', wrapText: true };
            });
          });
        }
      }

      // PAGE 4 : Suivie des Non conformités
      const sheet4 = workbook.getWorksheet('Suivie des Non conformités') || 
                       workbook.getWorksheet('Suivi des Non conformités') || 
                       workbook.getWorksheet(4);
      if (sheet4) {
        // Clear misplaced residual headers
        ['C2', 'C3', 'C4'].forEach(ref => { sheet4.getCell(ref).value = ''; });

        // Map values directly against template rows 6 to 11
        sheet4.getCell('C6').value = dnsSiteWeb;
        sheet4.getCell('C7').value = ipPublique;
        sheet4.getCell('C8').value = ipInterne;
        sheet4.getCell('C9').value = ipVirtuelle;
        sheet4.getCell('C10').value = publication;
        sheet4.getCell('C11').value = dateDerniereMaj;

        if (Array.isArray(securityCompliance) && securityCompliance.length > 0) {
          const startRow = 13;
          securityCompliance.forEach((ctrl, idx) => {
            const rowNum = startRow + idx;
            const refCell = sheet4.getCell(`B${rowNum}`);
            const statusCell = sheet4.getCell(`C${rowNum}`);
            const commentCell = sheet4.getCell(`D${rowNum}`);

            const status = ctrl.status || 'En attente';
            const comments = ctrl.comments || ctrl.commentaires || ctrl.action_plan || '/';

            refCell.value = ctrl.control_name || ctrl.ref || `NC-${idx + 1}`;
            statusCell.value = status;
            commentCell.value = comments;

            const statusVal = String(status).toLowerCase();
            if (statusVal.includes('conforme')) {
              statusCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.GREEN_BG } };
              statusCell.font = { name: 'Calibri', size: 10, bold: true, color: { argb: COLORS.GREEN_TEXT } };
            } else if (statusVal.includes('non')) {
              statusCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.RED_BG } };
              statusCell.font = { name: 'Calibri', size: 10, bold: true, color: { argb: COLORS.RED_TEXT } };
            } else {
              statusCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.YELLOW_BG } };
              statusCell.font = { name: 'Calibri', size: 10, bold: true, color: { argb: COLORS.YELLOW_TEXT } };
            }

            ['B', 'C', 'D'].forEach(col => {
              const cell = sheet4.getCell(`${col}${rowNum}`);
              cell.border = thinBorder;
              cell.alignment = { vertical: 'middle', wrapText: true };
            });
          });
        }
      }
    } else {
      // Dynamic Fallback Workbook
      workbook.creator = 'Sonatrach TRC';
      workbook.created = new Date();

      // Page 1
      const sheet1 = workbook.addWorksheet('Principale');
      sheet1.columns = [{ width: 35 }, { width: 50 }];
      const titleRow1 = sheet1.addRow(['Formulaire Technique de Publication et de Mise à Disposition des Ressources']);
      titleRow1.font = { bold: true, color: { argb: COLORS.HEADER_TEXT }, size: 11 };
      titleRow1.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.HEADER_BG } };
      sheet1.mergeCells('A1:B1');
      sheet1.addRow([]);

      [
        ['Pôle (*)', pole],
        ['Structure (*)', structure],
        ['Responsable de la Structure (*)', respStructure],
        ['Responsable du service à publier (*)', respService],
        ['Contact (*)', contact]
      ].forEach(([label, val]) => {
        const r = sheet1.addRow([label, val]);
        r.getCell(1).font = { bold: true };
        r.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.GRAY_BG } };
        r.getCell(1).border = thinBorder;
        r.getCell(2).border = thinBorder;
      });

      // Page 2
      const sheet2 = workbook.addWorksheet('Publication VM');
      sheet2.columns = [{ width: 38 }, { width: 22 }, { width: 25 }];
      [
        ['Type de publication', pubType],
        ['Population exploitant service', targetPop],
        ['Nom de l\'application (VM)', appName],
        ['Entrée DNS', dnsEntry],
        ['Adresse IP', ipAddr],
        ['Port', port],
        ['OS Serveur', osServer]
      ].forEach(([label, val]) => {
        const r = sheet2.addRow([label, val, '']);
        r.getCell(1).font = { bold: true };
        r.getCell(1).border = thinBorder;
        r.getCell(2).border = thinBorder;
      });

      sheet2.addRow([]);
      const stackHeader = sheet2.addRow(['Configuration Software', 'Existance', 'Version']);
      stackHeader.eachCell(cell => {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.ORANGE_TRC } };
        cell.font = { bold: true, color: { argb: COLORS.HEADER_TEXT } };
        cell.border = thinBorder;
      });

      softwareStack.forEach(sw => {
        const isPresent = sw.exists !== undefined ? sw.exists : (sw.is_present !== undefined ? sw.is_present : sw.installed);
        const r = sheet2.addRow([sw.software_name || sw.name, isPresent ? 'Oui' : 'Non', sw.version || '—']);
        r.eachCell(cell => { cell.border = thinBorder; });
      });

      // Page 3
      const sheet3 = workbook.addWorksheet('Informations liées au service');
      sheet3.columns = [{ width: 25 }, { width: 20 }, { width: 15 }, { width: 12 }, { width: 25 }, { width: 35 }];
      const archDescRow = sheet3.addRow([archDesc]);
      sheet3.mergeCells('A1:F1');
      archDescRow.getCell(1).border = thinBorder;
      sheet3.addRow([]);

      const flowHeader = sheet3.addRow(['Source', 'Destination', 'Service', 'Ports', 'Type de flux', 'Description']);
      flowHeader.eachCell(cell => {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.ORANGE_TRC } };
        cell.font = { bold: true, color: { argb: COLORS.HEADER_TEXT } };
        cell.border = thinBorder;
      });

      networkFlows.forEach(flow => {
        const r = sheet3.addRow([
          flow.source,
          flow.destination,
          flow.service || flow.port_protocol,
          flow.port,
          flow.flow_type || flow.type_flux,
          flow.description || flow.rules || '/'
        ]);
        r.eachCell(cell => { cell.border = thinBorder; });
      });

      // Page 4
      const sheet4 = workbook.addWorksheet('Suivie des Non conformités');
      sheet4.columns = [{ width: 35 }, { width: 20 }, { width: 45 }];

      [
        ['Entrée DNS (site web)', dnsSiteWeb],
        ['Adresse IP Publique', ipPublique],
        ['Adresse IP Interne (*)', ipInterne],
        ['Adresse IP Virtuelle F5', ipVirtuelle],
        ['Publication (*)', publication],
        ['Date / Heure de la dernière Mise à jour', dateDerniereMaj]
      ].forEach(([k, v]) => {
        const r = sheet4.addRow([k, v, '']);
        r.getCell(1).font = { bold: true };
        r.getCell(1).border = thinBorder;
        r.getCell(2).border = thinBorder;
      });

      sheet4.addRow([]);

      securityCompliance.forEach(ctrl => {
        const status = ctrl.status || 'En attente';
        const r = sheet4.addRow([ctrl.control_name || ctrl.ref, status, ctrl.comments || ctrl.commentaires || ctrl.action_plan || '/']);
        r.eachCell(cell => { cell.border = thinBorder; });
      });
    }

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const filename = `formulaire_création_VM_${appName ? appName.replace(/\s+/g, '_') : 'export'}.xlsx`;
    saveAs(blob, filename);

  } catch (error) {
    console.error("Erreur lors de l'exportation Excel :", error);
    alert(`Impossible d'exporter le fichier Excel : ${error.message}`);
  }
};