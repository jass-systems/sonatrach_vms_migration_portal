import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';

// Palette de couleurs officielles Sonatrach TRC
const COLORS = {
  HEADER_BG: '1E293B',     // Bleu nuit / Slate-800
  HEADER_TEXT: 'FFFFFF',   // Blanc
  ORANGE_TRC: 'D97706',    // Orange TRC / Amber-600
  GREEN_BG: 'D1FAE5',      // Vert clair (Oui / Conforme)
  GREEN_TEXT: '047857',    // Vert foncé
  RED_BG: 'FEE2E2',        // Rouge clair (Non validé)
  RED_TEXT: 'B91C1C',      // Rouge foncé
  GRAY_BG: 'F8FAFC'        // Gris neutre
};

// Bordure noire fine standard
const blackThinBorder = {
  top: { style: 'thin', color: { argb: '000000' } },
  left: { style: 'thin', color: { argb: '000000' } },
  bottom: { style: 'thin', color: { argb: '000000' } },
  right: { style: 'thin', color: { argb: '000000' } }
};

const styleCell = (cell, options = {}) => {
  cell.border = blackThinBorder;
  cell.font = {
    name: 'Calibri',
    size: options.fontSize || 10,
    bold: options.bold || false,
    color: { argb: options.textColor || '000000' }
  };
  cell.alignment = {
    vertical: 'middle',
    horizontal: options.align || 'left',
    wrapText: true
  };
  if (options.bgColor) {
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: options.bgColor }
    };
  }
};

export const exportVMToExcel = async (vmData) => {
  try {
    if (!vmData) {
      alert("Données de la VM introuvables pour l'exportation.");
      return;
    }

    // =========================================================================
    // 1. NORMALISATION DES DONNÉES (Compatible Formulaire React ET PostgreSQL)
    // =========================================================================
    const pole = vmData.structureInfo?.pole || vmData.pole || vmData.structure_info?.pole || 'ALGER';
    const structure = vmData.structureInfo?.structure || vmData.structure || vmData.structure_info?.structure || 'DTI';
    const respStructure = vmData.structureInfo?.responsable_structure || vmData.responsable_structure || vmData.structure_info?.responsable_structure || '/';
    const respService = vmData.structureInfo?.responsable_service || vmData.responsable_service || vmData.structure_info?.responsable_service || 'INTRANET';
    const contact = vmData.structureInfo?.contact || vmData.contact || vmData.structure_info?.contact || '';

    const pubType = vmData.formPublication?.publication_type || vmData.publication_type || vmData.form_publication?.publication_type || 'Intranet';
    const targetPop = vmData.formPublication?.target_population || vmData.target_population || vmData.form_publication?.target_population || '';
    const appName = vmData.formPublication?.app_name || vmData.app_name || vmData.form_publication?.app_name || 'VM_APP';
    const dnsEntry = vmData.formPublication?.dns_entry || vmData.dns_entry || vmData.form_publication?.dns_entry || 'N/A';
    const ipAddr = vmData.formPublication?.ip_address || vmData.ip_address || vmData.form_publication?.ip_address || '10.0.0.1';
    const port = vmData.formPublication?.port || vmData.port || vmData.form_publication?.port || '443';
    const osServer = vmData.formPublication?.os_server || vmData.os_server || vmData.form_publication?.os_server || 'Windows Server 2022';

    const archDesc = vmData.architecture_desc || vmData.architectureDesc || '';

    const softwareStack = vmData.softwareStack || vmData.software_stack || vmData.vm_software_stack || [];
    const networkFlows = vmData.networkFlows || vmData.network_flows || [];
    const securityCompliance = vmData.securityCompliance || vmData.security_compliance || [];

    const secParams = vmData.securityParams || vmData.security_params || {};
    const dnsSiteWeb = secParams.dns_site_web || vmData.dns_entry || dnsEntry;
    const ipPublique = secParams.ip_publique || vmData.public_ip || 'N/A';
    const ipInterne = secParams.ip_interne || vmData.ip_address || ipAddr;
    const ipVirtuelleF5 = secParams.ip_virtuelle_f5 || vmData.f5_virtual_ip || 'N/A';
    const publication = secParams.publication || vmData.publication || 'DEV';
    const dateDerniereMaj = secParams.date_derniere_maj || vmData.updated_at || new Date().toISOString().slice(0, 16);

    // =========================================================================
    // 2. CHARGEMENT DU FICHIER MODÈLE EXCEL
    // =========================================================================
    let workbook = new ExcelJS.Workbook();
    let loadedFromTemplate = false;

    // Liste des URLs possibles
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
        // Continuer vers la deuxième URL
      }
    }

    if (!loadedFromTemplate) {
      console.warn("Modèle Excel non trouvé dans /public, création dynamique du document...");
    }

    // -------------------------------------------------------------------------
    // ONGLET 1 : Principale
    // -------------------------------------------------------------------------
    let sheet1 = workbook.getWorksheet('Principale') || workbook.getWorksheet(1);
    if (!sheet1) sheet1 = workbook.addWorksheet('Principale');
    sheet1.columns = [{ width: 35 }, { width: 50 }];

    sheet1.getCell('B3').value = pole;
    sheet1.getCell('B4').value = structure;
    sheet1.getCell('B5').value = respStructure;
    sheet1.getCell('B6').value = respService;
    sheet1.getCell('B7').value = contact;

    for (let r = 3; r <= 7; r++) {
      styleCell(sheet1.getCell(`A${r}`), { bold: true, bgColor: COLORS.GRAY_BG });
      styleCell(sheet1.getCell(`B${r}`), { bold: false });
    }

    // -------------------------------------------------------------------------
    // ONGLET 2 : Publication VM
    // -------------------------------------------------------------------------
    let sheet2 = workbook.getWorksheet('Publication VM') || workbook.getWorksheet(2);
    if (!sheet2) sheet2 = workbook.addWorksheet('Publication VM');
    sheet2.columns = [{ width: 38 }, { width: 22 }, { width: 25 }];

    sheet2.getCell('B2').value = pubType;
    sheet2.getCell('B3').value = targetPop;
    sheet2.getCell('B4').value = appName;
    sheet2.getCell('B5').value = dnsEntry;
    sheet2.getCell('B6').value = ipAddr;
    sheet2.getCell('B7').value = port;
    sheet2.getCell('B8').value = osServer;

    for (let r = 2; r <= 8; r++) {
      styleCell(sheet2.getCell(`A${r}`), { bold: true });
      styleCell(sheet2.getCell(`B${r}`), { bold: false });
    }

    // Software Stack
    if (Array.isArray(softwareStack) && softwareStack.length > 0) {
      softwareStack.forEach((sw, idx) => {
        const rowNum = 11 + idx;
        const row = sheet2.getRow(rowNum);

        const name = sw.software_name || sw.name || '';
        const exists = sw.exists !== undefined ? sw.exists : (sw.is_present !== undefined ? sw.is_present : false);
        const version = sw.version || '—';

        row.getCell(1).value = name;
        row.getCell(2).value = exists ? 'Oui' : 'Non';
        row.getCell(3).value = exists ? version : '—';

        styleCell(row.getCell(1), { bold: true });
        if (exists) {
          styleCell(row.getCell(2), { bold: true, bgColor: COLORS.GREEN_BG, textColor: COLORS.GREEN_TEXT, align: 'center' });
          styleCell(row.getCell(3), { bold: true, bgColor: COLORS.GREEN_BG, textColor: COLORS.GREEN_TEXT, align: 'center' });
        } else {
          styleCell(row.getCell(2), { align: 'center', textColor: '64748B' });
          styleCell(row.getCell(3), { align: 'center', textColor: '94A3B8' });
        }
      });
    }

    // -------------------------------------------------------------------------
    // ONGLET 3 : Informations liées au service
    // -------------------------------------------------------------------------
    let sheet3 = workbook.getWorksheet('Informations liées au service') || workbook.getWorksheet(3);
    if (!sheet3) sheet3 = workbook.addWorksheet('Informations liées au service');
    sheet3.columns = [{ width: 25 }, { width: 20 }, { width: 15 }, { width: 12 }, { width: 25 }, { width: 35 }];

    if (archDesc) {
      const archCell = sheet3.getCell('A2');
      archCell.value = archDesc;
      styleCell(archCell, { fontSize: 10 });
    }

    if (Array.isArray(networkFlows) && networkFlows.length > 0) {
      networkFlows.forEach((flow, idx) => {
        const rowNum = 5 + idx;
        const row = sheet3.getRow(rowNum);

        row.getCell(1).value = flow.source || '';
        row.getCell(2).value = flow.destination || '';
        row.getCell(3).value = flow.service || '';
        row.getCell(4).value = flow.port || '';
        row.getCell(5).value = flow.flow_type || '';
        row.getCell(6).value = flow.description || '/';

        styleCell(row.getCell(1), { align: 'center' });
        styleCell(row.getCell(2), { align: 'center' });
        styleCell(row.getCell(3), { align: 'center' });
        styleCell(row.getCell(4), { align: 'center', bold: true });
        styleCell(row.getCell(5), { align: 'left' });
        styleCell(row.getCell(6), { align: 'left' });
      });
    }

    // -------------------------------------------------------------------------
    // ONGLET 4 : Suivi des Non conformités
    // -------------------------------------------------------------------------
    let sheet4 = workbook.getWorksheet('Suivie des Non conformités') || 
                   workbook.getWorksheet('Suivi des Non conformités') || 
                   workbook.getWorksheet(4);
    if (!sheet4) sheet4 = workbook.addWorksheet('Suivi des Non conformités');
    sheet4.columns = [{ width: 60 }, { width: 30 }, { width: 35 }];

    sheet4.getCell('B2').value = dnsSiteWeb;
    sheet4.getCell('B3').value = ipPublique;
    sheet4.getCell('B4').value = ipInterne;
    sheet4.getCell('B5').value = ipVirtuelleF5;
    sheet4.getCell('B6').value = publication;
    sheet4.getCell('B7').value = dateDerniereMaj;

    for (let r = 2; r <= 7; r++) {
      styleCell(sheet4.getCell(`A${r}`), { bold: true });
      styleCell(sheet4.getCell(`B${r}`), { bold: false });
    }

    if (Array.isArray(securityCompliance) && securityCompliance.length > 0) {
      securityCompliance.forEach((ctrl, idx) => {
        const rowNum = 10 + idx;
        const row = sheet4.getRow(rowNum);

        const name = ctrl.control_name || ctrl.name || '';
        const status = ctrl.status || 'En attente';
        const comments = ctrl.comments || '/';

        row.getCell(1).value = name;
        row.getCell(2).value = status;
        row.getCell(3).value = comments;

        styleCell(row.getCell(1), { bold: true });
        styleCell(row.getCell(3), { align: 'left' });

        if (status.startsWith('Conforme')) {
          styleCell(row.getCell(2), { bold: true, bgColor: COLORS.GREEN_BG, textColor: COLORS.GREEN_TEXT, align: 'center' });
        } else if (status === 'Non validé' || status === 'Non Conforme') {
          styleCell(row.getCell(2), { bold: true, bgColor: COLORS.RED_BG, textColor: COLORS.RED_TEXT, align: 'center' });
        } else {
          styleCell(row.getCell(2), { align: 'center', textColor: '475569' });
        }
      });
    }

    // =========================================================================
    // 3. TÉLÉCHARGEMENT DU FICHIER
    // =========================================================================
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const filename = `formulaire_création_VM_${ipAddr || appName || 'export'}.xlsx`;
    saveAs(blob, filename);

  } catch (error) {
    console.error("Erreur détaillée lors de l'exportation Excel :", error);
    alert(`Impossible d'exporter le fichier Excel : ${error.message}`);
  }
};