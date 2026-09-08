import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';

// --- PALETTE DE COULEURS ET STYLES SONATRACH TRC ---
const COLORS = {
  PRIMARY_NAVY: '0F172A',
  HEADER_BLUE: '1E293B',
  SECTION_HEADER: '334155',
  AMBER_TRC: 'D97706',
  GRAY_BG: 'F8FAFC',
  GRAY_ALT: 'F1F5F9',
  BORDER_COLOR: 'CBD5E1',
  WHITE: 'FFFFFF',
  
  GREEN_BG: 'D1FAE5',
  GREEN_TEXT: '047857',
  RED_BG: 'FEE2E2',
  RED_TEXT: 'B91C1C',
  YELLOW_BG: 'FEF3C7',
  YELLOW_TEXT: '92400E'
};

const thinBorder = {
  top: { style: 'thin', color: { argb: COLORS.BORDER_COLOR } },
  left: { style: 'thin', color: { argb: COLORS.BORDER_COLOR } },
  bottom: { style: 'thin', color: { argb: COLORS.BORDER_COLOR } },
  right: { style: 'thin', color: { argb: COLORS.BORDER_COLOR } }
};

const DEFAULT_STACK = [
  'Apache Struts', 'Apache Tomcat', 'Apache/NCSA HTTP Server', 'ASP', 'ASP.NET',
  'BEA Systems WebLogic Server', 'CGI', 'Cisco', 'Citrix', 'Django',
  'Elasticsearch', 'Front Page Server Extensions (FPSE)', 'IBM DB2', 'IIS',
  'Java Servlets/JSP', 'Java', 'Js', 'JavaServer Faces (JSF)', 'JBoss',
  'Jetty', 'Joomla', 'jQuery', 'Lotus Domino', 'Macromedia ColdFusion',
  'Macromedia JRun', 'Microsoft SQL Server', 'Microsoft Windows', 'MongoDB',
  'MySQL', 'Node.js', 'Novell', 'Oracle', 'Outlook Web Access', 'PHP',
  'PostgreSQL', 'Proxy Servers', 'Ruby', 'SSI (Server Side Includes)',
  'Sybase/ASE', 'WebDAV', 'WordPress', 'C/C++', 'Python', 'Sharepoint',
  'XML', 'Ngnix (web server & revers proxy)', 'WSL (Ubuntu server 24.04)'
];

const loadLogoImageId = async (workbook, logoPath = '/logo.png') => {
  try {
    const response = await fetch(logoPath);
    if (!response.ok) {
      const altResponse = await fetch('/public/logo.png');
      if (!altResponse.ok) return null;
      const buffer = await altResponse.arrayBuffer();
      return workbook.addImage({ buffer, extension: 'png' });
    }
    const buffer = await response.arrayBuffer();
    return workbook.addImage({ buffer, extension: 'png' });
  } catch (err) {
    console.warn("Logo non trouvé ou impossible à charger :", err);
    return null;
  }
};

const formatDate = (dateVal) => {
  if (!dateVal) return new Date().toLocaleString('fr-FR');
  if (dateVal instanceof Date && !isNaN(dateVal.getTime())) {
    return dateVal.toLocaleString('fr-FR');
  }
  const parsed = new Date(dateVal);
  if (!isNaN(parsed.getTime())) {
    return parsed.toLocaleString('fr-FR');
  }
  return String(dateVal);
};

export const exportVMToExcel = async (rawVmData) => {
  try {
    if (!rawVmData) {
      alert("Données de la VM introuvables.");
      return;
    }

    // --- NORMALISATION DES DONNÉES ---
    const pole = rawVmData.structureInfo?.pole || rawVmData.pole || 'ALGER';
    const structure = rawVmData.structureInfo?.structure || rawVmData.structure || 'TRC Siège / EXP';
    const respStructure = rawVmData.structureInfo?.responsable_structure || rawVmData.responsable_structure || 'Benali Mohamed';
    const respService = rawVmData.structureInfo?.responsable_service || rawVmData.responsable_service || "Direction Systèmes d'Information";
    const contact = rawVmData.structureInfo?.contact || rawVmData.contact || 'Berkat Siham';

    const pubType = rawVmData.formPublication?.publication_type || rawVmData.publication_type || 'Internet';
    const targetPop = rawVmData.formPublication?.target_population || rawVmData.target_population || "Agents de la Direction EXP ainsi que les agents d'exploitation des 9 Directions Régionales.";
    const appName = rawVmData.formPublication?.app_name || rawVmData.app_name || rawVmData.nom_application || 'vm_app_dev_1';
    const dnsEntry = rawVmData.formPublication?.dns_entry || rawVmData.dns_entry || 'app-dev1.sonatrach.dz';
    const ipAddr = rawVmData.formPublication?.ip_address || rawVmData.ip_address || '10.118.100.63';
    const port = rawVmData.formPublication?.port || rawVmData.port || '80, 443';
    const osServer = rawVmData.formPublication?.os_server || rawVmData.os_server || 'Linux RHEL 8';

    const archDesc = rawVmData.architecture_desc || rawVmData.architectureDesc || 'Architecture Web / App / BDD : Reverse Proxy Nginx, API Python/Node.js et base de données Oracle.';
    const softwareStack = rawVmData.softwareStack || rawVmData.software_stack || [];
    const networkFlows = rawVmData.networkFlows || rawVmData.network_flows || [];
    const securityCompliance = rawVmData.securityCompliance || rawVmData.security_compliance || [];

    const secParams = rawVmData.securityParams || rawVmData.security_params || {};
    const dnsSiteWeb = secParams.dns_site_web || dnsEntry;
    const ipPublique = secParams.ip_publique || '197.200.x.x (DMZ)';
    const ipInterne = secParams.ip_interne || ipAddr;
    const ipVirtuelle = secParams.ip_virtuelle || secParams.ip_virtuelle_f5 || '10.118.100.64';
    const publication = secParams.publication || 'PROD';
    const dateDerniereMaj = formatDate(secParams.date_derniere_maj);

    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'SONATRACH TRC';
    workbook.created = new Date();

    const logoImageId = await loadLogoImageId(workbook, '/logo.png');

    // ==========================================
    // TAB 1: Principale
    // ==========================================
    const sheet1 = workbook.addWorksheet('Principale');
    sheet1.views = [{ showGridLines: true }];
    sheet1.columns = [
      { width: 38 }, { width: 30 }, { width: 20 }, { width: 20 }, { width: 20 }
    ];

    if (logoImageId !== null) {
      sheet1.addImage(logoImageId, {
        tl: { col: 0, row: 1 },
        ext: { width: 130, height: 45 },
        editAs: 'oneCell'
      });
    }

    sheet1.mergeCells('B2:E4');
    const headerCell1 = sheet1.getCell('B2');
    headerCell1.value = 'SONATRACH - DIRECTION TRC\nFormulaire Technique de Publication et de Mise à Disposition des Ressources';
    headerCell1.font = { name: 'Calibri', size: 14, bold: true, color: { argb: COLORS.WHITE } };
    headerCell1.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.PRIMARY_NAVY } };
    headerCell1.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };

    sheet1.mergeCells('A6:E6');
    const overviewTitle = sheet1.getCell('A6');
    overviewTitle.value = 'Résumé de la demande';
    overviewTitle.font = { bold: true, color: { argb: COLORS.WHITE }, size: 11 };
    overviewTitle.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.AMBER_TRC } };
    overviewTitle.alignment = { vertical: 'middle', indent: 1 };

    const quickFields = [
      ['Application / Service', appName],
      ['Adresse IP Interne', ipAddr],
      ['Type de Publication', pubType],
      ['Système d\'Exploitation', osServer]
    ];

    quickFields.forEach(([lbl, val], idx) => {
      const rowNum = 7 + idx;
      sheet1.getCell(`A${rowNum}`).value = lbl;
      sheet1.getCell(`A${rowNum}`).font = { bold: true };
      sheet1.getCell(`A${rowNum}`).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.GRAY_BG } };
      sheet1.getCell(`A${rowNum}`).border = thinBorder;

      sheet1.mergeCells(`B${rowNum}:E${rowNum}`);
      const valCell = sheet1.getCell(`B${rowNum}`);
      valCell.value = val;
      valCell.alignment = { vertical: 'middle' };

      ['B', 'C', 'D', 'E'].forEach(col => {
        sheet1.getCell(`${col}${rowNum}`).border = thinBorder;
      });
    });

    sheet1.mergeCells('A12:E12');
    const structHeader = sheet1.getCell('A12');
    structHeader.value = 'Informations Structure & Demandeur TRC';
    structHeader.font = { bold: true, color: { argb: COLORS.WHITE }, size: 11 };
    structHeader.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.HEADER_BLUE } };
    structHeader.alignment = { vertical: 'middle', indent: 1 };

    const structFields = [
      { label: 'Pôle (*)', value: pole, row: 13 },
      { label: 'Structure (*)', value: structure, row: 14 },
      { label: 'Responsable de la Structure (*)', value: respStructure, row: 15 },
      { label: 'Responsable du service à publier (*)', value: respService, row: 16 },
      { label: 'Contact (*)', value: contact, row: 17 }
    ];

    structFields.forEach(item => {
      sheet1.getCell(`A${item.row}`).value = item.label;
      sheet1.getCell(`A${item.row}`).font = { bold: true };
      sheet1.getCell(`A${item.row}`).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.GRAY_BG } };
      sheet1.getCell(`A${item.row}`).border = thinBorder;

      sheet1.mergeCells(`B${item.row}:E${item.row}`);
      const valCell = sheet1.getCell(`B${item.row}`);
      valCell.value = item.value;
      valCell.alignment = { vertical: 'middle' };

      ['B', 'C', 'D', 'E'].forEach(col => {
        sheet1.getCell(`${col}${item.row}`).border = thinBorder;
      });
    });

    // ==========================================
    // TAB 2: Publication VM
    // ==========================================
    const sheet2 = workbook.addWorksheet('Publication VM');
    sheet2.views = [{ showGridLines: true }];
    sheet2.columns = [
      { width: 5 }, { width: 20 }, { width: 20 }, { width: 38 }, { width: 25 }, { width: 25 }
    ];

    if (logoImageId !== null) {
      sheet2.addImage(logoImageId, {
        tl: { col: 1, row: 1 },
        ext: { width: 130, height: 45 },
        editAs: 'oneCell'
      });
    }

    sheet2.mergeCells('D2:F3');
    const headerCell2 = sheet2.getCell('D2');
    headerCell2.value = 'FICHE TECHNIQUE DE PUBLICATION DES VM';
    headerCell2.font = { name: 'Calibri', size: 13, bold: true, color: { argb: COLORS.WHITE } };
    headerCell2.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.PRIMARY_NAVY } };
    headerCell2.alignment = { horizontal: 'center', vertical: 'middle' };

    sheet2.getCell('D5').value = `Date : ${new Date().toLocaleDateString('fr-FR')}`;
    sheet2.getCell('D5').font = { bold: true, color: { argb: COLORS.AMBER_TRC } };

    const pubFormFields = [
      { cellLabel: 'D7', cellVal: 'E7', label: 'Type de publication (*)', val: pubType, isDropdown: true },
      { cellLabel: 'D8', cellVal: 'E8', label: 'Population exploitant service (*)', val: targetPop },
      { cellLabel: 'D9', cellVal: 'E9', label: 'Description du service (Nom VM) (*)', val: appName },
      { cellLabel: 'D10', cellVal: 'E10', label: 'Entrée DNS (Optionnel)', val: dnsEntry },
      { cellLabel: 'D11', cellVal: 'E11', label: 'Adresse IP (*)', val: ipAddr },
      { cellLabel: 'D12', cellVal: 'E12', label: 'Port (*)', val: port },
      { cellLabel: 'D13', cellVal: 'E13', label: 'OS Serveur (*)', val: osServer }
    ];

    pubFormFields.forEach(f => {
      sheet2.getCell(f.cellLabel).value = f.label;
      sheet2.getCell(f.cellLabel).font = { bold: true };
      sheet2.getCell(f.cellLabel).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.GRAY_BG } };
      sheet2.getCell(f.cellLabel).border = thinBorder;

      const row = f.cellVal.substring(1);
      sheet2.mergeCells(`E${row}:F${row}`);
      const valCell = sheet2.getCell(`E${row}`);
      valCell.value = f.val;
      valCell.alignment = { vertical: 'middle', wrapText: true };

      sheet2.getCell(`E${row}`).border = thinBorder;
      sheet2.getCell(`F${row}`).border = thinBorder;

      if (f.isDropdown) {
        valCell.dataValidation = {
          type: 'list',
          allowBlank: true,
          formulae: ['"Intranet,Extranet,Internet,VPN"']
        };
      }
    });

    sheet2.getCell('D16').value = 'Composants Software & Tech Stack';
    sheet2.getCell('E16').value = 'Existence';
    sheet2.getCell('F16').value = 'Version / OS';

    ['D16', 'E16', 'F16'].forEach(ref => {
      const c = sheet2.getCell(ref);
      c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.AMBER_TRC } };
      c.font = { bold: true, color: { argb: COLORS.WHITE } };
      c.border = thinBorder;
      c.alignment = { horizontal: ref === 'D16' ? 'left' : 'center', vertical: 'middle' };
    });

    const stackMap = new Map();
    softwareStack.forEach(item => {
      const name = (item.software_name || item.name || '').trim().toLowerCase();
      if (name) stackMap.set(name, item);
    });

    DEFAULT_STACK.forEach((softName, idx) => {
      const rowNum = 17 + idx;
      const key = softName.toLowerCase().trim();
      const matched = stackMap.get(key);

      let isPresent = false;
      if (matched) {
        const rawVal = matched.exists !== undefined ? matched.exists : matched.is_present;
        if (typeof rawVal === 'boolean') isPresent = rawVal;
        else if (typeof rawVal === 'string') isPresent = ['oui', 'yes', 'true', '1'].includes(rawVal.toLowerCase().trim());
        else if (typeof rawVal === 'number') isPresent = rawVal === 1;
      }

      const versionStr = matched && isPresent ? (matched.version || '—') : '—';

      const cellD = sheet2.getCell(`D${rowNum}`);
      const cellE = sheet2.getCell(`E${rowNum}`);
      const cellF = sheet2.getCell(`F${rowNum}`);

      cellD.value = softName;
      cellE.value = isPresent ? 'Oui' : 'Non';
      cellF.value = versionStr;

      cellE.dataValidation = {
        type: 'list',
        allowBlank: false,
        formulae: ['"Oui,Non"'],
        showErrorMessage: true,
        errorTitle: 'Choix invalide',
        error: 'Veuillez choisir Oui ou Non.'
      };

      [cellD, cellE, cellF].forEach(c => {
        c.border = thinBorder;
        c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: idx % 2 === 0 ? COLORS.WHITE : COLORS.GRAY_BG } };
      });

      cellE.alignment = { horizontal: 'center', vertical: 'middle' };
      cellF.alignment = { horizontal: 'center', vertical: 'middle' };
      if (isPresent) {
        cellE.font = { bold: true, color: { argb: COLORS.GREEN_TEXT } };
      }
    });

    // ==========================================
    // TAB 3: Informations liées au service
    // ==========================================
    const sheet3 = workbook.addWorksheet('Informations liées au service');
    sheet3.views = [{ showGridLines: true }];
    sheet3.columns = [
      { width: 5 }, { width: 10 }, { width: 28 }, { width: 22 },
      { width: 18 }, { width: 12 }, { width: 25 }, { width: 35 }
    ];

    sheet3.mergeCells('C2:H3');
    const headerCell3 = sheet3.getCell('C2');
    headerCell3.value = 'ARCHITECTURE DU SERVICE ET MATRICE DES FLUX';
    headerCell3.font = { name: 'Calibri', size: 13, bold: true, color: { argb: COLORS.WHITE } };
    headerCell3.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.PRIMARY_NAVY } };
    headerCell3.alignment = { horizontal: 'center', vertical: 'middle' };

    sheet3.mergeCells('C5:H5');
    const archTitle = sheet3.getCell('C5');
    archTitle.value = 'Description de l\'Architecture Technique';
    archTitle.font = { bold: true, color: { argb: COLORS.WHITE } };
    archTitle.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.HEADER_BLUE } };

    sheet3.mergeCells('C6:H9');
    const archCell = sheet3.getCell('C6');
    archCell.value = archDesc;
    archCell.alignment = { vertical: 'top', wrapText: true };
    archCell.border = thinBorder;

    sheet3.mergeCells('C11:H11');
    const flowTitle = sheet3.getCell('C11');
    flowTitle.value = 'Matrice des Flux Réseau / Firewall';
    flowTitle.font = { bold: true, color: { argb: COLORS.WHITE } };
    flowTitle.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.HEADER_BLUE } };

    const flowHeaders = [
      { col: 'C', title: 'Source (*)' },
      { col: 'D', title: 'Destination (*)' },
      { col: 'E', title: 'Service (*)' },
      { col: 'F', title: 'Ports (*)' },
      { col: 'G', title: 'Type de flux (*)' },
      { col: 'H', title: 'Description / Remarques' }
    ];

    flowHeaders.forEach(fh => {
      const cell = sheet3.getCell(`${fh.col}12`);
      cell.value = fh.title;
      cell.font = { bold: true, color: { argb: COLORS.WHITE } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.AMBER_TRC } };
      cell.border = thinBorder;
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
    });

    const startFlowRow = 13;
    const defaultFlows = [
      { source: '10.118.100.0/24', destination: ipAddr, service: 'HTTP', port: 80, flow_type: 'Flux Web (Reverse Proxy)', description: 'Accès Web HTTP Utilisateurs' },
      { source: '10.118.100.0/24', destination: ipAddr, service: 'HTTPS', port: 443, flow_type: 'Flux Web Sécurisé', description: 'Accès Web HTTPS SSL/TLS' },
      { source: ipAddr, destination: '10.10.1.50', service: 'Oracle SQL', port: 1521, flow_type: 'Flux Base de données', description: 'Connexion BDD Oracle ERP' },
      { source: ipAddr, destination: '10.10.0.10', service: 'DNS / NTP', port: '53, 123', flow_type: 'Flux Infrastructure', description: 'Résolution de noms & Synchro' }
    ];

    const flowsToRender = networkFlows.length > 0 ? networkFlows : defaultFlows;

    flowsToRender.forEach((flow, idx) => {
      const r = startFlowRow + idx;
      sheet3.getCell(`C${r}`).value = flow.source || '';
      sheet3.getCell(`D${r}`).value = flow.destination || '';
      sheet3.getCell(`E${r}`).value = flow.service || flow.port_protocol || '';

      const portVal = flow.port !== undefined ? flow.port : flow.ports;
      const portCell = sheet3.getCell(`F${r}`);
      if (typeof portVal === 'number') {
        portCell.value = portVal;
        portCell.numFmt = '0';
      } else {
        portCell.value = String(portVal || '');
      }

      sheet3.getCell(`G${r}`).value = flow.flow_type || flow.type_flux || '';
      sheet3.getCell(`H${r}`).value = flow.description || flow.rules || '/';

      ['C', 'D', 'E', 'F', 'G', 'H'].forEach(col => {
        const cell = sheet3.getCell(`${col}${r}`);
        cell.border = thinBorder;
        cell.alignment = { vertical: 'middle', wrapText: true, horizontal: (col === 'E' || col === 'F') ? 'center' : 'left' };
      });
    });

    // ==========================================
    // TAB 4: Suivi des Non conformités
    // ==========================================
    const sheet4 = workbook.addWorksheet('Suivi des Non conformités');
    sheet4.views = [{ showGridLines: true }];
    sheet4.columns = [
      { width: 8 }, { width: 50 }, { width: 30 }, { width: 45 }
    ];

    sheet4.mergeCells('B2:D3');
    const headerCell4 = sheet4.getCell('B2');
    headerCell4.value = 'SUIVI DES CONTRÔLES ET DE LA CONFORMITÉ SÉCURITÉ SI';
    headerCell4.font = { name: 'Calibri', size: 13, bold: true, color: { argb: COLORS.WHITE } };
    headerCell4.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.PRIMARY_NAVY } };
    headerCell4.alignment = { horizontal: 'center', vertical: 'middle' };

    const secParamsFields = [
      { label: 'Entrée DNS (site web)', val: dnsSiteWeb, row: 5 },
      { label: 'Adresse IP Publique', val: ipPublique, row: 6 },
      { label: 'Adresse IP Interne (*)', val: ipInterne, row: 7 },
      { label: 'Adresse IP Virtuelle F5', val: ipVirtuelle, row: 8 },
      { label: 'Publication (*)', val: publication, row: 9 },
      { label: 'Date / Heure de la dernière Mise à jour (*)', val: dateDerniereMaj, row: 10 }
    ];

    secParamsFields.forEach(p => {
      sheet4.getCell(`B${p.row}`).value = p.label;
      sheet4.getCell(`B${p.row}`).font = { bold: true };
      sheet4.getCell(`B${p.row}`).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.GRAY_BG } };
      sheet4.getCell(`B${p.row}`).border = thinBorder;

      sheet4.getCell(`C${p.row}`).value = p.val;
      sheet4.getCell(`C${p.row}`).border = thinBorder;
      sheet4.getCell(`C${p.row}`).alignment = { vertical: 'middle' };
    });

    const secHeaders = [
      { col: 'A', title: 'N°' },
      { col: 'B', title: 'Contrôle de Sécurité SI' },
      { col: 'C', title: 'Statut de Conformité' },
      { col: 'D', title: 'Commentaires / Actions' }
    ];

    secHeaders.forEach(sh => {
      const cell = sheet4.getCell(`${sh.col}12`);
      cell.value = sh.title;
      cell.font = { bold: true, color: { argb: COLORS.WHITE } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.AMBER_TRC } };
      cell.border = thinBorder;
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
    });

    const startSecRow = 13;
    const defaultControls = [
      { name: 'Vérification de la mise en place du Service à publier dans la Zone DMZ', status: 'Conforme', comments: 'Positionné en DMZ Sécurisée' },
      { name: 'Vérification de l\'installation de l\'Antivirus (avec une base de signature à jour)', status: 'Conforme (Actif & À jour)', comments: 'Agent EDR actif' },
      { name: 'Vérification de l\'utilisation de certificat TLS', status: 'Conforme (Certificat TLS actif)', comments: 'Certificat valide' },
      { name: 'Réalisation d\'un scan de vulnérabilités authentifié', status: 'Conforme', comments: 'Scan validé par RSSI' },
      { name: 'Réalisation d\'un scan de Vulnérabilités Web', status: 'Conforme', comments: 'Rapport Acunetix validé' },
      { name: 'Scan de Conformité des configurations de sécurité appliquées', status: 'Conforme', comments: 'Hardening CIS RHEL 8 appliqué' },
      { name: 'Revue de Code source de l\'application', status: 'Conforme', comments: 'Analyse SonarQube validée' },
      { name: 'Vérification de l\'application de la politique du moindre privilège', status: 'Conforme (PoLP respecté)', comments: 'Accès RBAC configurés' }
    ];

    const controlsToRender = securityCompliance.length > 0 
      ? securityCompliance 
      : defaultControls.map((c, i) => ({ id: i + 1, control_name: c.name, status: c.status, comments: c.comments }));

    controlsToRender.forEach((ctrl, idx) => {
      const r = startSecRow + idx;
      const statusStr = ctrl.status || 'Conforme';
      const commentsStr = ctrl.comments || ctrl.commentaires || '/';

      sheet4.getCell(`A${r}`).value = idx + 1;
      sheet4.getCell(`B${r}`).value = ctrl.control_name || ctrl.ref || defaultControls[idx]?.name || '';
      
      const statusCell = sheet4.getCell(`C${r}`);
      statusCell.value = statusStr;

      statusCell.dataValidation = {
        type: 'list',
        allowBlank: true,
        formulae: ['"Conforme,Non conforme,En attente,Non applicable"']
      };

      sheet4.getCell(`D${r}`).value = commentsStr;

      const statusLower = String(statusStr).toLowerCase();
      if (statusLower.includes('non') && !statusLower.includes('applicable')) {
        statusCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.RED_BG } };
        statusCell.font = { name: 'Calibri', size: 10, bold: true, color: { argb: COLORS.RED_TEXT } };
      } else if (statusLower.includes('conforme')) {
        statusCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.GREEN_BG } };
        statusCell.font = { name: 'Calibri', size: 10, bold: true, color: { argb: COLORS.GREEN_TEXT } };
      } else {
        statusCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.YELLOW_BG } };
        statusCell.font = { name: 'Calibri', size: 10, bold: true, color: { argb: COLORS.YELLOW_TEXT } };
      }

      ['A', 'B', 'C', 'D'].forEach(col => {
        const cell = sheet4.getCell(`${col}${r}`);
        cell.border = thinBorder;
        cell.alignment = { vertical: 'middle', wrapText: true, horizontal: col === 'A' || col === 'C' ? 'center' : 'left' };
      });
    });

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const filename = `formulaire_création_VM_${appName ? appName.replace(/\s+/g, '_') : 'export'}.xlsx`;
    saveAs(blob, filename);

  } catch (error) {
    console.error("Erreur lors de la génération du fichier Excel :", error);
    alert(`Impossible de générer le fichier Excel : ${error.message}`);
  }
};