import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';

const COLORS = {
  ORANGE: 'ED7D31',
  HEADER_YELLOW: 'FFF2CC',
  WHITE: 'FFFFFF',
  BORDER_BLACK: '000000',
  BORDER_GRAY: 'D9D9D9',
  CANVAS_BG: 'F2F4F8',
  TEXT_DARK: '1F2937',
  TEXT_ORANGE: 'C65911',
  TEMPLATE_BLUE: '4F81BD',
  TEMPLATE_BORDER_BLUE: '5B9BD5'
};

const thinBlackBorder = {
  top: { style: 'thin', color: { argb: COLORS.BORDER_BLACK } },
  left: { style: 'thin', color: { argb: COLORS.BORDER_BLACK } },
  bottom: { style: 'thin', color: { argb: COLORS.BORDER_BLACK } },
  right: { style: 'thin', color: { argb: COLORS.BORDER_BLACK } }
};

const DEFAULT_TECH_STACK = [
  'OS Serveur', 'Apache Struts', 'Apache Tomcat', 'Apache/NCSA HTTP Server', 'ASP', 'ASP.NET',
  'BEA Systems WebLogic Server', 'CGI', 'Cisco', 'Citrix', 'Django', 'Elasticsearch',
  'Front Page Server Extensions (FPSE)', 'IBM DB2', 'IIS', 'Java Servlets/JSP', 'Java', 'Js',
  'JavaServer Faces (JSF)', 'JBoss', 'Jetty', 'Joomla', 'jQuery', 'Lotus Domino',
  'Macromedia ColdFusion', 'Macromedia JRun', 'Microsoft SQL Server', 'Microsoft Windows',
  'MongoDB', 'MySQL', 'Node.js', 'Novell', 'Oracle', 'Outlook Web Access', 'PHP', 'PostgreSQL',
  'Proxy Servers', 'Ruby', 'SSI (Server Side Includes)', 'Sybase/ASE', 'WebDAV', 'WordPress',
  'C/C++', 'Python', 'Sharepoint', 'XML', 'Ngnix (web server & revers proxy)'
];

const NUM_AUTRES_ROWS = 5;

const loadCanvasImage = async (src) => {
  if (typeof window === 'undefined') return null;
  const pathsToTry = [src, `/public${src}`, `public${src}`, src.replace(/^\//, '')];
  for (const path of pathsToTry) {
    try {
      const img = new Image();
      img.crossOrigin = 'Anonymous';
      const loaded = await new Promise((resolve) => {
        img.onload = () => resolve(img);
        img.onerror = () => resolve(null);
        img.src = path;
      });
      if (loaded) return loaded;
    } catch {
      // Continue on fail
    }
  }
  return null;
};

const loadImageId = async (workbook, imagePath) => {
  if (!imagePath) return null;

  if (typeof imagePath === 'string' && imagePath.startsWith('data:image')) {
    try {
      const match = imagePath.match(/^data:image\/(\w+);base64,/);
      const ext = match ? (match[1] === 'jpeg' ? 'jpg' : match[1]) : 'png';
      const base64Data = imagePath.replace(/^data:image\/\w+;base64,/, '');
      return workbook.addImage({ base64: base64Data, extension: ext });
    } catch {
      return null;
    }
  }

  const cleanPath = imagePath.startsWith('/') ? imagePath : `/${imagePath}`;
  const pathCandidates = [cleanPath, `/public${cleanPath}`, `public${cleanPath}`, imagePath];

  for (const path of pathCandidates) {
    try {
      const response = await fetch(path);
      if (response.ok) {
        const buffer = await response.arrayBuffer();
        return workbook.addImage({ buffer, extension: 'png' });
      }
    } catch {
      // Continue on fail
    }
  }
  return null;
};

const styleRange = (ws, startCol, startRow, endCol, endRow, style) => {
  const colToIdx = col => {
    let idx = 0;
    for (let i = 0; i < col.length; i++) idx = idx * 26 + col.charCodeAt(i) - 64;
    return idx;
  };
  const idxToCol = idx => {
    let col = '';
    while (idx > 0) {
      let rem = (idx - 1) % 26;
      col = String.fromCharCode(65 + rem) + col;
      idx = Math.floor((idx - 1) / 26);
    }
    return col;
  };

  for (let r = startRow; r <= endRow; r++) {
    for (let c = colToIdx(startCol); c <= colToIdx(endCol); c++) {
      const cell = ws.getCell(`${idxToCol(c)}${r}`);
      if (style.border) cell.border = style.border;
      if (style.fill) cell.fill = style.fill;
      if (style.font) cell.font = style.font;
      if (style.alignment) cell.alignment = style.alignment;
    }
  }
};

const formatDate = (dateVal) => {
  if (!dateVal) return new Date().toLocaleDateString('fr-FR');
  if (dateVal instanceof Date && !isNaN(dateVal.getTime())) return dateVal.toLocaleDateString('fr-FR');
  const parsed = new Date(dateVal);
  return !isNaN(parsed.getTime()) ? parsed.toLocaleDateString('fr-FR') : String(dateVal);
};

const drawArrowHead = (ctx, x, y, angle, color = '#2B78E4', size = 10) => {
  ctx.save();
  ctx.beginPath();
  ctx.fillStyle = color;
  ctx.moveTo(x, y);
  ctx.lineTo(x - size * Math.cos(angle - Math.PI / 6), y - size * Math.sin(angle - Math.PI / 6));
  ctx.lineTo(x - size * Math.cos(angle + Math.PI / 6), y - size * Math.sin(angle + Math.PI / 6));
  ctx.closePath();
  ctx.fill();
  ctx.restore();
};

const generateExactTemplateArchitectureDiagram = async (rawVmData) => {
  if (typeof document === 'undefined') return null;

  // 1. Extraction dynamique de la matrice des flux (networkFlows)
  const flows = rawVmData.networkFlows || rawVmData.network_flows || [];

  // 2. Identification dynamique des lignes de flux selon le type/service
  const webFlow = flows.find(f => 
    (f.flow_type || f.type || '').toLowerCase().includes('web') || 
    (f.flow_type || f.type || '').toLowerCase().includes('applicatif')
  );

  const dbFlow = flows.find(f => 
    (f.flow_type || f.type || '').toLowerCase().includes('base') || 
    (f.flow_type || f.type || '').toLowerCase().includes('db') ||
    (f.flow_type || f.type || '').toLowerCase().includes('bdd') ||
    (f.service || '').toLowerCase().includes('sql') ||
    (f.service || '').toLowerCase().includes('oracle') ||
    (f.service || '').toLowerCase().includes('postgres')
  );

  const sshFlow = flows.find(f => 
    (f.service || '').toLowerCase().includes('ssh') || 
    String(f.port || f.ports) === '22'
  );

  const rdpFlow = flows.find(f => 
    (f.service || '').toLowerCase().includes('rdp') || 
    String(f.port || f.ports) === '3389'
  );

  // 3. Extraction des paramètres dynamiques avec repli sur le formulaire
  const vmName = rawVmData.vm_name || rawVmData.formPublication?.app_name || rawVmData.app_name || rawVmData.nom_application || 'vm_app_dev_1';
  
  // Flux Web (WAN -> VM)
  const httpProtocol = webFlow?.service || rawVmData.formPublication?.publication_type || rawVmData.publication_type || 'INTERNET';
  const httpPortVal = webFlow?.port || webFlow?.ports || rawVmData.formPublication?.port || rawVmData.port || '443';

  // Base de données (VM -> DB & Admin -> DB)
  const dbServiceVal = dbFlow?.service || rawVmData.db_service || rawVmData.formPublication?.db_service || 'Oracle DB';
  const dbPortVal = dbFlow?.port || dbFlow?.ports || rawVmData.db_port || rawVmData.formPublication?.db_port || '1521';

  // Administration (Admin -> VM)
  const sshPortVal = sshFlow?.port || sshFlow?.ports || rawVmData.ssh_port || '22';
  const rdpPortVal = rdpFlow?.port || rdpFlow?.ports || rawVmData.rdp_port || '3389';

  // Chargeur d'icônes
  const [imgCloud, imgDb, imgAdmin] = await Promise.all([
    loadCanvasImage('/img_0.png'),
    loadCanvasImage('/img_1.png'),
    loadCanvasImage('/img_2.png')
  ]);

  // Canvas High DPI Setup
  const canvas = document.createElement('canvas');
  const width = 1400;
  const height = 580;
  canvas.width = width * 2;
  canvas.height = height * 2;
  const ctx = canvas.getContext('2d');
  ctx.scale(2, 2);
  ctx.imageSmoothingEnabled = true;
  ctx.imageSmoothingQuality = 'high';

  // Arrière-plan
  ctx.fillStyle = '#FFFFFF';
  ctx.fillRect(0, 0, width, height);

  // Conteneur interne
  const boxX = 30, boxY = 20, boxW = 1340, boxH = 540;
  ctx.fillStyle = '#' + COLORS.CANVAS_BG;
  ctx.fillRect(boxX, boxY, boxW, boxH);
  ctx.strokeStyle = '#' + COLORS.TEMPLATE_BORDER_BLUE;
  ctx.lineWidth = 2;
  ctx.strokeRect(boxX, boxY, boxW, boxH);

  // 1. Composant Cloud (SH-WAN)
  const cloudX = 120, cloudY = 60, cloudW = 190, cloudH = 95;
  if (imgCloud) {
    ctx.drawImage(imgCloud, cloudX, cloudY, cloudW, cloudH);
  }

  // 2. Boîte VM Principale
  const appBoxX = 540, appBoxY = 230, appBoxW = 280, appBoxH = 90;
  const appCenterX = appBoxX + appBoxW / 2;

  // Flèche WAN -> VM
  ctx.beginPath();
  ctx.strokeStyle = '#2B78E4';
  ctx.lineWidth = 2.5;
  ctx.moveTo(cloudX + cloudW, cloudY + cloudH / 2);
  ctx.lineTo(appCenterX, cloudY + cloudH / 2);
  ctx.lineTo(appCenterX, appBoxY);
  ctx.stroke();
  drawArrowHead(ctx, appCenterX, appBoxY, Math.PI / 2, '#2B78E4', 10);

  // Boîte d'information WAN
  const wanBoxX = 370, wanBoxY = 80;
  ctx.fillStyle = '#FFFFFF';
  ctx.fillRect(wanBoxX, wanBoxY, 145, 40);
  ctx.strokeStyle = '#CBD5E1';
  ctx.lineWidth = 1;
  ctx.strokeRect(wanBoxX, wanBoxY, 145, 40);

  ctx.fillStyle = '#1F2937';
  ctx.font = 'bold 11px Arial, sans-serif';
  ctx.textAlign = 'left';
  ctx.fillText('Service/ Ports:', wanBoxX + 8, wanBoxY + 15);
  ctx.font = '11px Arial, sans-serif';
  ctx.fillText(`${httpProtocol.toUpperCase()}/ TCP ${httpPortVal};`, wanBoxX + 8, wanBoxY + 31);

  // Dessin du rectangle VM
  ctx.fillStyle = '#' + COLORS.TEMPLATE_BLUE;
  ctx.fillRect(appBoxX, appBoxY, appBoxW, appBoxH);

  ctx.fillStyle = '#FFFFFF';
  ctx.font = 'bold 18px Arial, sans-serif';
  ctx.textAlign = 'center';
  ctx.fillText(vmName, appBoxX + appBoxW / 2, appBoxY + 52);

  // 3. Composant Base de données
  const dbX = 1040, dbY = 225, dbW = 75, dbH = 85;

  ctx.beginPath();
  ctx.strokeStyle = '#2B78E4';
  ctx.lineWidth = 2.5;
  ctx.moveTo(appBoxX + appBoxW, appBoxY + appBoxH / 2);
  ctx.lineTo(dbX, appBoxY + appBoxH / 2);
  ctx.stroke();
  drawArrowHead(ctx, dbX, appBoxY + appBoxH / 2, 0, '#2B78E4', 10);

  // Boîte d'information VM -> DB
  const appCalloutX = 845, appCalloutY = 245;
  ctx.fillStyle = '#FFFFFF';
  ctx.fillRect(appCalloutX, appCalloutY, 160, 40);
  ctx.strokeStyle = '#CBD5E1';
  ctx.lineWidth = 1;
  ctx.strokeRect(appCalloutX, appCalloutY, 160, 40);

  ctx.fillStyle = '#1F2937';
  ctx.font = 'bold 11px Arial, sans-serif';
  ctx.textAlign = 'left';
  ctx.fillText('Service/ Ports:', appCalloutX + 8, appCalloutY + 15);
  ctx.font = '11px Arial, sans-serif';
  ctx.fillText(`${dbServiceVal}/ TCP ${dbPortVal}`, appCalloutX + 8, appCalloutY + 31);

  if (imgDb) {
    const sW = imgDb.naturalWidth || imgDb.width;
    const sH = imgDb.naturalHeight || imgDb.height;
    ctx.drawImage(imgDb, 0, 0, sW, Math.floor(sH * 0.65), dbX, dbY, dbW, dbH * 0.65);
  }

  // Libellé de la Base de données
  const dbLabelBoxW = 135, dbLabelBoxH = 26;
  const dbLabelBoxX = dbX + (dbW / 2) - (dbLabelBoxW / 2);
  const dbLabelBoxY = dbY + (dbH * 0.65) + 8;

  ctx.fillStyle = '#FFFFFF';
  ctx.fillRect(dbLabelBoxX, dbLabelBoxY, dbLabelBoxW, dbLabelBoxH);
  ctx.strokeStyle = '#CBD5E1';
  ctx.lineWidth = 1;
  ctx.strokeRect(dbLabelBoxX, dbLabelBoxY, dbLabelBoxW, dbLabelBoxH);

  ctx.fillStyle = '#1F2937';
  ctx.font = 'bold 11px Arial, sans-serif';
  ctx.textAlign = 'center';
  ctx.fillText('Base de données', dbLabelBoxX + dbLabelBoxW / 2, dbLabelBoxY + 17);

  // 4. Composant Administrateur Interne
  const adminX = 110, adminY = 360, adminW = 55, adminH = 70;
  if (imgAdmin) {
    const sW = imgAdmin.naturalWidth || imgAdmin.width;
    const sH = imgAdmin.naturalHeight || imgAdmin.height;
    ctx.drawImage(imgAdmin, 0, 0, sW, Math.floor(sH * 0.65), adminX, adminY, adminW, adminH * 0.65);
  }

  // Libellé Administrateur Interne
  const adminLabelBoxW = 150, adminLabelBoxH = 26;
  const adminLabelBoxX = adminX + (adminW / 2) - (adminLabelBoxW / 2);
  const adminLabelBoxY = adminY + (adminH * 0.65) + 8;

  ctx.fillStyle = '#FFFFFF';
  ctx.fillRect(adminLabelBoxX, adminLabelBoxY, adminLabelBoxW, adminLabelBoxH);
  ctx.strokeStyle = '#CBD5E1';
  ctx.lineWidth = 1;
  ctx.strokeRect(adminLabelBoxX, adminLabelBoxY, adminLabelBoxW, adminLabelBoxH);

  ctx.fillStyle = '#1F2937';
  ctx.font = 'bold 11px Arial, sans-serif';
  ctx.textAlign = 'center';
  ctx.fillText('Administrateur interne', adminLabelBoxX + adminLabelBoxW / 2, adminLabelBoxY + 17);

  // Connexion Admin -> VM (Ligne Orange Pointillée)
  const orangeColor = '#' + COLORS.ORANGE;
  ctx.strokeStyle = orangeColor;
  ctx.lineWidth = 2.5;
  ctx.setLineDash([6, 4]);
  ctx.beginPath();
  ctx.moveTo(adminX + adminW + 15, adminY + 20);
  ctx.lineTo(440, adminY + 20);
  ctx.lineTo(440, appBoxY + 45);
  ctx.lineTo(appBoxX, appBoxY + 45);
  ctx.stroke();

  ctx.setLineDash([]);
  drawArrowHead(ctx, appBoxX, appBoxY + 45, 0, orangeColor, 10);

  // Boîte d'information Admin (SSH / RDP)
  const adminCalloutX = 210, adminCalloutY = 360;
  ctx.fillStyle = '#FFFFFF';
  ctx.fillRect(adminCalloutX, adminCalloutY, 160, 40);
  ctx.strokeStyle = '#CBD5E1';
  ctx.lineWidth = 1;
  ctx.strokeRect(adminCalloutX, adminCalloutY, 160, 40);

  ctx.fillStyle = '#1F2937';
  ctx.font = '11px Arial, sans-serif';
  ctx.textAlign = 'left';
  ctx.fillText(`SSH/TCP ${sshPortVal} ;`, adminCalloutX + 8, adminCalloutY + 15);
  ctx.fillText(`RDP/TCP ${rdpPortVal};`, adminCalloutX + 8, adminCalloutY + 31);

  // Connexion Admin -> DB (Ligne Orange Pointillée)
  ctx.strokeStyle = orangeColor;
  ctx.lineWidth = 2.5;
  ctx.setLineDash([6, 4]);
  ctx.beginPath();

  const adminLineStartX = adminLabelBoxX + adminLabelBoxW / 2;
  const adminLineStartY = adminLabelBoxY + adminLabelBoxH;
  const dbLineTargetX = dbLabelBoxX + dbLabelBoxW / 2;
  const dbLineTargetY = dbLabelBoxY + dbLabelBoxH;

  ctx.moveTo(adminLineStartX, adminLineStartY + 2);
  ctx.lineTo(adminLineStartX, 515);
  ctx.lineTo(dbLineTargetX, 515);
  ctx.lineTo(dbLineTargetX, dbLineTargetY + 2);
  ctx.stroke();

  ctx.setLineDash([]);
  drawArrowHead(ctx, dbLineTargetX, dbLineTargetY + 2, -Math.PI / 2, orangeColor, 10);

  // Boîte d'information Admin -> DB
  const adminDbCalloutW = 180, adminDbCalloutH = 40;
  const adminDbCalloutX = (adminLineStartX + dbLineTargetX) / 2 - adminDbCalloutW / 2;
  const adminDbCalloutY = 515 - adminDbCalloutH / 2;

  ctx.fillStyle = '#FFFFFF';
  ctx.fillRect(adminDbCalloutX, adminDbCalloutY, adminDbCalloutW, adminDbCalloutH);
  ctx.strokeStyle = '#CBD5E1';
  ctx.lineWidth = 1;
  ctx.strokeRect(adminDbCalloutX, adminDbCalloutY, adminDbCalloutW, adminDbCalloutH);

  ctx.fillStyle = '#1F2937';
  ctx.font = 'bold 11px Arial, sans-serif';
  ctx.textAlign = 'left';
  ctx.fillText('Service/ Ports:', adminDbCalloutX + 8, adminDbCalloutY + 15);
  ctx.font = '11px Arial, sans-serif';
  ctx.fillText(`${dbServiceVal}/ TCP ${dbPortVal}`, adminDbCalloutX + 8, adminDbCalloutY + 31);

  return canvas.toDataURL('image/png');
};

export const exportVMToExcel = async (rawVmData) => {
  try {
    if (!rawVmData) {
      alert("Données introuvables.");
      return;
    }

    const pole = rawVmData.structureInfo?.pole || rawVmData.pole || '';
    const structure = rawVmData.structureInfo?.structure || rawVmData.structure || '';
    const respStructure = rawVmData.structureInfo?.responsable_structure || rawVmData.responsable_structure || '';
    const respService = rawVmData.structureInfo?.responsable_service || rawVmData.responsable_service || '';
    const contact = rawVmData.structureInfo?.contact || rawVmData.contact || '';

    const pubType = rawVmData.formPublication?.publication_type || rawVmData.publication_type || '';
    const targetPop = rawVmData.formPublication?.target_population || rawVmData.target_population || '';
    const appName = rawVmData.formPublication?.app_name || rawVmData.app_name || rawVmData.nom_application || '';
    const dnsEntry = rawVmData.formPublication?.dns_entry || rawVmData.dns_entry || '';
    const ipAddr = rawVmData.formPublication?.ip_address || rawVmData.ip_address || '';
    const port = rawVmData.formPublication?.port || rawVmData.port || '';

    const extractedOs = (
      rawVmData.formPublication?.os_server ||
      rawVmData.os_server ||
      rawVmData.formPublication?.os_serveur ||
      rawVmData.os_serveur ||
      rawVmData.osServeur ||
      rawVmData.os ||
      ''
    ).trim();

    const softwareStack = rawVmData.softwareStack || rawVmData.software_stack || [];
    const networkFlows = rawVmData.networkFlows || rawVmData.network_flows || [];
    const securityCompliance = rawVmData.securityCompliance || rawVmData.security_compliance || [];

    const secParams = rawVmData.securityParams || rawVmData.security_params || {};
    const dnsSiteWeb = secParams.dns_site_web || dnsEntry || '';
    const ipPublique = secParams.ip_publique || '';
    const ipInterne = secParams.ip_interne || ipAddr || '';
    const ipVirtuelle = secParams.ip_virtuelle || secParams.ip_virtuelle_f5 || '';
    const publication = secParams.publication || pubType || '';
    const dateDerniereMaj = formatDate(secParams.date_derniere_maj || new Date());

    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'SONATRACH DSI';
    workbook.created = new Date();

    const principaleImageId = await loadImageId(workbook, '/principale.png');
    const logoImageId = await loadImageId(workbook, '/logo.png');

    let archImageId = null;
    const dynamicDiagramUrl = await generateExactTemplateArchitectureDiagram(rawVmData);
    if (dynamicDiagramUrl) {
      archImageId = await loadImageId(workbook, dynamicDiagramUrl);
    }

    // Sheet 1: Principale
    const ws1 = workbook.addWorksheet('Principale');
    ws1.views = [{ showGridLines: true }];
    ws1.pageSetup = { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 1 };

    ['A', 'B', 'C', 'D', 'E'].forEach(col => ws1.getColumn(col).width = 32);
    for (let r = 2; r <= 14; r++) ws1.getRow(r).height = 30;

    ws1.mergeCells('A2:E14');
    const activePrincipaleLogo = principaleImageId ?? logoImageId;
    if (activePrincipaleLogo !== null) {
      ws1.addImage(activePrincipaleLogo, {
        tl: { col: 0, row: 1 },
        br: { col: 5, row: 14 },
        editAs: 'twoCell'
      });
    }

    const metaRows = [
      ['Pôle (*):', pole],
      ['Structure (*):', structure],
      ['Responsable de la Structure (*):', respStructure],
      ['Responsable du service à publier (*):', respService],
      ['Contact (*):', contact]
    ];

    metaRows.forEach((row, idx) => {
      const rIdx = 16 + idx;
      ws1.getRow(rIdx).height = 24;

      ws1.mergeCells(`A${rIdx}:B${rIdx}`);
      const cellLabel = ws1.getCell(`A${rIdx}`);
      cellLabel.value = row[0];
      cellLabel.font = { bold: true, size: 9.5 };
      cellLabel.alignment = { vertical: 'middle' };

      ws1.mergeCells(`C${rIdx}:E${rIdx}`);
      const cellVal = ws1.getCell(`C${rIdx}`);
      cellVal.value = row[1];
      cellVal.alignment = { vertical: 'middle' };
      if (row[0].includes('Responsable du service')) {
        cellVal.font = { bold: true, color: { argb: COLORS.TEXT_ORANGE } };
      }

      styleRange(ws1, 'A', rIdx, 'E', rIdx, { border: thinBlackBorder });
    });

    ws1.getRow(22).height = 20;
    const reqCell = ws1.getCell('A22');
    reqCell.value = '(*) obligatoire';
    reqCell.font = { bold: true, size: 10 };
    reqCell.alignment = { vertical: 'middle', horizontal: 'left' };

    // Sheet 2: Publication VM
    const ws2 = workbook.addWorksheet('Publication VM');
    ws2.views = [{ showGridLines: true }];

    ws2.getColumn('A').width = 3;
    ws2.getColumn('B').width = 3;
    ws2.getColumn('C').width = 28;
    ws2.getColumn('D').width = 38;
    ws2.getColumn('E').width = 20;
    ws2.getColumn('F').width = 22;

    for (let r = 5; r <= 8; r++) ws2.getRow(r).height = 22;

    ws2.mergeCells('C5:C8');
    const logoCell = ws2.getCell('C5');
    logoCell.value = "Direction Centrale Informatique et\nSystème d'Information";
    logoCell.font = { bold: true, size: 8.5 };
    logoCell.alignment = { horizontal: 'center', vertical: 'bottom', wrapText: true };

    const activeSheet2Logo = logoImageId ?? principaleImageId;
    if (activeSheet2Logo !== null) {
      ws2.addImage(activeSheet2Logo, {
        tl: { col: 2.60, row: 4.12 },
        ext: { width: 70, height: 70 }
      });
    }

    ws2.mergeCells('D5:F5');
    const headerTitle2 = ws2.getCell('D5');
    headerTitle2.value = 'Formulaire technique de publication';
    headerTitle2.font = { bold: true, size: 13 };
    headerTitle2.alignment = { horizontal: 'center', vertical: 'middle' };

    ws2.getCell('D6').value = 'Code document :';
    ws2.getCell('E6').value = 'Rev: 1.0';
    ws2.getCell('D7').value = 'Document de référence :';

    ws2.mergeCells('D8:E8');
    ws2.getCell('D8').value = `Date : ${formatDate(new Date())}`;

    ws2.mergeCells('F6:F8');
    const pageCell = ws2.getCell('F6');
    pageCell.value = 'Page :1/1';
    pageCell.font = { bold: true, size: 10 };
    pageCell.alignment = { horizontal: 'center', vertical: 'middle' };

    styleRange(ws2, 'C', 5, 'F', 8, { border: thinBlackBorder });

    // Section Banner
    ws2.getRow(10).height = 25;
    ws2.mergeCells('C10:F10');
    styleRange(ws2, 'C', 10, 'F', 10, {
      fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.ORANGE } },
      border: thinBlackBorder
    });
    const banner10 = ws2.getCell('C10');
    banner10.value = 'Formulaire technique de publication';
    banner10.font = { bold: true, color: { argb: COLORS.WHITE }, size: 11 };
    banner10.alignment = { horizontal: 'center', vertical: 'middle' };

    // Section: Informations
    ws2.getRow(13).height = 24;
    ws2.getRow(14).height = 32;

    ws2.mergeCells('C13:C14');
    styleRange(ws2, 'C', 13, 'C', 14, {
      fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.ORANGE } },
      border: thinBlackBorder
    });
    const infoLabel = ws2.getCell('C13');
    infoLabel.value = 'Informations';
    infoLabel.font = { bold: true, color: { argb: COLORS.WHITE }, size: 10 };
    infoLabel.alignment = { horizontal: 'center', vertical: 'middle' };

    ws2.getCell('D13').value = 'Type de publication :';
    ws2.getCell('D14').value = 'Population exploitant service à publier :';

    ws2.mergeCells('E13:F13');
    ws2.getCell('E13').value = pubType;

    ws2.mergeCells('E14:F14');
    ws2.getCell('E14').value = targetPop;

    styleRange(ws2, 'D', 13, 'F', 14, { border: thinBlackBorder });

    // Section: Description
    ws2.getRow(15).height = 45;
    styleRange(ws2, 'C', 15, 'C', 15, {
      fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.ORANGE } },
      border: thinBlackBorder
    });
    const descLabel = ws2.getCell('C15');
    descLabel.value = 'Description du service à publier';
    descLabel.font = { bold: true, color: { argb: COLORS.WHITE }, size: 10 };
    descLabel.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };

    ws2.mergeCells('D15:F15');
    ws2.getCell('D15').value = appName;

    styleRange(ws2, 'D', 15, 'F', 15, { border: thinBlackBorder });

    // Section: Paramètres
    for (let r = 16; r <= 18; r++) ws2.getRow(r).height = 22;

    ws2.mergeCells('C16:C18');
    styleRange(ws2, 'C', 16, 'C', 18, {
      fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.ORANGE } },
      border: thinBlackBorder
    });
    const paramLabel = ws2.getCell('C16');
    paramLabel.value = 'Paramètres';
    paramLabel.font = { bold: true, color: { argb: COLORS.WHITE }, size: 10 };
    paramLabel.alignment = { horizontal: 'center', vertical: 'middle' };

    ws2.getCell('D16').value = 'Entrée DNS';
    ws2.getCell('D17').value = 'Adresse IP';
    ws2.getCell('D18').value = 'Port';

    ws2.mergeCells('E16:F16');
    ws2.getCell('E16').value = dnsEntry;

    ws2.mergeCells('E17:F17');
    ws2.getCell('E17').value = ipAddr;

    ws2.mergeCells('E18:F18');
    ws2.getCell('E18').value = port;

    styleRange(ws2, 'D', 16, 'F', 18, { border: thinBlackBorder });

    // Software Stack Dynamic Mapping
    const defaultTechEndRow = 19 + DEFAULT_TECH_STACK.length;
    const autresStartRow = defaultTechEndRow + 1;
    const totalTechEndRow = autresStartRow + NUM_AUTRES_ROWS - 1;

    ws2.mergeCells(`C19:C${defaultTechEndRow}`);
    styleRange(ws2, 'C', 19, 'C', defaultTechEndRow, {
      fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.ORANGE } },
      border: thinBlackBorder
    });
    const configLabel = ws2.getCell('C19');
    configLabel.value = 'Configuration software';
    configLabel.font = { bold: true, color: { argb: COLORS.WHITE }, size: 10 };
    configLabel.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };

    ws2.mergeCells(`C${autresStartRow}:C${totalTechEndRow}`);
    styleRange(ws2, 'C', autresStartRow, 'C', totalTechEndRow, {
      fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.ORANGE } },
      border: thinBlackBorder
    });
    ws2.getCell(`C${autresStartRow}`).value = 'AUTRE';
    ws2.getCell(`C${autresStartRow}`).font = { bold: true, color: { argb: COLORS.WHITE } };
    ws2.getCell(`C${autresStartRow}`).alignment = { horizontal: 'center', vertical: 'middle' };

    ws2.getRow(19).height = 24;
    ws2.getCell('D19').value = 'Technologie';
    ws2.getCell('E19').value = 'Existance';
    ws2.getCell('F19').value = 'version';

    ['D19', 'E19', 'F19'].forEach(cellKey => {
      const cell = ws2.getCell(cellKey);
      cell.font = { bold: true, color: { argb: COLORS.TEXT_ORANGE }, size: 10 };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.HEADER_YELLOW } };
      cell.alignment = { horizontal: cellKey === 'D19' ? 'left' : 'center', vertical: 'middle' };
    });

    const stackMap = new Map();
    softwareStack.forEach(item => {
      const name = (item.software_name || item.name || '').trim();
      if (!name) return;
      const matchedDefault = DEFAULT_TECH_STACK.find(t => t.toLowerCase() === name.toLowerCase());
      if (matchedDefault) stackMap.set(matchedDefault.toLowerCase(), item);
    });

    DEFAULT_TECH_STACK.forEach((tech, idx) => {
      const rowNum = 20 + idx;
      ws2.getRow(rowNum).height = 20;

      const techCell = ws2.getCell(`D${rowNum}`);
      techCell.value = tech;
      techCell.alignment = { vertical: 'middle' };

      let isPresent = false;
      let versionVal = '—';

      if (tech.toLowerCase() === 'os serveur') {
        if (extractedOs) {
          isPresent = true;
          versionVal = extractedOs;
        } else {
          const matched = stackMap.get('os serveur');
          if (matched) {
            const rawVal = matched.exists !== undefined ? matched.exists : matched.is_present;
            isPresent = typeof rawVal === 'boolean' ? rawVal : ['oui', 'yes', 'true', '1'].includes(String(rawVal).toLowerCase().trim());
            versionVal = matched.version || matched.value || matched.software_name || '—';
          }
        }
      } else {
        const matched = stackMap.get(tech.toLowerCase());
        if (matched) {
          const rawVal = matched.exists !== undefined ? matched.exists : matched.is_present;
          isPresent = typeof rawVal === 'boolean' ? rawVal : ['oui', 'yes', 'true', '1'].includes(String(rawVal).toLowerCase().trim());
          if (isPresent) versionVal = matched.version || '—';
        }
      }

      const existCell = ws2.getCell(`E${rowNum}`);
      existCell.value = isPresent ? 'Oui' : 'Non';
      existCell.alignment = { horizontal: 'center', vertical: 'middle' };

      const verCell = ws2.getCell(`F${rowNum}`);
      verCell.value = isPresent ? versionVal : '—';
      verCell.alignment = { horizontal: 'center', vertical: 'middle' };
    });

    for (let i = 0; i < NUM_AUTRES_ROWS; i++) {
      const rowNum = autresStartRow + i;
      ws2.getRow(rowNum).height = 20;
      ['D', 'E', 'F'].forEach(col => ws2.getCell(`${col}${rowNum}`).value = '');
    }

    styleRange(ws2, 'D', 19, 'F', totalTechEndRow, { border: thinBlackBorder });

    // Sheet 3: Informations liées au service
    const ws3 = workbook.addWorksheet('Informations liées au service');
    ws3.views = [{ showGridLines: true }];

    const sheet3Cols = ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R'];
    sheet3Cols.forEach(c => ws3.getColumn(c).width = 12);

    let currentR3 = 23;

    // Banner Architecture
    ws3.getRow(currentR3).height = 25;
    ws3.mergeCells(`B${currentR3}:R${currentR3}`);
    styleRange(ws3, 'B', currentR3, 'R', currentR3, {
      fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.ORANGE } },
      border: thinBlackBorder
    });
    const schemaBanner = ws3.getCell(`B${currentR3}`);
    schemaBanner.value = "Architecture du service";
    schemaBanner.font = { bold: true, color: { argb: COLORS.WHITE }, size: 12 };
    schemaBanner.alignment = { horizontal: 'center', vertical: 'middle' };

    currentR3++;

    // Image du Schéma d'Architecture Dynamique
    if (archImageId !== null) {
      const imgStartRow = currentR3 + 1;
      const imgEndRow = currentR3 + 22;

      for (let r = imgStartRow; r <= imgEndRow; r++) ws3.getRow(r).height = 20;

      ws3.addImage(archImageId, {
        tl: { col: 1, row: imgStartRow - 1 },
        br: { col: 17, row: imgEndRow },
        editAs: 'twoCell'
      });

      currentR3 = imgEndRow + 2;
    } else {
      currentR3 += 2;
    }

    // Banner Matrice des flux
    ws3.getRow(currentR3).height = 25;
    ws3.mergeCells(`B${currentR3}:R${currentR3}`);
    styleRange(ws3, 'B', currentR3, 'R', currentR3, {
      fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.ORANGE } },
      border: thinBlackBorder
    });
    const bannerFlux = ws3.getCell(`B${currentR3}`);
    bannerFlux.value = 'Matrice des flux';
    bannerFlux.font = { bold: true, color: { argb: COLORS.WHITE }, size: 12 };
    bannerFlux.alignment = { horizontal: 'center', vertical: 'middle' };

    currentR3 += 2;

    // Flow Table Headers
    ws3.getRow(currentR3).height = 22;

    ws3.mergeCells(`B${currentR3}:D${currentR3}`);
    ws3.mergeCells(`E${currentR3}:G${currentR3}`);
    ws3.mergeCells(`H${currentR3}:I${currentR3}`);
    ws3.mergeCells(`J${currentR3}:K${currentR3}`);
    ws3.mergeCells(`L${currentR3}:N${currentR3}`);
    ws3.mergeCells(`O${currentR3}:R${currentR3}`);

    const headerConfigs = [
      { cell: `B${currentR3}`, title: 'Source' },
      { cell: `E${currentR3}`, title: 'Destination' },
      { cell: `H${currentR3}`, title: 'Service' },
      { cell: `J${currentR3}`, title: 'Ports' },
      { cell: `L${currentR3}`, title: 'Type de flux' },
      { cell: `O${currentR3}`, title: 'Description' }
    ];

    headerConfigs.forEach(h => {
      const cell = ws3.getCell(h.cell);
      cell.value = h.title;
      cell.font = { bold: true, color: { argb: COLORS.WHITE } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.ORANGE } };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
    });
    styleRange(ws3, 'B', currentR3, 'R', currentR3, { border: thinBlackBorder });

    currentR3++;

    // Flow Data Rows
    const matrixData = networkFlows.map(f => ({
      src: f.source || f.src || '',
      dest: f.destination || f.dest || '',
      service: f.service || '',
      port: f.port || f.ports || '',
      type: f.flow_type || f.type || '',
      desc: f.description || f.desc || ''
    }));

    if (matrixData.length === 0) {
      const defaultSrc = ipAddr ? `${ipAddr}/16` : '10.10.0.1/16';
      ws3.getRow(currentR3).height = 20;
      ws3.mergeCells(`B${currentR3}:D${currentR3}`);
      ws3.mergeCells(`E${currentR3}:G${currentR3}`);
      ws3.mergeCells(`H${currentR3}:I${currentR3}`);
      ws3.mergeCells(`J${currentR3}:K${currentR3}`);
      ws3.mergeCells(`L${currentR3}:N${currentR3}`);
      ws3.mergeCells(`O${currentR3}:R${currentR3}`);

      ws3.getCell(`B${currentR3}`).value = defaultSrc;
      ws3.getCell(`E${currentR3}`).value = defaultSrc;
      ws3.getCell(`H${currentR3}`).value = `TCP/${port || 80}`;
      ws3.getCell(`J${currentR3}`).value = port || '80';
      ws3.getCell(`L${currentR3}`).value = 'Flux de base de données';
      ws3.getCell(`O${currentR3}`).value = appName || 'erp';

      styleRange(ws3, 'B', currentR3, 'R', currentR3, { border: thinBlackBorder });
    } else {
      matrixData.forEach((row) => {
        ws3.getRow(currentR3).height = 20;
        ws3.mergeCells(`B${currentR3}:D${currentR3}`);
        ws3.mergeCells(`E${currentR3}:G${currentR3}`);
        ws3.mergeCells(`H${currentR3}:I${currentR3}`);
        ws3.mergeCells(`J${currentR3}:K${currentR3}`);
        ws3.mergeCells(`L${currentR3}:N${currentR3}`);
        ws3.mergeCells(`O${currentR3}:R${currentR3}`);

        ws3.getCell(`B${currentR3}`).value = row.src;
        ws3.getCell(`E${currentR3}`).value = row.dest;
        ws3.getCell(`H${currentR3}`).value = row.service;
        ws3.getCell(`J${currentR3}`).value = row.port;
        ws3.getCell(`L${currentR3}`).value = row.type;
        ws3.getCell(`O${currentR3}`).value = row.desc;

        styleRange(ws3, 'B', currentR3, 'R', currentR3, { border: thinBlackBorder });
        currentR3++;
      });
    }

    // Sheet 4: Suivi des Non conformités
    const ws4 = workbook.addWorksheet('Suivi des Non conformités');
    ws4.views = [{ showGridLines: true }];
    ws4.getColumn('B').width = 50;
    ws4.getColumn('C').width = 30;
    ws4.getColumn('D').width = 35;

    ws4.getRow(5).height = 24;
    ws4.mergeCells('B5:D5');
    styleRange(ws4, 'B', 5, 'D', 5, {
      fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.ORANGE } },
      border: thinBlackBorder
    });
    const bannerPub = ws4.getCell('B5');
    bannerPub.value = 'Service Publié';
    bannerPub.font = { bold: true, color: { argb: COLORS.WHITE } };
    bannerPub.alignment = { horizontal: 'center', vertical: 'middle' };

    const pubFields = [
      ['Entrée DNS (site web)', dnsSiteWeb],
      ['Adresse IP Publique', ipPublique],
      ['Adresse IP Interne (*)', ipInterne],
      ['Adresse IP Virtuelle F5', ipVirtuelle],
      ['Publication (*)', publication],
      ['Date / Heure de la dernière Mise à jour (*)', dateDerniereMaj]
    ];

    pubFields.forEach((item, idx) => {
      const r = 6 + idx;
      ws4.getRow(r).height = 22;
      ws4.getCell(`B${r}`).value = item[0];
      ws4.getCell(`B${r}`).font = { bold: true };
      ws4.getCell(`B${r}`).alignment = { vertical: 'middle' };

      const valCell = ws4.getCell(`C${r}`);
      valCell.value = item[1];
      valCell.alignment = { vertical: 'middle' };

      styleRange(ws4, 'B', r, 'D', r, { border: thinBlackBorder });
    });

    ws4.getRow(12).height = 24;
    const ctrlHeaders = ['Contrôles', 'Etat et Conformité', 'Commentaires'];
    ['B', 'C', 'D'].forEach((c, i) => {
      const cell = ws4.getCell(`${c}12`);
      cell.value = ctrlHeaders[i];
      cell.font = { bold: true, color: { argb: COLORS.WHITE } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.ORANGE } };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      cell.border = thinBlackBorder;
    });

    const defaultControls = [
      'Vérification de la mise en place du Service à publier dans la Zone DMZ',
      "Vérification de l'installation de l'Antivirus (avec une base de signature à jour)",
      "Vérification de l'utilisation de certificat TLS",
      "Réalisation d'un scan de vulnérabilités authentifié",
      "Réalisation d'un scan de Vulnérabilités Web",
      'Scan de Conformité des configurations de sécurité appliquées',
      "Revue de Code source de l'application",
      "Vérification de l'application de la politique du moindre privilège pour chaque type d'utilisateur ayant accès au service."
    ];

    defaultControls.forEach((ctrl, idx) => {
      const r = 13 + idx;
      ws4.getRow(r).height = 22;
      const matched = securityCompliance[idx] || {};
      const statusStr = matched.status || matched.etat || 'Conforme';
      const commentStr = matched.comment || matched.comments || '/';

      ws4.getCell(`B${r}`).value = ctrl;
      ws4.getCell(`B${r}`).alignment = { vertical: 'middle' };

      const statusCell = ws4.getCell(`C${r}`);
      statusCell.value = statusStr;
      statusCell.alignment = { horizontal: 'center', vertical: 'middle' };
      statusCell.font = { bold: true, color: { argb: COLORS.TEXT_DARK } };

      const commentCell = ws4.getCell(`D${r}`);
      commentCell.value = commentStr;
      commentCell.alignment = { vertical: 'middle' };

      styleRange(ws4, 'B', r, 'D', r, { border: thinBlackBorder });
    });

    // Sauvegarde du fichier Excel
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const safeAppName = appName ? appName.replace(/[/\\?%*:|"<>]/g, '_').replace(/\s+/g, '_') : 'export';
    saveAs(blob, `formulaire_publication_VM_${safeAppName}.xlsx`);

  } catch (error) {
    console.error("Erreur d'exportation Excel :", error);
    alert(`Erreur : ${error.message}`);
  }
};