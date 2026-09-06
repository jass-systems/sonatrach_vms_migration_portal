import React, { useState } from 'react';
import { createPortal } from 'react-dom';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';

// =========================================================================
// FONCTION D'EXPORTATION EXCEL DE LA MACHINE VIRTUELLE (INDIVIDUELLE)
// =========================================================================
const exportVMToExcel = async (data) => {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Sonatrach TRC';
  workbook.created = new Date();

  // Palette de couleurs Sonatrach TRC
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

  const thinBorder = {
    top: { style: 'thin', color: { argb: 'CBD5E1' } },
    left: { style: 'thin', color: { argb: 'CBD5E1' } },
    bottom: { style: 'thin', color: { argb: 'CBD5E1' } },
    right: { style: 'thin', color: { argb: 'CBD5E1' } }
  };

  // -------------------------------------------------------------------------
  // ONGLET 1 : Principale
  // -------------------------------------------------------------------------
  const sheet1 = workbook.addWorksheet('Principale');
  sheet1.columns = [{ width: 35 }, { width: 50 }];

  const titleRow1 = sheet1.addRow(['Formulaire Technique de Publication et de Mise à Disposition des Ressources']);
  titleRow1.font = { bold: true, color: { argb: COLORS.HEADER_TEXT }, size: 11 };
  titleRow1.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.HEADER_BG } };
  sheet1.mergeCells('A1:B1');
  titleRow1.alignment = { vertical: 'middle', horizontal: 'center' };

  sheet1.addRow([]); // Espace

  const infoData1 = [
    ['Pôle (*)', data.structureInfo?.pole || 'ALGER'],
    ['Structure (*)', data.structureInfo?.structure || 'DTI'],
    ['Responsable de la Structure (*)', data.structureInfo?.responsable_structure || '/'],
    ['Responsable du service à publier (*)', data.structureInfo?.responsable_service || 'INTRANET'],
    ['Contact (*)', data.structureInfo?.contact || '']
  ];

  infoData1.forEach(([label, val]) => {
    const r = sheet1.addRow([label, val]);
    r.getCell(1).font = { bold: true };
    r.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.GRAY_BG } };
    r.getCell(1).border = thinBorder;
    r.getCell(2).border = thinBorder;
  });

  // -------------------------------------------------------------------------
  // ONGLET 2 : Publication VM
  // -------------------------------------------------------------------------
  const sheet2 = workbook.addWorksheet('Publication VM');
  sheet2.columns = [{ width: 38 }, { width: 22 }, { width: 25 }];

  const h2_1 = sheet2.addRow(['Formulaire Technique de Publication', '', '']);
  h2_1.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.HEADER_BG } };
  h2_1.font = { bold: true, color: { argb: COLORS.HEADER_TEXT } };
  sheet2.mergeCells('A1:C1');

  const pubData = [
    ['Type de publication', data.formPublication?.publication_type || 'Intranet'],
    ['Population exploitant service', data.formPublication?.target_population || ''],
    ['Nom de l\'application (VM)', data.formPublication?.app_name || ''],
    ['Entrée DNS', data.formPublication?.dns_entry || 'N/A'],
    ['Adresse IP', data.formPublication?.ip_address || ''],
    ['Port', data.formPublication?.port || '443'],
    ['OS Serveur', data.formPublication?.os_server || '']
  ];

  pubData.forEach(([label, val]) => {
    const r = sheet2.addRow([label, val, '']);
    r.getCell(1).font = { bold: true };
    r.getCell(1).border = thinBorder;
    r.getCell(2).border = thinBorder;
  });

  sheet2.addRow([]); // Espace

  const stackHeader = sheet2.addRow(['Configuration Software', 'Existance', 'Version']);
  stackHeader.eachCell((cell) => {
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.ORANGE_TRC } };
    cell.font = { bold: true, color: { argb: COLORS.HEADER_TEXT } };
    cell.border = thinBorder;
  });

  (data.softwareStack || []).forEach((sw) => {
    const existsText = sw.exists ? 'Oui' : 'Non';
    const r = sheet2.addRow([sw.software_name, existsText, sw.version || '—']);
    
    r.getCell(1).border = thinBorder;
    r.getCell(2).border = thinBorder;
    r.getCell(3).border = thinBorder;

    if (sw.exists) {
      r.getCell(2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.GREEN_BG } };
      r.getCell(2).font = { bold: true, color: { argb: COLORS.GREEN_TEXT } };
    }
  });

  // -------------------------------------------------------------------------
  // ONGLET 3 : Informations liées au service
  // -------------------------------------------------------------------------
  const sheet3 = workbook.addWorksheet('Informations liées au service');
  sheet3.columns = [{ width: 25 }, { width: 20 }, { width: 15 }, { width: 12 }, { width: 25 }, { width: 35 }];

  const archHeader = sheet3.addRow(['Architecture du service', '', '', '', '', '']);
  archHeader.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.HEADER_BG } };
  archHeader.font = { bold: true, color: { argb: COLORS.HEADER_TEXT } };
  sheet3.mergeCells('A1:F1');

  const archDescRow = sheet3.addRow([data.architecture_desc || '']);
  sheet3.mergeCells('A2:F2');
  archDescRow.getCell(1).border = thinBorder;

  sheet3.addRow([]); // Espace

  const flowHeader = sheet3.addRow(['Source', 'Destination', 'Service', 'Ports', 'Type de flux', 'Description']);
  flowHeader.eachCell((cell) => {
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.ORANGE_TRC } };
    cell.font = { bold: true, color: { argb: COLORS.HEADER_TEXT } };
    cell.border = thinBorder;
    cell.alignment = { horizontal: 'center' };
  });

  (data.networkFlows || []).forEach((flow) => {
    const r = sheet3.addRow([
      flow.source,
      flow.destination,
      flow.service,
      flow.port,
      flow.flow_type,
      flow.description || '/'
    ]);
    r.eachCell((cell) => { cell.border = thinBorder; });
  });

  // -------------------------------------------------------------------------
  // ONGLET 4 : Suivie des Non conformités
  // -------------------------------------------------------------------------
  const sheet4 = workbook.addWorksheet('Suivie des Non conformités');
  sheet4.columns = [{ width: 60 }, { width: 30 }, { width: 35 }];

  const secHeader = sheet4.addRow(['Service Publié (Paramètres de Sécurité SI)', '', '']);
  secHeader.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.HEADER_BG } };
  secHeader.font = { bold: true, color: { argb: COLORS.HEADER_TEXT } };
  sheet4.mergeCells('A1:C1');

  const secParams = [
    ['Entrée DNS (site web)', data.securityParams?.dns_site_web || 'N/A'],
    ['Adresse IP Publique', data.securityParams?.ip_publique || 'N/A'],
    ['Adresse IP Interne (*)', data.securityParams?.ip_interne || ''],
    ['Adresse IP Virtuelle F5', data.securityParams?.ip_virtuelle_f5 || 'N/A'],
    ['Publication (*)', data.securityParams?.publication || 'DEV'],
    ['Date / Heure de la dernière Mise à jour', data.securityParams?.date_derniere_maj || '']
  ];

  secParams.forEach(([k, v]) => {
    const r = sheet4.addRow([k, v, '']);
    r.getCell(1).font = { bold: true };
    r.getCell(1).border = thinBorder;
    r.getCell(2).border = thinBorder;
  });

  sheet4.addRow([]); // Espace

  const ctrlHeader = sheet4.addRow(['Contrôles Sécurité SI', 'État et Conformité', 'Commentaires']);
  ctrlHeader.eachCell((cell) => {
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.ORANGE_TRC } };
    cell.font = { bold: true, color: { argb: COLORS.HEADER_TEXT } };
    cell.border = thinBorder;
  });

  (data.securityCompliance || []).forEach((ctrl) => {
    const r = sheet4.addRow([ctrl.control_name, ctrl.status, ctrl.comments || '/']);
    r.getCell(1).border = thinBorder;
    r.getCell(2).border = thinBorder;
    r.getCell(3).border = thinBorder;

    if (ctrl.status.startsWith('Conforme')) {
      r.getCell(2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.GREEN_BG } };
      r.getCell(2).font = { bold: true, color: { argb: COLORS.GREEN_TEXT } };
    } else if (ctrl.status === 'Non validé') {
      r.getCell(2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.RED_BG } };
      r.getCell(2).font = { bold: true, color: { argb: COLORS.RED_TEXT } };
    }
  });

  // Génération et Téléchargement du fichier .xlsx
  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const ipName = data.formPublication?.ip_address || data.formPublication?.app_name || 'vm';
  const filename = `formulaire_création_VM_${ipName}.xlsx`;
  saveAs(blob, filename);
};

// =========================================================================
// COMPOSANT MODAL DE GESTION DE VM
// =========================================================================
export default function VMFormModal({ requestId, initialData, onClose, onSuccess }) {
  const [activeTab, setActiveTab] = useState(1);
  const [loading, setLoading] = useState(false);

  // Initialisation des états avec initialData s'il existe
  const [structureInfo, setStructureInfo] = useState({
    pole: initialData?.structureInfo?.pole || 'ALGER',
    structure: initialData?.structureInfo?.structure || 'DTI',
    responsable_structure: initialData?.structureInfo?.responsable_structure || '/',
    responsable_service: initialData?.structureInfo?.responsable_service || 'INTRANET',
    contact: initialData?.structureInfo?.contact || 'younes samia, berkat siham, aloui adel'
  });

  const [formPublication, setFormPublication] = useState({
    publication_type: initialData?.formPublication?.publication_type || 'Intranet',
    target_population: initialData?.formPublication?.target_population || "Agents de la Direction EXP ainsi que les agents d'exploitation des 9 Directions Régionales.",
    app_name: initialData?.formPublication?.app_name || initialData?.app_name || 'vm_app_dev',
    dns_entry: initialData?.formPublication?.dns_entry || 'N/A',
    ip_address: initialData?.formPublication?.ip_address || initialData?.ip_address || '10.118.100.64',
    port: initialData?.formPublication?.port || '443',
    os_server: initialData?.formPublication?.os_server || 'Windows Server 2022'
  });

  const defaultStack = [
    { software_name: 'Apache Struts', exists: false, version: '' },
    { software_name: 'Apache Tomcat', exists: false, version: '' },
    { software_name: 'Apache/NCSA HTTP Server', exists: false, version: '' },
    { software_name: 'ASP', exists: false, version: '' },
    { software_name: 'ASP.NET', exists: false, version: '' },
    { software_name: 'BEA Systems WebLogic Server', exists: false, version: '' },
    { software_name: 'CGI', exists: false, version: '' },
    { software_name: 'Cisco', exists: false, version: '' },
    { software_name: 'Citrix', exists: false, version: '' },
    { software_name: 'Django', exists: false, version: '' },
    { software_name: 'Elasticsearch', exists: false, version: '' },
    { software_name: 'Front Page Server Extensions (FPSE)', exists: false, version: '' },
    { software_name: 'IBM DB2', exists: false, version: '' },
    { software_name: 'IIS', exists: false, version: '' },
    { software_name: 'Java Servlets/JSP', exists: false, version: '' },
    { software_name: 'Java', exists: false, version: '' },
    { software_name: 'Js', exists: false, version: '' },
    { software_name: 'JavaServer Faces (JSF)', exists: false, version: '' },
    { software_name: 'JBoss', exists: false, version: '' },
    { software_name: 'Jetty', exists: false, version: '' },
    { software_name: 'Joomla', exists: false, version: '' },
    { software_name: 'jQuery', exists: false, version: '' },
    { software_name: 'Lotus Domino', exists: false, version: '' },
    { software_name: 'Macromedia ColdFusion', exists: false, version: '' },
    { software_name: 'Macromedia JRun', exists: false, version: '' },
    { software_name: 'Microsoft SQL Server', exists: false, version: '' },
    { software_name: 'Microsoft Windows', exists: true, version: 'Server 2022' },
    { software_name: 'MongoDB', exists: false, version: '' },
    { software_name: 'MySQL', exists: false, version: '' },
    { software_name: 'Node.js', exists: true, version: '18.x' },
    { software_name: 'Novell', exists: false, version: '' },
    { software_name: 'Oracle', exists: true, version: '19c' },
    { software_name: 'Outlook Web Access', exists: false, version: '' },
    { software_name: 'PHP', exists: false, version: '' },
    { software_name: 'PostgreSQL', exists: false, version: '' },
    { software_name: 'Proxy Servers', exists: false, version: '' },
    { software_name: 'Ruby', exists: false, version: '' },
    { software_name: 'SSI (Server Side Includes)', exists: false, version: '' },
    { software_name: 'Sybase/ASE', exists: false, version: '' },
    { software_name: 'WebDAV', exists: false, version: '' },
    { software_name: 'WordPress', exists: false, version: '' },
    { software_name: 'C/C++', exists: false, version: '' },
    { software_name: 'Python', exists: true, version: '3.10' },
    { software_name: 'Sharepoint', exists: false, version: '' },
    { software_name: 'XML', exists: false, version: '' },
    { software_name: 'Ngnix (web server & revers proxy)', exists: true, version: '1.24' },
    { software_name: 'WSL (Ubuntu server 24.04)', exists: true, version: '24.04' }
  ];

  const [softwareStack, setSoftwareStack] = useState(initialData?.softwareStack || defaultStack);

  const [architectureDesc, setArchitectureDesc] = useState(
    initialData?.architecture_desc || 'Architecture Web / App / BDD : Reverse Proxy Nginx, API Python/Node.js et base de données Oracle.'
  );

  const [networkFlows, setNetworkFlows] = useState(
    initialData?.networkFlows || [
      {
        source: '10.10.0.0/16 (UTILISATEURS)',
        destination: '10.118.100.64',
        service: 'TCP/443',
        port: '443',
        flow_type: 'Flux applicatif Web',
        description: 'Trafic Web sécurisé HTTPS (Flux applicatif Web)'
      },
      {
        source: '10.5.1.0/24 (ADMINS)',
        destination: '10.118.100.64',
        service: 'TCP/22',
        port: '22',
        flow_type: "Flux d'administration",
        description: "Accès d'administration sécurisé SSH"
      }
    ]
  );

  const addNetworkFlow = () => {
    setNetworkFlows([
      ...networkFlows,
      { source: '', destination: '', service: 'TCP/80', port: '80', flow_type: 'Flux applicatif Web', description: '' }
    ]);
  };

  const removeNetworkFlow = (idx) => {
    setNetworkFlows(networkFlows.filter((_, i) => i !== idx));
  };

  const [securityParams, setSecurityParams] = useState(
    initialData?.securityParams || {
      dns_site_web: 'N/A',
      ip_publique: 'N/A',
      ip_interne: '10.118.100.64',
      ip_virtuelle_f5: 'N/A',
      publication: 'DEV',
      date_derniere_maj: '2026-09-06T09:04'
    }
  );

  const [securityCompliance, setSecurityCompliance] = useState(
    initialData?.securityCompliance || [
      { id: 1, control_name: 'Vérification de la mise en place du Service à publier dans la Zone DMZ', status: 'Non validé', comments: '/' },
      { id: 2, control_name: "Vérification de l'installation de l'Antivirus (avec une base de signature à jour)", status: 'Conforme (Actif & À jour)', comments: '' },
      { id: 3, control_name: "Vérification de l'utilisation de certificat TLS", status: 'Conforme (Certificat TLS actif)', comments: '' },
      { id: 4, control_name: "Réalisation d'un scan de vulnérabilités authentifié", status: 'En attente', comments: '' },
      { id: 5, control_name: 'Réalisation d\'un scan de Vulnérabilités Web', status: 'En attente', comments: '' },
      { id: 6, control_name: 'Scan de Conformité des configurations de sécurité appliquées', status: 'Conforme', comments: '' },
      { id: 7, control_name: "Revue de Code source de l'application", status: 'Conforme', comments: '' },
      { id: 8, control_name: "Vérification de l'application de la politique du moindre privilège pour chaque type d'utilisateur ayant accès au service.", status: 'Conforme (PoLP respecté)', comments: '' }
    ]
  );

  // Soumission au backend (Création ou Édition)
  const handleSubmit = async (e) => {
    e.preventDefault();
    setLoading(true);

    const payload = {
      structureInfo,
      formPublication,
      softwareStack,
      networkFlows,
      securityCompliance,
      securityParams,
      architecture_desc: architectureDesc
    };

    try {
      const isEdit = Boolean(initialData?.id);
      const url = isEdit
        ? `http://localhost:5000/api/vms/${initialData.id}`
        : `http://localhost:5000/api/migrations/${requestId}/vms`;

      const method = isEdit ? 'PUT' : 'POST';

      const res = await fetch(url, {
        method,
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });

      if (res.ok) {
        if (onSuccess) onSuccess();
        onClose();
      } else {
        alert("Erreur lors de l'enregistrement de la VM");
      }
    } catch (err) {
      console.error(err);
      alert("Une erreur réseau s'est produite.");
    } finally {
      setLoading(false);
    }
  };

  const handleExport = () => {
    exportVMToExcel({
      structureInfo,
      formPublication,
      softwareStack,
      networkFlows,
      securityCompliance,
      securityParams,
      architecture_desc: architectureDesc
    });
  };

  return createPortal(
    <div className="fixed inset-0 bg-black/80 flex justify-center items-center z-[9999] p-4 backdrop-blur-sm overflow-y-auto">
      <div className="bg-white rounded-2xl shadow-2xl w-full max-w-5xl border border-slate-200 overflow-hidden flex flex-col max-h-[88vh] my-auto">
        
        {/* Entête Modal */}
        <div className="bg-slate-900 text-white px-6 py-4 flex justify-between items-center border-b border-slate-800 shrink-0">
          <div>
            <h3 className="font-bold text-base flex items-center gap-2">
              <span>📋</span> Formulaire Technique VM : <span className="text-amber-400 font-mono">{formPublication.app_name || 'Nouvelle VM'}</span>
            </h3>
            <p className="text-[11px] text-amber-300 mt-0.5">
              Fiche technique officielle Sonatrach TRC | IP: <span className="font-mono">{formPublication.ip_address}</span>
            </p>
          </div>
          <button 
            type="button"
            onClick={onClose} 
            className="text-slate-400 hover:text-white font-bold text-xl h-8 w-8 rounded-lg hover:bg-slate-800 flex items-center justify-center transition"
          >
            ✕
          </button>
        </div>

        {/* Navigation des 4 Onglets */}
        <div className="flex border-b bg-slate-100 text-xs font-semibold overflow-x-auto shrink-0">
          {[
            { id: 1, title: 'Principale', desc: 'Demandeur TRC' },
            { id: 2, title: 'Publication VM', desc: 'Fiche Technique & Stack' },
            { id: 3, title: 'Informations liées au service', desc: 'Flux & Architecture' },
            { id: 4, title: 'Suivie des Non conformités', desc: 'Sécurité SI' }
          ].map(tab => (
            <button
              key={tab.id}
              type="button"
              onClick={() => setActiveTab(tab.id)}
              className={`flex-1 min-w-[190px] py-3 px-3 text-left border-b-2 transition ${
                activeTab === tab.id 
                  ? 'border-amber-600 text-amber-700 bg-white font-bold shadow-sm' 
                  : 'border-transparent text-slate-600 hover:text-slate-900 hover:bg-slate-200/50'
              }`}
            >
              <div className="font-bold">{tab.id}. {tab.title}</div>
              <div className="text-[10px] text-slate-500 font-normal">{tab.desc}</div>
            </button>
          ))}
        </div>

        {/* Corps du Formulaire */}
        <form onSubmit={handleSubmit} className="p-6 overflow-y-auto flex-1 space-y-6 text-xs bg-slate-50/50">
          
          {/* ONGLET 1 : PRINCIPALE */}
          {activeTab === 1 && (
            <div className="space-y-4">
              <div className="bg-white p-5 rounded-xl border border-slate-200 shadow-sm space-y-4">
                <div className="border-b pb-2">
                  <h4 className="font-bold text-sm text-slate-800">Informations Demandeur & Structure Sonatrach TRC</h4>
                  <p className="text-[11px] text-slate-500">Renseignez les champs de la structure initiatrice de la demande.</p>
                </div>

                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  <div>
                    <label className="block font-semibold text-slate-700 mb-1">
                      Pôle <span className="text-red-500 font-bold">*</span>
                    </label>
                    <select 
                      required 
                      className="w-full border p-2.5 rounded-lg bg-white text-xs outline-none focus:ring-2 focus:ring-amber-500"
                      value={structureInfo.pole}
                      onChange={e => setStructureInfo({...structureInfo, pole: e.target.value})}
                    >
                      <option value="ALGER">ALGER</option>
                      <option value="HAOUD EL HAMRA">HAOUD EL HAMRA</option>
                      <option value="SKIKDA">SKIKDA</option>
                      <option value="ARZEW">ARZEW</option>
                      <option value="BEJAIA">BEJAIA</option>
                      <option value="LAGHOUAT">LAGHOUAT</option>
                    </select>
                  </div>

                  <div>
                    <label className="block font-semibold text-slate-700 mb-1">
                      Structure <span className="text-red-500 font-bold">*</span>
                    </label>
                    <input 
                      type="text" 
                      required 
                      className="w-full border p-2.5 rounded-lg bg-white text-xs outline-none focus:ring-2 focus:ring-amber-500"
                      value={structureInfo.structure}
                      onChange={e => setStructureInfo({...structureInfo, structure: e.target.value})}
                    />
                  </div>

                  <div>
                    <label className="block font-semibold text-slate-700 mb-1">
                      Responsable de la Structure <span className="text-red-500 font-bold">*</span>
                    </label>
                    <input 
                      type="text" 
                      required 
                      className="w-full border p-2.5 rounded-lg bg-white text-xs outline-none focus:ring-2 focus:ring-amber-500"
                      value={structureInfo.responsable_structure}
                      onChange={e => setStructureInfo({...structureInfo, responsable_structure: e.target.value})}
                    />
                  </div>

                  <div>
                    <label className="block font-semibold text-slate-700 mb-1">
                      Responsable du service à publier <span className="text-red-500 font-bold">*</span>
                    </label>
                    <input 
                      type="text" 
                      required 
                      className="w-full border p-2.5 rounded-lg bg-white text-xs outline-none focus:ring-2 focus:ring-amber-500 text-red-600 font-bold"
                      value={structureInfo.responsable_service}
                      onChange={e => setStructureInfo({...structureInfo, responsable_service: e.target.value})}
                    />
                  </div>

                  <div className="md:col-span-2">
                    <label className="block font-semibold text-slate-700 mb-1">
                      Contact <span className="text-red-500 font-bold">*</span>
                    </label>
                    <input 
                      type="text" 
                      required 
                      className="w-full border p-2.5 rounded-lg bg-white text-xs outline-none focus:ring-2 focus:ring-amber-500"
                      value={structureInfo.contact}
                      onChange={e => setStructureInfo({...structureInfo, contact: e.target.value})}
                    />
                  </div>
                </div>
              </div>

              <div className="bg-amber-50/70 border border-amber-300 p-4 rounded-xl text-slate-700 space-y-2">
                <p className="font-bold text-amber-900 text-xs">(*) obligatoire</p>
                <div className="text-[11px] space-y-1">
                  <p>• Les onglets <strong>"Informations liées au service"</strong> sont à renseigner par la Structure demanderesse.</p>
                  <p>• L'onglet <strong>"Suivi des non conformités"</strong> est à renseigner par la direction Sécurité SI.</p>
                </div>
              </div>
            </div>
          )}

          {/* ONGLET 2 : PUBLICATION VM */}
          {activeTab === 2 && (
            <div className="space-y-5">
              <div className="bg-white p-5 rounded-xl border border-slate-200 shadow-sm space-y-4">
                <div className="border-b pb-2">
                  <h4 className="font-bold text-sm text-slate-800">Formulaire technique de publication</h4>
                </div>

                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  <div>
                    <label className="block font-semibold text-slate-700 mb-1">
                      Type de publication <span className="text-red-500 font-bold">*</span>
                    </label>
                    <select 
                      required 
                      className="w-full border p-2.5 rounded-lg bg-white text-xs outline-none focus:ring-2 focus:ring-amber-500"
                      value={formPublication.publication_type}
                      onChange={e => setFormPublication({...formPublication, publication_type: e.target.value})}
                    >
                      <option value="Intranet">Intranet</option>
                      <option value="Internet">Internet</option>
                    </select>
                  </div>

                  <div>
                    <label className="block font-semibold text-slate-700 mb-1">
                      Population exploitant service à publier <span className="text-red-500 font-bold">*</span>
                    </label>
                    <input 
                      type="text" 
                      required 
                      className="w-full border p-2.5 rounded-lg bg-white text-xs outline-none focus:ring-2 focus:ring-amber-500"
                      value={formPublication.target_population}
                      onChange={e => setFormPublication({...formPublication, target_population: e.target.value})}
                    />
                  </div>

                  <div className="md:col-span-2">
                    <label className="block font-semibold text-slate-700 mb-1">
                      Description du service à publier (Nom VM) <span className="text-red-500 font-bold">*</span>
                    </label>
                    <input 
                      type="text" 
                      required 
                      className="w-full border p-2.5 rounded-lg bg-white text-xs outline-none focus:ring-2 focus:ring-amber-500 font-mono"
                      value={formPublication.app_name}
                      onChange={e => setFormPublication({...formPublication, app_name: e.target.value})}
                    />
                  </div>

                  <div>
                    <label className="block font-semibold text-slate-500 mb-1">Entrée DNS (Optionnel)</label>
                    <input 
                      type="text" 
                      className="w-full border p-2.5 rounded-lg bg-slate-50 text-xs outline-none focus:ring-2 focus:ring-amber-500"
                      value={formPublication.dns_entry}
                      onChange={e => setFormPublication({...formPublication, dns_entry: e.target.value})}
                    />
                  </div>

                  <div>
                    <label className="block font-semibold text-slate-700 mb-1">
                      Adresse IP <span className="text-red-500 font-bold">*</span>
                    </label>
                    <input 
                      type="text" 
                      required 
                      className="w-full border p-2.5 rounded-lg bg-white text-xs outline-none focus:ring-2 focus:ring-amber-500 font-mono"
                      value={formPublication.ip_address}
                      onChange={e => setFormPublication({...formPublication, ip_address: e.target.value})}
                    />
                  </div>

                  <div>
                    <label className="block font-semibold text-slate-700 mb-1">
                      Port <span className="text-red-500 font-bold">*</span>
                    </label>
                    <input 
                      type="text" 
                      required 
                      className="w-full border p-2.5 rounded-lg bg-white text-xs outline-none focus:ring-2 focus:ring-amber-500 font-mono"
                      value={formPublication.port}
                      onChange={e => setFormPublication({...formPublication, port: e.target.value})}
                    />
                  </div>

                  <div>
                    <label className="block font-semibold text-slate-700 mb-1">
                      OS Serveur <span className="text-red-500 font-bold">*</span>
                    </label>
                    <select 
                      required 
                      className="w-full border p-2.5 rounded-lg bg-white text-xs outline-none focus:ring-2 focus:ring-amber-500"
                      value={formPublication.os_server}
                      onChange={e => setFormPublication({...formPublication, os_server: e.target.value})}
                    >
                      <option value="Windows Server 2022">Windows Server 2022</option>
                      <option value="Windows Server 2019">Windows Server 2019</option>
                      <option value="Linux RHEL 8">Linux RHEL 8</option>
                      <option value="Linux RHEL 9">Linux RHEL 9</option>
                      <option value="Ubuntu Server 24.04">Ubuntu Server 24.04</option>
                    </select>
                  </div>
                </div>
              </div>

              <div className="bg-white p-5 rounded-xl border border-slate-200 shadow-sm space-y-3">
                <div className="border-b pb-2 flex justify-between items-center">
                  <h4 className="font-bold text-sm text-slate-800">Configuration Software & Technologies</h4>
                  <span className="text-[11px] text-slate-500 font-medium">Existance (Oui / Non) & Version</span>
                </div>

                <div className="grid grid-cols-1 md:grid-cols-2 gap-2 max-h-80 overflow-y-auto pr-1">
                  {softwareStack.map((sw, idx) => (
                    <div key={sw.software_name || idx} className="flex items-center justify-between bg-slate-50 p-2.5 rounded-lg border border-slate-200 gap-2">
                      <span className="font-medium text-slate-800 text-xs truncate w-1/2" title={sw.software_name}>
                        {sw.software_name}
                      </span>
                      
                      <div className="flex items-center space-x-2">
                        <select 
                          className={`border p-1 rounded text-xs font-bold ${
                            sw.exists ? 'bg-emerald-50 text-emerald-700 border-emerald-300' : 'bg-slate-100 text-slate-500'
                          }`}
                          value={sw.exists ? 'Oui' : 'Non'}
                          onChange={e => {
                            const u = [...softwareStack];
                            u[idx].exists = e.target.value === 'Oui';
                            if (e.target.value === 'Non') u[idx].version = '';
                            setSoftwareStack(u);
                          }}
                        >
                          <option value="Non">Non</option>
                          <option value="Oui">Oui</option>
                        </select>

                        <input 
                          type="text" 
                          disabled={!sw.exists}
                          placeholder={sw.exists ? "version" : "—"} 
                          className={`border p-1 rounded text-xs w-24 ${!sw.exists ? 'bg-slate-100 text-slate-400' : 'bg-white'}`}
                          value={sw.version}
                          onChange={e => {
                            const u = [...softwareStack];
                            u[idx].version = e.target.value;
                            setSoftwareStack(u);
                          }}
                        />
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            </div>
          )}

          {/* ONGLET 3 : INFORMATIONS LIÉES AU SERVICE / FLUX */}
          {activeTab === 3 && (
            <div className="space-y-5">
              <div className="bg-white p-5 rounded-xl border border-slate-200 shadow-sm space-y-3">
                <div className="border-b pb-2">
                  <h4 className="font-bold text-sm text-slate-800">Architecture du service</h4>
                </div>
                <textarea 
                  required
                  rows="3"
                  className="w-full border p-2.5 rounded-lg bg-white text-xs outline-none focus:ring-2 focus:ring-amber-500"
                  value={architectureDesc}
                  onChange={e => setArchitectureDesc(e.target.value)}
                />
              </div>

              <div className="bg-white p-5 rounded-xl border border-slate-200 shadow-sm space-y-3">
                <div className="flex justify-between items-center border-b pb-2">
                  <div>
                    <h4 className="font-bold text-sm text-slate-800">Matrice des flux</h4>
                    <p className="text-[11px] text-slate-500">Renseignez les lignes d'autorisations de trafic réseau.</p>
                  </div>
                  <button 
                    type="button" 
                    onClick={addNetworkFlow}
                    className="bg-amber-600 hover:bg-amber-700 text-white text-xs font-bold px-3 py-1.5 rounded-lg shadow-sm transition"
                  >
                    + Ajouter une ligne de flux
                  </button>
                </div>

                <div className="space-y-3">
                  {networkFlows.map((flow, idx) => (
                    <div key={idx} className="bg-slate-50 p-3 rounded-lg border border-slate-200 space-y-2">
                      <div className="flex justify-between items-center border-b pb-1">
                        <span className="font-bold text-slate-700">Flux #{idx + 1}</span>
                        {networkFlows.length > 1 && (
                          <button 
                            type="button" 
                            onClick={() => removeNetworkFlow(idx)}
                            className="text-red-600 hover:text-red-800 font-bold text-xs"
                          >
                            Supprimer
                          </button>
                        )}
                      </div>

                      <div className="grid grid-cols-1 md:grid-cols-6 gap-2">
                        <div>
                          <label className="block text-[10px] font-bold text-slate-600 mb-0.5">Source *</label>
                          <input type="text" required placeholder="ex: 10.10.0.0/16" className="w-full border p-1.5 rounded bg-white text-xs font-mono" value={flow.source} onChange={e => { const u = [...networkFlows]; u[idx].source = e.target.value; setNetworkFlows(u); }} />
                        </div>
                        <div>
                          <label className="block text-[10px] font-bold text-slate-600 mb-0.5">Destination *</label>
                          <input type="text" required placeholder="ex: 10.118.100.64" className="w-full border p-1.5 rounded bg-white text-xs font-mono" value={flow.destination} onChange={e => { const u = [...networkFlows]; u[idx].destination = e.target.value; setNetworkFlows(u); }} />
                        </div>
                        <div>
                          <label className="block text-[10px] font-bold text-slate-600 mb-0.5">Service *</label>
                          <input type="text" required placeholder="ex: TCP/443" className="w-full border p-1.5 rounded bg-white text-xs" value={flow.service} onChange={e => { const u = [...networkFlows]; u[idx].service = e.target.value; setNetworkFlows(u); }} />
                        </div>
                        <div>
                          <label className="block text-[10px] font-bold text-slate-600 mb-0.5">Ports *</label>
                          <input type="text" required placeholder="ex: 443" className="w-full border p-1.5 rounded bg-white text-xs font-mono" value={flow.port} onChange={e => { const u = [...networkFlows]; u[idx].port = e.target.value; setNetworkFlows(u); }} />
                        </div>
                        <div>
                          <label className="block text-[10px] font-bold text-slate-600 mb-0.5">Type de flux *</label>
                          <select className="w-full border p-1.5 rounded bg-white text-xs" value={flow.flow_type} onChange={e => { const u = [...networkFlows]; u[idx].flow_type = e.target.value; setNetworkFlows(u); }}>
                            <option value="Flux applicatif Web">Flux applicatif Web</option>
                            <option value="Flux d'administration">Flux d'administration</option>
                            <option value="Flux de base de données">Flux de base de données</option>
                          </select>
                        </div>
                        <div>
                          <label className="block text-[10px] font-bold text-slate-600 mb-0.5">Description (Optionnel)</label>
                          <input type="text" placeholder="Description" className="w-full border p-1.5 rounded bg-white text-xs" value={flow.description} onChange={e => { const u = [...networkFlows]; u[idx].description = e.target.value; setNetworkFlows(u); }} />
                        </div>
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            </div>
          )}

          {/* ONGLET 4 : SUIVI DES NON CONFORMITÉS */}
          {activeTab === 4 && (
            <div className="space-y-5">
              <div className="bg-white p-5 rounded-xl border border-slate-200 shadow-sm space-y-4">
                <div className="border-b pb-2">
                  <h4 className="font-bold text-sm text-slate-800">Service Publié (Paramètres de Sécurité SI)</h4>
                </div>

                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  <div>
                    <label className="block font-semibold text-slate-500 mb-1">Entrée DNS (site web) - Optionnel</label>
                    <input 
                      type="text" 
                      className="w-full border p-2.5 rounded-lg bg-slate-50 text-xs outline-none focus:ring-2 focus:ring-amber-500"
                      value={securityParams.dns_site_web}
                      onChange={e => setSecurityParams({...securityParams, dns_site_web: e.target.value})}
                    />
                  </div>

                  <div>
                    <label className="block font-semibold text-slate-500 mb-1">Adresse IP Publique - Optionnel</label>
                    <input 
                      type="text" 
                      className="w-full border p-2.5 rounded-lg bg-slate-50 text-xs outline-none focus:ring-2 focus:ring-amber-500"
                      value={securityParams.ip_publique}
                      onChange={e => setSecurityParams({...securityParams, ip_publique: e.target.value})}
                    />
                  </div>

                  <div>
                    <label className="block font-semibold text-slate-700 mb-1">
                      Adresse IP Interne <span className="text-red-500 font-bold">*</span>
                    </label>
                    <input 
                      type="text" 
                      required 
                      className="w-full border p-2.5 rounded-lg bg-white text-xs outline-none focus:ring-2 focus:ring-amber-500 font-mono"
                      value={securityParams.ip_interne}
                      onChange={e => setSecurityParams({...securityParams, ip_interne: e.target.value})}
                    />
                  </div>

                  <div>
                    <label className="block font-semibold text-slate-500 mb-1">Adresse IP Virtuelle F5 - Optionnel</label>
                    <input 
                      type="text" 
                      className="w-full border p-2.5 rounded-lg bg-slate-50 text-xs outline-none focus:ring-2 focus:ring-amber-500"
                      value={securityParams.ip_virtuelle_f5}
                      onChange={e => setSecurityParams({...securityParams, ip_virtuelle_f5: e.target.value})}
                    />
                  </div>

                  <div>
                    <label className="block font-semibold text-slate-700 mb-1">
                      Publication <span className="text-red-500 font-bold">*</span>
                    </label>
                    <select 
                      required 
                      className="w-full border p-2.5 rounded-lg bg-white text-xs outline-none focus:ring-2 focus:ring-amber-500 font-bold"
                      value={securityParams.publication}
                      onChange={e => setSecurityParams({...securityParams, publication: e.target.value})}
                    >
                      <option value="DEV">DEV</option>
                      <option value="PROD">PROD</option>
                    </select>
                  </div>

                  <div>
                    <label className="block font-semibold text-slate-700 mb-1">
                      Date / Heure de la dernière Mise à jour <span className="text-red-500 font-bold">*</span>
                    </label>
                    <input 
                      type="datetime-local" 
                      required 
                      className="w-full border p-2.5 rounded-lg bg-white text-xs outline-none focus:ring-2 focus:ring-amber-500 text-amber-700 font-bold"
                      value={securityParams.date_derniere_maj}
                      onChange={e => setSecurityParams({...securityParams, date_derniere_maj: e.target.value})}
                    />
                  </div>
                </div>
              </div>

              <div className="bg-white p-5 rounded-xl border border-slate-200 shadow-sm space-y-3">
                <div className="border-b pb-2">
                  <h4 className="font-bold text-sm text-slate-800">Contrôles de Conformité Sécurité SI</h4>
                </div>

                <div className="space-y-2.5">
                  {securityCompliance.map((ctrl, idx) => (
                    <div key={ctrl.id || idx} className="bg-slate-50 p-3 rounded-lg border border-slate-200 flex flex-col md:flex-row md:items-center justify-between gap-3">
                      <div className="flex-1">
                        <span className="font-bold text-amber-700 mr-2">#{idx + 1}</span>
                        <span className="font-semibold text-slate-800 text-xs">{ctrl.control_name}</span>
                      </div>

                      <div className="flex items-center space-x-2">
                        <select 
                          className={`border p-1.5 rounded-md font-bold text-xs bg-white ${
                            ctrl.status.startsWith('Conforme') ? 'text-emerald-700 border-emerald-300' :
                            ctrl.status === 'Non validé' ? 'text-red-600 border-red-300' : 'text-slate-600'
                          }`}
                          value={ctrl.status}
                          onChange={e => {
                            const updated = [...securityCompliance];
                            updated[idx].status = e.target.value;
                            setSecurityCompliance(updated);
                          }}
                        >
                          <option value="Conforme">Conforme</option>
                          <option value="Conforme (Actif & À jour)">Conforme (Actif & À jour)</option>
                          <option value="Conforme (Certificat TLS actif)">Conforme (Certificat TLS actif)</option>
                          <option value="Conforme (PoLP respecté)">Conforme (PoLP respecté)</option>
                          <option value="Non validé">Non validé</option>
                          <option value="En attente">En attente</option>
                        </select>

                        <input 
                          type="text" 
                          placeholder="Commentaires (Direction Sécurité SI)"
                          className="border p-1.5 rounded-md text-xs w-64 bg-white"
                          value={ctrl.comments}
                          onChange={e => {
                            const updated = [...securityCompliance];
                            updated[idx].comments = e.target.value;
                            setSecurityCompliance(updated);
                          }}
                        />
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            </div>
          )}

          {/* Pied de Page & Actions */}
          <div className="pt-4 border-t flex justify-between items-center bg-white p-4 -mx-6 -mb-6 mt-4 shrink-0">
            <button 
              type="button" 
              onClick={onClose} 
              className="px-4 py-2 border rounded-lg font-semibold text-xs text-slate-600 hover:bg-slate-100 transition"
            >
              Annuler
            </button>
            
            <div className="flex items-center space-x-2">
              <button 
                type="button"
                onClick={handleExport}
                className="px-4 py-2 bg-emerald-600 hover:bg-emerald-700 text-white font-bold rounded-lg text-xs shadow transition flex items-center gap-1.5"
                title="Télécharger la fiche Excel spécifique à cette VM"
              >
                <span>📊</span> Exporter cette VM (.xlsx)
              </button>

              {activeTab > 1 && (
                <button 
                  type="button" 
                  onClick={() => setActiveTab(activeTab - 1)} 
                  className="px-4 py-2 border rounded-lg font-semibold text-xs text-slate-700 hover:bg-slate-100 transition"
                >
                  ← Précédent
                </button>
              )}

              {activeTab < 4 ? (
                <button 
                  type="button" 
                  onClick={() => setActiveTab(activeTab + 1)} 
                  className="px-4 py-2 bg-slate-900 hover:bg-slate-800 text-white font-bold rounded-lg text-xs shadow transition"
                >
                  Suivant →
                </button>
              ) : (
                <button 
                  type="submit" 
                  disabled={loading}
                  className="px-5 py-2 bg-amber-600 hover:bg-amber-700 text-white font-bold rounded-lg text-xs shadow transition disabled:opacity-50"
                >
                  {loading ? 'Enregistrement...' : initialData?.id ? 'Enregistrer les Modifications' : 'Enregistrer la Fiche Complète'}
                </button>
              )}
            </div>
          </div>

        </form>

      </div>
    </div>,
    document.body
  );
}