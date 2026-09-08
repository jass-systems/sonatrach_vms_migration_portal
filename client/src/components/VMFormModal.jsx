import React, { useState, useEffect } from 'react';
import { createPortal } from 'react-dom';
import { exportVMToExcel } from '../utils/exportVMExcel';

export { exportVMToExcel };

const DEFAULT_SOFTWARE_STACK = [
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

const DEFAULT_NETWORK_FLOWS = [
  {
    id: 'flow-1',
    source: '10.10.0.0/16 (UTILISATEURS)',
    destination: '10.118.100.64',
    service: 'TCP/443',
    port: '443',
    flow_type: 'Flux applicatif Web',
    description: 'Trafic Web sécurisé HTTPS (Flux applicatif Web)'
  },
  {
    id: 'flow-2',
    source: '10.5.1.0/24 (ADMINS)',
    destination: '10.118.100.64',
    service: 'TCP/22',
    port: '22',
    flow_type: "Flux d'administration",
    description: "Accès d'administration sécurisé SSH"
  }
];

const DEFAULT_SECURITY_COMPLIANCE = [
  { id: 1, control_name: 'Vérification de la mise en place du Service à publier dans la Zone DMZ', status: 'Non validé', comments: '/' },
  { id: 2, control_name: "Vérification de l'installation de l'Antivirus (avec une base de signature à jour)", status: 'Conforme (Actif & À jour)', comments: '/' },
  { id: 3, control_name: "Vérification de l'utilisation de certificat TLS", status: 'Conforme (Certificat TLS actif)', comments: '/' },
  { id: 4, control_name: "Réalisation d'un scan de vulnérabilités authentifié", status: 'En attente', comments: '/' },
  { id: 5, control_name: 'Réalisation d\'un scan de Vulnérabilités Web', status: 'En attente', comments: '/' },
  { id: 6, control_name: 'Scan de Conformité des configurations de sécurité appliquées', status: 'Conforme', comments: '/' },
  { id: 7, control_name: "Revue de Code source de l'application", status: 'Conforme', comments: '/' },
  { id: 8, control_name: "Vérification de l'application de la politique du moindre privilège pour chaque type d'utilisateur ayant accès au service.", status: 'Conforme (PoLP respecté)', comments: '/' }
];

const getInitialStack = (data) => {
  const rawStack = data?.softwareStack || data?.software_stack;
  if (!rawStack?.length) return DEFAULT_SOFTWARE_STACK;

  return DEFAULT_SOFTWARE_STACK.map(defaultItem => {
    const found = rawStack.find(
      item => (item.software_name || item.name || '').trim().toLowerCase() === defaultItem.software_name.trim().toLowerCase()
    );
    if (found) {
      return {
        software_name: defaultItem.software_name,
        exists: found.exists !== undefined ? Boolean(found.exists) : Boolean(found.is_present),
        version: found.version || ''
      };
    }
    return defaultItem;
  });
};

const getInitialFlows = (data) => {
  const source = data?.networkFlows || data?.network_flows || DEFAULT_NETWORK_FLOWS;
  return source.map((flow, idx) => ({ ...flow, id: flow.id || `flow-${Date.now()}-${idx}` }));
};

const getInitialSecurityParams = (data) => ({
  dns_site_web: data?.securityParams?.dns_site_web || data?.security_params?.dns_site_web || data?.dns_entry || 'N/A',
  ip_publique: data?.securityParams?.ip_publique || data?.security_params?.ip_publique || data?.public_ip || 'N/A',
  ip_interne: data?.securityParams?.ip_interne || data?.security_params?.ip_interne || data?.ip_address || '10.118.100.64',
  ip_virtuelle: data?.securityParams?.ip_virtuelle || data?.security_params?.ip_virtuelle || '/',
  publication: data?.securityParams?.publication || data?.security_params?.publication || 'DEV',
  date_derniere_maj: data?.securityParams?.date_derniere_maj || data?.security_params?.date_derniere_maj || new Date().toISOString().slice(0, 16)
});

const getInitialCompliance = (data) => {
  return data?.securityCompliance || data?.security_compliance || DEFAULT_SECURITY_COMPLIANCE;
};

export default function VMFormModal({ requestId, initialData, onClose, onSuccess }) {
  const [activeTab, setActiveTab] = useState(1);
  const [loading, setLoading] = useState(false);

  const [structureInfo, setStructureInfo] = useState({
    pole: initialData?.structureInfo?.pole || initialData?.pole || 'ALGER',
    structure: initialData?.structureInfo?.structure || initialData?.structure || 'TRC Siège / EXP',
    responsable_structure: initialData?.structureInfo?.responsable_structure || initialData?.responsable_structure || '/',
    responsable_service: initialData?.structureInfo?.responsable_service || initialData?.responsable_service || 'INTRANET',
    contact: initialData?.structureInfo?.contact || initialData?.contact || 'Berkat Siham'
  });

  const [formPublication, setFormPublication] = useState({
    publication_type: initialData?.formPublication?.publication_type || initialData?.publication_type || 'Intranet',
    target_population: initialData?.formPublication?.target_population || initialData?.target_population || "Agents de la Direction EXP ainsi que les agents d'exploitation des 9 Directions Régionales.",
    app_name: initialData?.formPublication?.app_name || initialData?.app_name || 'vm_app_dev',
    dns_entry: initialData?.formPublication?.dns_entry || initialData?.dns_entry || 'N/A',
    ip_address: initialData?.formPublication?.ip_address || initialData?.ip_address || '10.118.100.64',
    port: initialData?.formPublication?.port || initialData?.port || '443',
    os_server: initialData?.formPublication?.os_server || initialData?.os_server || 'Windows Server 2022'
  });

  const [softwareStack, setSoftwareStack] = useState(() => getInitialStack(initialData));
  const [architectureDesc, setArchitectureDesc] = useState(
    initialData?.architecture_desc || initialData?.architectureDesc || 'Architecture Web / App / BDD : Reverse Proxy Nginx, API Python/Node.js et base de données Oracle.'
  );
  const [networkFlows, setNetworkFlows] = useState(() => getInitialFlows(initialData));
  const [securityParams, setSecurityParams] = useState(() => getInitialSecurityParams(initialData));
  const [securityCompliance, setSecurityCompliance] = useState(() => getInitialCompliance(initialData));

  useEffect(() => {
    if (!initialData) return;

    setStructureInfo({
      pole: initialData.structureInfo?.pole || initialData.pole || 'ALGER',
      structure: initialData.structureInfo?.structure || initialData.structure || 'TRC Siège / EXP',
      responsable_structure: initialData.structureInfo?.responsable_structure || initialData.responsable_structure || '/',
      responsable_service: initialData.structureInfo?.responsable_service || initialData.responsable_service || 'INTRANET',
      contact: initialData.structureInfo?.contact || initialData.contact || 'Berkat Siham'
    });

    setFormPublication({
      publication_type: initialData.formPublication?.publication_type || initialData.publication_type || 'Intranet',
      target_population: initialData.formPublication?.target_population || initialData.target_population || "Agents de la Direction EXP ainsi que les agents d'exploitation des 9 Directions Régionales.",
      app_name: initialData.formPublication?.app_name || initialData.app_name || 'vm_app_dev',
      dns_entry: initialData.formPublication?.dns_entry || initialData.dns_entry || 'N/A',
      ip_address: initialData.formPublication?.ip_address || initialData.ip_address || '10.118.100.64',
      port: initialData.formPublication?.port || initialData.port || '443',
      os_server: initialData.formPublication?.os_server || initialData.os_server || 'Windows Server 2022'
    });

    setSoftwareStack(getInitialStack(initialData));
    setArchitectureDesc(
      initialData.architecture_desc || initialData.architectureDesc || 'Architecture Web / App / BDD : Reverse Proxy Nginx, API Python/Node.js et base de données Oracle.'
    );
    setNetworkFlows(getInitialFlows(initialData));
    setSecurityParams(getInitialSecurityParams(initialData));
    setSecurityCompliance(getInitialCompliance(initialData));
  }, [initialData]);

  const addNetworkFlow = () => {
    setNetworkFlows(prev => [
      ...prev,
      { id: `flow-${Date.now()}`, source: '', destination: '', service: 'TCP/80', port: '80', flow_type: 'Flux applicatif Web', description: '' }
    ]);
  };

  const removeNetworkFlow = (id) => {
    setNetworkFlows(prev => prev.filter((flow) => flow.id !== id));
  };

  const updateNetworkFlow = (id, field, value) => {
    setNetworkFlows(prev => prev.map(flow => (flow.id === id ? { ...flow, [field]: value } : flow)));
  };

  const getPayload = () => ({
    id: initialData?.id || initialData?._id || initialData?.vm_id,
    ...structureInfo,
    ...formPublication,
    structureInfo,
    formPublication,
    softwareStack,
    networkFlows,
    securityCompliance,
    securityParams,
    architecture_desc: architectureDesc,
    app_name: formPublication.app_name,
    ip_address: formPublication.ip_address,
    port: formPublication.port,
    os_server: formPublication.os_server
  });

  const handleSubmit = async (e) => {
    e.preventDefault();
    setLoading(true);

    const vmId = initialData?.id || initialData?._id || initialData?.vm_id;
    const isEdit = Boolean(vmId);
    const payload = getPayload();

    try {
      const url = isEdit
        ? `http://localhost:5000/api/vms/${vmId}`
        : `http://localhost:5000/api/migrations/${requestId}/vms`;

      const res = await fetch(url, {
        method: isEdit ? 'PUT' : 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });

      if (!res.ok) {
        const errorData = await res.json().catch(() => ({}));
        throw new Error(errorData.message || `Erreur serveur: ${res.status}`);
      }

      const responseData = await res.json().catch(() => ({}));
      const returnedVm = responseData.vm || responseData || {};

      const completeVmData = {
        ...initialData,
        ...payload,
        ...returnedVm,
        softwareStack: returnedVm.softwareStack || returnedVm.software_stack || payload.softwareStack,
        networkFlows: returnedVm.networkFlows || returnedVm.network_flows || payload.networkFlows,
        securityCompliance: returnedVm.securityCompliance || returnedVm.security_compliance || payload.securityCompliance,
        securityParams: returnedVm.securityParams || returnedVm.security_params || payload.securityParams,
      };

      if (onSuccess) onSuccess(completeVmData);
      onClose();
    } catch (err) {
      console.error('Erreur lors de la sauvegarde de la VM :', err);
      alert(`Impossible d'enregistrer les modifications : ${err.message}`);
    } finally {
      setLoading(false);
    }
  };

  return createPortal(
    <div className="fixed inset-0 bg-black/80 flex justify-center items-center z-[9999] p-4 backdrop-blur-sm overflow-y-auto">
      <div className="bg-white rounded-2xl shadow-2xl w-full max-w-5xl border border-slate-200 overflow-hidden flex flex-col max-h-[88vh] my-auto">
        
        {/* Header */}
        <div className="bg-slate-900 text-white px-6 py-4 flex justify-between items-center border-b border-slate-800 shrink-0">
          <div>
            <h3 className="font-bold text-base flex items-center gap-2">
              <span>📋</span> Formulaire Technique VM : <span className="text-amber-400 font-mono">{formPublication.app_name || 'Nouvelle VM'}</span>
            </h3>
            <p className="text-[11px] text-amber-300 mt-0.5">
              Fiche technique Sonatrach TRC | IP: <span className="font-mono">{formPublication.ip_address}</span>
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

        {/* Navigation Tabs */}
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

        {/* Form Body */}
        <form onSubmit={handleSubmit} className="p-6 overflow-y-auto flex-1 space-y-6 text-xs bg-slate-50/50">
          
          {/* TAB 1 : PRINCIPALE */}
          {activeTab === 1 && (
            <div className="space-y-4">
              <div className="bg-white p-5 rounded-xl border border-slate-200 shadow-sm space-y-4">
                <div className="border-b pb-2">
                  <h4 className="font-bold text-sm text-slate-800">Informations Demandeur & Structure Sonatrach TRC</h4>
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
                      className="w-full border p-2.5 rounded-lg bg-white text-xs outline-none focus:ring-2 focus:ring-amber-500 font-bold"
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
            </div>
          )}

          {/* TAB 2 : PUBLICATION VM */}
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
                </div>

                <div className="grid grid-cols-1 md:grid-cols-2 gap-2 max-h-80 overflow-y-auto pr-1">
                  {softwareStack.map((sw, idx) => (
                    <div key={sw.software_name || `sw-${idx}`} className="flex items-center justify-between bg-slate-50 p-2.5 rounded-lg border border-slate-200 gap-2">
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
                            const exists = e.target.value === 'Oui';
                            setSoftwareStack(prev =>
                              prev.map((item, i) =>
                                i === idx ? { ...item, exists, version: exists ? item.version : '' } : item
                              )
                            );
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
                            const version = e.target.value;
                            setSoftwareStack(prev =>
                              prev.map((item, i) => (i === idx ? { ...item, version } : item))
                            );
                          }}
                        />
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            </div>
          )}

          {/* TAB 3 : FLUX & ARCHITECTURE */}
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
                  <h4 className="font-bold text-sm text-slate-800">Matrice des flux</h4>
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
                    <div key={flow.id} className="bg-slate-50 p-3 rounded-lg border border-slate-200 space-y-2">
                      <div className="flex justify-between items-center border-b pb-1">
                        <span className="font-bold text-slate-700">Flux #{idx + 1}</span>
                        {networkFlows.length > 1 && (
                          <button 
                            type="button" 
                            onClick={() => removeNetworkFlow(flow.id)}
                            className="text-red-600 hover:text-red-800 font-bold text-xs"
                          >
                            Supprimer
                          </button>
                        )}
                      </div>

                      <div className="grid grid-cols-1 md:grid-cols-6 gap-2">
                        <div>
                          <label className="block text-[10px] font-bold text-slate-600 mb-0.5">Source *</label>
                          <input 
                            type="text" 
                            required 
                            className="w-full border p-1.5 rounded bg-white text-xs font-mono" 
                            value={flow.source} 
                            onChange={e => updateNetworkFlow(flow.id, 'source', e.target.value)} 
                          />
                        </div>
                        <div>
                          <label className="block text-[10px] font-bold text-slate-600 mb-0.5">Destination *</label>
                          <input 
                            type="text" 
                            required 
                            className="w-full border p-1.5 rounded bg-white text-xs font-mono" 
                            value={flow.destination} 
                            onChange={e => updateNetworkFlow(flow.id, 'destination', e.target.value)} 
                          />
                        </div>
                        <div>
                          <label className="block text-[10px] font-bold text-slate-600 mb-0.5">Service *</label>
                          <input 
                            type="text" 
                            required 
                            className="w-full border p-1.5 rounded bg-white text-xs" 
                            value={flow.service} 
                            onChange={e => updateNetworkFlow(flow.id, 'service', e.target.value)} 
                          />
                        </div>
                        <div>
                          <label className="block text-[10px] font-bold text-slate-600 mb-0.5">Ports *</label>
                          <input 
                            type="text" 
                            required 
                            className="w-full border p-1.5 rounded bg-white text-xs font-mono" 
                            value={flow.port} 
                            onChange={e => updateNetworkFlow(flow.id, 'port', e.target.value)} 
                          />
                        </div>
                        <div>
                          <label className="block text-[10px] font-bold text-slate-600 mb-0.5">Type de flux *</label>
                          <select 
                            className="w-full border p-1.5 rounded bg-white text-xs" 
                            value={flow.flow_type} 
                            onChange={e => updateNetworkFlow(flow.id, 'flow_type', e.target.value)}
                          >
                            <option value="Flux applicatif Web">Flux applicatif Web</option>
                            <option value="Flux d'administration">Flux d'administration</option>
                            <option value="Flux de base de données">Flux de base de données</option>
                          </select>
                        </div>
                        <div>
                          <label className="block text-[10px] font-bold text-slate-600 mb-0.5">Description</label>
                          <input 
                            type="text" 
                            className="w-full border p-1.5 rounded bg-white text-xs" 
                            value={flow.description} 
                            onChange={e => updateNetworkFlow(flow.id, 'description', e.target.value)} 
                          />
                        </div>
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            </div>
          )}

          {/* TAB 4 : SÉCURITÉ SI */}
          {activeTab === 4 && (
            <div className="space-y-5">
              <div className="bg-white p-5 rounded-xl border border-slate-200 shadow-sm space-y-4">
                <div className="border-b pb-2">
                  <h4 className="font-bold text-sm text-slate-800">Service Publié (Paramètres de Sécurité SI)</h4>
                </div>

                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  <div>
                    <label className="block font-semibold text-slate-500 mb-1">Entrée DNS (site web)</label>
                    <input 
                      type="text" 
                      className="w-full border p-2.5 rounded-lg bg-slate-50 text-xs outline-none"
                      value={securityParams.dns_site_web}
                      onChange={e => setSecurityParams({...securityParams, dns_site_web: e.target.value})}
                    />
                  </div>

                  <div>
                    <label className="block font-semibold text-slate-500 mb-1">Adresse IP Publique</label>
                    <input 
                      type="text" 
                      className="w-full border p-2.5 rounded-lg bg-slate-50 text-xs outline-none"
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
                      className="w-full border p-2.5 rounded-lg bg-white text-xs font-mono"
                      value={securityParams.ip_interne}
                      onChange={e => setSecurityParams({...securityParams, ip_interne: e.target.value})}
                    />
                  </div>

                  <div>
                    <label className="block font-semibold text-slate-500 mb-1">Adresse IP Virtuelle F5</label>
                    <input 
                      type="text" 
                      className="w-full border p-2.5 rounded-lg bg-slate-50 text-xs outline-none"
                      value={securityParams.ip_virtuelle || ''}
                      onChange={e => setSecurityParams({...securityParams, ip_virtuelle: e.target.value})}
                    />
                  </div>

                  <div>
                    <label className="block font-semibold text-slate-700 mb-1">
                      Publication <span className="text-red-500 font-bold">*</span>
                    </label>
                    <select 
                      required 
                      className="w-full border p-2.5 rounded-lg bg-white text-xs font-bold"
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
                      className="w-full border p-2.5 rounded-lg bg-white text-xs text-amber-700 font-bold"
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
                    <div key={ctrl.id || `ctrl-${idx}`} className="bg-slate-50 p-3 rounded-lg border border-slate-200 flex flex-col md:flex-row md:items-center justify-between gap-3">
                      <div className="flex-1">
                        <span className="font-bold text-amber-700 mr-2">#{idx + 1}</span>
                        <span className="font-semibold text-slate-800 text-xs">{ctrl.control_name}</span>
                      </div>

                      <div className="flex items-center space-x-2">
                        <select 
                          className={`border p-1.5 rounded-md font-bold text-xs bg-white ${
                            (ctrl.status || '').startsWith('Conforme') ? 'text-emerald-700 border-emerald-300' :
                            ctrl.status === 'Non validé' ? 'text-red-600 border-red-300' : 'text-slate-600'
                          }`}
                          value={ctrl.status}
                          onChange={e => {
                            const val = e.target.value;
                            setSecurityCompliance(prev =>
                              prev.map((item, i) => (i === idx ? { ...item, status: val } : item))
                            );
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
                          placeholder="Commentaires"
                          className="border p-1.5 rounded-md text-xs w-64 bg-white"
                          value={ctrl.comments}
                          onChange={e => {
                            const val = e.target.value;
                            setSecurityCompliance(prev =>
                              prev.map((item, i) => (i === idx ? { ...item, comments: val } : item))
                            );
                          }}
                        />
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            </div>
          )}

          {/* Footer Actions */}
          <div className="pt-4 border-t flex justify-between items-center bg-white p-4 -mx-6 -mb-6 mt-4 shrink-0">
            <div className="flex items-center space-x-2">
              <button 
                type="button" 
                onClick={onClose} 
                className="px-4 py-2 border rounded-lg font-semibold text-xs text-slate-600 hover:bg-slate-100 transition"
              >
                Annuler
              </button>

              <button 
                type="button" 
                onClick={() => exportVMToExcel(getPayload())} 
                className="px-4 py-2 bg-emerald-700 hover:bg-emerald-800 text-white font-bold rounded-lg text-xs shadow transition flex items-center gap-1.5"
              >
                <span>📊</span> Exporter Excel
              </button>
            </div>
            
            <div className="flex items-center space-x-2">
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
                  {loading ? 'Enregistrement...' : 'Enregistrer la Fiche'}
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