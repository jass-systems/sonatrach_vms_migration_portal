import React, { useState, useEffect } from 'react';
import { createPortal } from 'react-dom';
import VMFormModal from './VMFormModal';
import { exportVMToExcel } from '../utils/exportVMExcel';

export default function RequestDetailsModal({ request, onClose, onRefresh }) {
  const [vms, setVms] = useState([]);
  const [loading, setLoading] = useState(true);
  const [selectedVmForEdit, setSelectedVmForEdit] = useState(null);
  const [showAddVmModal, setShowAddVmModal] = useState(false);

  // Charger la liste des VMs associées à cette demande
  const fetchVms = async () => {
    setLoading(true);
    try {
      const res = await fetch(`http://localhost:5000/api/migrations/${request.id}/vms`);
      if (res.ok) {
        const data = await res.json();
        setVms(data);
      }
    } catch (err) {
      console.error("Erreur chargement VMs:", err);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    if (request?.id) fetchVms();
  }, [request?.id]);

  // SUPPRESSION D'UNE VM SPÉCIFIQUE
  const handleDeleteVM = async (vm) => {
    const vmId = vm.id || vm._id;
    const vmIp = vm.formPublication?.ip_address || vm.ip_address || vm.app_name || 'cette VM';

    if (!vmId) {
      alert("Erreur: ID de la machine virtuelle introuvable.");
      return;
    }

    if (!window.confirm(`Voulez-vous vraiment supprimer la machine virtuelle (${vmIp}) ?`)) return;

    try {
      let res = await fetch(`http://localhost:5000/api/vms/${vmId}`, { method: 'DELETE' });

      if (res.status === 404) {
        res = await fetch(`http://localhost:5000/api/migrations/vms/${vmId}`, { method: 'DELETE' });
      }

      if (res.ok) {
        fetchVms();
        if (onRefresh) onRefresh();
      } else {
        const errData = await res.json().catch(() => ({}));
        alert(`Erreur lors de la suppression : ${errData.message || res.statusText}`);
      }
    } catch (err) {
      console.error("Erreur lors de la suppression de la VM:", err);
      alert("Erreur réseau lors de la suppression de la VM.");
    }
  };

  if (!request) return null;

  return createPortal(
    <div className="fixed inset-0 bg-black/80 flex justify-center items-center z-[9000] p-4 backdrop-blur-sm overflow-y-auto">
      <div className="bg-white rounded-2xl shadow-2xl w-full max-w-6xl border border-slate-200 overflow-hidden flex flex-col max-h-[90vh] my-auto">
        
        {/* Entête Demande */}
        <div className="bg-slate-900 text-white px-6 py-4 flex justify-between items-center border-b border-slate-800 shrink-0">
          <div>
            <div className="flex items-center gap-3">
              <span className="text-xs font-bold px-2.5 py-1 bg-amber-500 text-slate-950 rounded-md">
                Demande {request.id}
              </span>
              <h3 className="font-bold text-lg">{request.title}</h3>
            </div>
            <p className="text-xs text-slate-400 mt-1">
              Structure : <strong className="text-white">{request.structure || 'DTI'}</strong> | 
              Pôle : <strong className="text-white">{request.pole || 'ALGER'}</strong>
            </p>
          </div>
          <button 
            onClick={onClose}
            className="text-slate-400 hover:text-white font-bold text-xl h-9 w-9 rounded-lg hover:bg-slate-800 flex items-center justify-center transition"
          >
            ✕
          </button>
        </div>

        {/* Corps : Liste des VMs */}
        <div className="p-6 overflow-y-auto flex-1 space-y-6 bg-slate-50/50">
          
          <div className="flex justify-between items-center bg-white p-4 rounded-xl border border-slate-200 shadow-sm">
            <div>
              <h4 className="font-bold text-sm text-slate-800 flex items-center gap-2">
                <span>🖥️</span> Machines Virtuelles de la demande ({vms.length})
              </h4>
              <p className="text-xs text-slate-500">Consultez, modifiez, exportez en Excel ou supprimez les machines de cette demande.</p>
            </div>
            <button
              onClick={() => setShowAddVmModal(true)}
              className="px-4 py-2 bg-amber-600 hover:bg-amber-700 text-white font-bold text-xs rounded-lg shadow transition flex items-center gap-2"
            >
              <span>+</span> Ajouter une VM
            </button>
          </div>

          {/* Tableau des VMs */}
          <div className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden">
            {loading ? (
              <div className="p-8 text-center text-xs text-slate-500">Chargement des machines virtuelles...</div>
            ) : vms.length === 0 ? (
              <div className="p-8 text-center space-y-3">
                <p className="text-slate-500 text-xs">Aucune machine virtuelle n'est rattachée à cette demande.</p>
                <button
                  onClick={() => setShowAddVmModal(true)}
                  className="px-4 py-2 bg-slate-900 text-white font-bold text-xs rounded-lg"
                >
                  Ajouter la première VM
                </button>
              </div>
            ) : (
              <table className="w-full text-left border-collapse text-xs">
                <thead>
                  <tr className="bg-slate-100 border-b border-slate-200 text-slate-700 font-bold">
                    <th className="p-3">N°</th>
                    <th className="p-3">Nom VM / Application</th>
                    <th className="p-3">Adresse IP</th>
                    <th className="p-3">Port</th>
                    <th className="p-3">OS Serveur</th>
                    <th className="p-3 text-right">Actions</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-200 font-medium">
                  {vms.map((vm, index) => {
                    const ip = vm.formPublication?.ip_address || vm.ip_address || '—';
                    const appName = vm.formPublication?.app_name || vm.app_name || 'VM';
                    const port = vm.formPublication?.port || vm.port || '443';
                    const osServer = vm.formPublication?.os_server || vm.os_server || 'Windows Server';

                    return (
                      <tr key={vm.id || vm._id || index} className="hover:bg-slate-50/80 transition">
                        <td className="p-3 font-bold text-slate-500">{index + 1}</td>
                        <td className="p-3 font-bold text-slate-900">{appName}</td>
                        <td className="p-3 font-mono text-amber-700 font-bold">{ip}</td>
                        <td className="p-3 font-mono text-slate-600">{port}</td>
                        <td className="p-3 text-slate-700">{osServer}</td>
                        <td className="p-3 text-right space-x-2">
                          {/* Voir / Éditer */}
                          <button
                            onClick={() => setSelectedVmForEdit(vm)}
                            className="px-2.5 py-1.5 bg-slate-100 hover:bg-slate-200 text-slate-800 font-bold rounded-lg border border-slate-300 transition"
                            title="Modifier cette VM"
                          >
                            👁️ Éditer
                          </button>

                          {/* Export Excel */}
                          <button
                            onClick={() => exportVMToExcel(vm)}
                            className="px-2.5 py-1.5 bg-emerald-600 hover:bg-emerald-700 text-white font-bold rounded-lg shadow transition inline-flex items-center gap-1"
                            title="Télécharger le modèle Excel de cette VM"
                          >
                            <span>📊</span> Export Excel
                          </button>

                          {/* SUPPRESSION VM */}
                          <button
                            onClick={() => handleDeleteVM(vm)}
                            className="px-2.5 py-1.5 bg-red-50 hover:bg-red-100 text-red-600 border border-red-200 font-bold rounded-lg transition"
                            title="Supprimer cette VM"
                          >
                            🗑️ Supprimer
                          </button>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            )}
          </div>

        </div>

        <div className="p-4 bg-white border-t flex justify-end shrink-0">
          <button
            onClick={onClose}
            className="px-5 py-2 bg-slate-900 text-white font-bold rounded-lg text-xs hover:bg-slate-800 transition"
          >
            Fermer
          </button>
        </div>

      </div>

      {/* Modal d'ajout d'une VM */}
      {showAddVmModal && (
        <VMFormModal
          requestId={request.id}
          onClose={() => setShowAddVmModal(false)}
          onSuccess={() => {
            setShowAddVmModal(false);
            fetchVms();
            if (onRefresh) onRefresh();
          }}
        />
      )}

      {/* Modal d'édition d'une VM */}
      {selectedVmForEdit && (
        <VMFormModal
          requestId={request.id}
          initialData={selectedVmForEdit}
          onClose={() => setSelectedVmForEdit(null)}
          onSuccess={() => {
            setSelectedVmForEdit(null);
            fetchVms();
            if (onRefresh) onRefresh();
          }}
        />
      )}
    </div>,
    document.body
  );
}