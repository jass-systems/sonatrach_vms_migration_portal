import React, { useState, useEffect } from 'react';
import RequestDetailsModal from './RequestDetailsModal';
import VMFormModal from './VMFormModal';

export default function Dashboard() {
  const [requests, setRequests] = useState([]);
  const [showNewRequest, setShowNewRequest] = useState(false);
  const [selectedRequestForDetails, setSelectedRequestForDetails] = useState(null);
  const [selectedReqForNewVM, setSelectedReqForNewVM] = useState(null);
  const [title, setTitle] = useState('');
  const [loading, setLoading] = useState(false);

  // Charger la liste des demandes
  const loadRequests = async () => {
    setLoading(true);
    try {
      const res = await fetch('http://localhost:5000/api/migrations');
      if (res.ok) {
        const data = await res.json();
        setRequests(data);
      }
    } catch (err) {
      console.error("Erreur lors du chargement des demandes :", err);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => { 
    loadRequests(); 
  }, []);

  // Créer une nouvelle demande globale
  const handleCreateRequest = async (e) => {
    e.preventDefault();
    try {
      const res = await fetch('http://localhost:5000/api/migrations', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ title })
      });
      if (res.ok) {
        setTitle('');
        setShowNewRequest(false);
        loadRequests();
      } else {
        alert("Erreur lors de l'initialisation de la demande");
      }
    } catch (err) {
      console.error("Erreur lors de la création :", err);
    }
  };

  // Supprimer une demande
  const handleDeleteRequest = async (id) => {
    if (!window.confirm('Voulez-vous vraiment supprimer cette demande et toutes ses VMs associées ?')) return;
    try {
      const res = await fetch(`http://localhost:5000/api/migrations/${id}`, { method: 'DELETE' });
      if (res.ok) {
        loadRequests();
      }
    } catch (err) {
      console.error("Erreur lors de la suppression :", err);
    }
  };

  return (
    <div className="w-full space-y-6">
      
      {/* En-tête de la page responsive */}
      <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4 bg-white p-5 sm:p-6 rounded-2xl border border-slate-200 shadow-sm w-full">
        <div>
          <h2 className="text-xl font-bold text-slate-800 flex items-center gap-2">
            <span>📦</span> Demandes de Migration Sonatrach TRC
          </h2>
          <p className="text-xs text-slate-500 mt-1">
            Gérez vos demandes et accédez au détail des machines virtuelles (VMs) pour l'exportation Excel individuelle (.xlsx).
          </p>
        </div>
        <button 
          onClick={() => setShowNewRequest(true)}
          className="bg-amber-600 hover:bg-amber-700 text-white text-xs font-bold px-4 py-2.5 rounded-lg transition shadow-sm flex items-center gap-2 whitespace-nowrap shrink-0 w-full sm:w-auto justify-center"
        >
          <span>+</span> Initialiser une demande
        </button>
      </div>

      {/* Modal : Création d'une nouvelle Demande */}
      {showNewRequest && (
        <div className="fixed inset-0 bg-black/60 flex justify-center items-center z-50 p-4 backdrop-blur-sm">
          <div className="bg-white p-6 rounded-2xl shadow-2xl w-full max-w-md border border-slate-200">
            <h3 className="font-bold text-slate-800 text-sm mb-4 flex items-center gap-2">
              <span>📝</span> Initialiser une Demande de Migration
            </h3>
            <form onSubmit={handleCreateRequest} className="space-y-4">
              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">
                  Titre / Description de la demande <span className="text-red-500">*</span>
                </label>
                <input 
                  type="text" 
                  required
                  placeholder="ex: Migration Application Intranet TRC 2026" 
                  className="w-full border p-2.5 rounded-lg text-xs outline-none focus:ring-2 focus:ring-amber-500 bg-white"
                  value={title}
                  onChange={e => setTitle(e.target.value)}
                />
              </div>
              <div className="flex justify-end space-x-2 pt-2">
                <button 
                  type="button" 
                  onClick={() => setShowNewRequest(false)} 
                  className="px-4 py-2 text-xs border rounded-lg font-semibold text-slate-600 hover:bg-slate-100 transition"
                >
                  Annuler
                </button>
                <button 
                  type="submit" 
                  className="px-4 py-2 text-xs bg-amber-600 hover:bg-amber-700 text-white font-bold rounded-lg transition shadow"
                >
                  Créer la demande
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* Tableau Principal des Demandes avec Largeur Minimale */}
      <div className="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-x-auto w-full">
        {loading ? (
          <div className="p-8 text-center text-xs text-slate-500">Chargement des demandes de migration...</div>
        ) : requests.length === 0 ? (
          <div className="p-12 text-center space-y-3">
            <p className="text-slate-500 text-xs font-medium">Aucune demande de migration enregistrée.</p>
            <button 
              onClick={() => setShowNewRequest(true)}
              className="px-4 py-2 bg-slate-900 text-white font-bold text-xs rounded-lg hover:bg-slate-800 transition"
            >
              Initialiser la première demande
            </button>
          </div>
        ) : (
          <table className="w-full text-left text-xs min-w-[850px]">
            <thead className="bg-slate-100 text-slate-700 border-b border-slate-200 font-bold uppercase text-[11px]">
              <tr>
                <th className="p-4 whitespace-nowrap w-16">ID</th>
                <th className="p-4 whitespace-nowrap">Titre de la Demande</th>
                <th className="p-4 whitespace-nowrap">Structure / Pôle</th>
                <th className="p-4 whitespace-nowrap">Machines Virtuelles</th>
                <th className="p-4 text-right whitespace-nowrap">Actions</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-200 font-medium">
              {requests.map(req => (
                <tr key={req.id} className="hover:bg-slate-50/80 transition">
                  <td className="p-4 font-bold text-slate-500 whitespace-nowrap">#{req.id}</td>
                  <td className="p-4 font-bold text-slate-900 whitespace-nowrap">{req.title}</td>
                  <td className="p-4 text-slate-600 whitespace-nowrap">
                    <span className="font-semibold text-slate-800">{req.pole || 'ALGER'}</span> - {req.structure || 'TRC Siège / EXP'}
                  </td>
                  <td className="p-4 whitespace-nowrap">
                    <span className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-md text-[11px] font-bold bg-amber-100 text-amber-800 border border-amber-200 whitespace-nowrap">
                      🖥️ {req.vm_count || req.vms?.length || 0} VM(s)
                    </span>
                  </td>
                  
                  {/* Actions alignées sans troncature */}
                  <td className="p-4 text-right whitespace-nowrap">
                    <div className="flex items-center justify-end gap-2">
                      <button 
                        onClick={() => setSelectedRequestForDetails(req)}
                        className="bg-slate-900 hover:bg-slate-800 text-white text-xs px-3 py-2 rounded-lg transition font-bold shadow-sm inline-flex items-center gap-1.5 whitespace-nowrap"
                        title="Voir les détails de la demande et exporter les fiches Excel des VMs"
                      >
                        <span>👁️</span> Voir / Gérer les VMs
                      </button>

                      <button 
                        onClick={() => setSelectedReqForNewVM(req.id)}
                        className="bg-amber-600 hover:bg-amber-700 text-white text-xs px-3 py-2 rounded-lg transition font-bold shadow-sm inline-flex items-center gap-1 whitespace-nowrap"
                        title="Ajouter une machine virtuelle à cette demande"
                      >
                        <span>+</span> Saisir VM
                      </button>

                      <button 
                        onClick={() => handleDeleteRequest(req.id)}
                        className="bg-red-50 hover:bg-red-100 text-red-600 hover:text-red-800 border border-red-200 text-xs p-2 rounded-lg transition font-bold inline-flex items-center justify-center shrink-0"
                        title="Supprimer cette demande"
                      >
                        🗑️
                      </button>
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </div>

      {/* Modaux */}
      {selectedRequestForDetails && (
        <RequestDetailsModal 
          request={selectedRequestForDetails} 
          onClose={() => setSelectedRequestForDetails(null)} 
          onRefresh={loadRequests} 
        />
      )}

      {selectedReqForNewVM && (
        <VMFormModal 
          requestId={selectedReqForNewVM} 
          onClose={() => setSelectedReqForNewVM(null)} 
          onSuccess={() => {
            setSelectedReqForNewVM(null);
            loadRequests();
          }} 
        />
      )}

    </div>
  );
}