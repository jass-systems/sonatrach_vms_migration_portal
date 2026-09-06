import React, { useEffect, useState } from 'react';

export default function StatsDashboard() {
  const [stats, setStats] = useState({
    totalRequests: 8,
    totalVMs: 24,
    osDistribution: [
      { os_server: 'Windows Server 2022', count: 14 },
      { os_server: 'Ubuntu Server 24.04', count: 7 },
      { os_server: 'RHEL 9', count: 3 }
    ],
    stackDistribution: [
      { software_name: 'Python / Django', count: 12 },
      { software_name: 'Node.js', count: 8 },
      { software_name: 'Oracle DB', count: 5 },
      { software_name: 'Nginx', count: 18 }
    ]
  });

  useEffect(() => {
    fetch('http://localhost:5000/api/stats')
      .then(res => res.json())
      .then(data => setStats(data))
      .catch(() => {});
  }, []);

  return (
    <div className="space-y-6">
      <div>
        <h2 className="text-2xl font-bold text-slate-800">Tableau de Bord & Statistiques</h2>
        <p className="text-xs text-slate-500">Synthèse globale du parc applicatif et des composants VM</p>
      </div>

      {/* Cartes Métriques (KPIs) */}
      <div className="grid grid-cols-3 gap-6">
        <div className="bg-white p-6 rounded-xl border border-slate-200 shadow-sm">
          <span className="text-xs font-semibold text-slate-400 uppercase">Demandes de Migration</span>
          <div className="text-3xl font-extrabold text-slate-900 mt-2">{stats.totalRequests}</div>
        </div>
        <div className="bg-white p-6 rounded-xl border border-slate-200 shadow-sm">
          <span className="text-xs font-semibold text-slate-400 uppercase">Total VMs Enregistrées</span>
          <div className="text-3xl font-extrabold text-amber-600 mt-2">{stats.totalVMs}</div>
        </div>
        <div className="bg-white p-6 rounded-xl border border-slate-200 shadow-sm">
          <span className="text-xs font-semibold text-slate-400 uppercase">Taux de Conformité SI</span>
          <div className="text-3xl font-extrabold text-emerald-600 mt-2">100%</div>
        </div>
      </div>

      {/* Graphiques et Répartitions */}
      <div className="grid grid-cols-2 gap-6">
        <div className="bg-white p-6 rounded-xl border border-slate-200 shadow-sm">
          <h3 className="font-bold text-slate-800 text-sm mb-4">Répartition par Système d'Exploitation (OS)</h3>
          <div className="space-y-3">
            {stats.osDistribution.map((item, idx) => (
              <div key={idx} className="space-y-1">
                <div className="flex justify-between text-xs font-semibold text-slate-700">
                  <span>{item.os_server}</span>
                  <span>{item.count} VMs</span>
                </div>
                <div className="w-full bg-slate-100 h-2 rounded-full overflow-hidden">
                  <div className="bg-amber-500 h-2 rounded-full" style={{ width: `${(item.count / (stats.totalVMs || 1)) * 100}%` }}></div>
                </div>
              </div>
            ))}
          </div>
        </div>

        <div className="bg-white p-6 rounded-xl border border-slate-200 shadow-sm">
          <h3 className="font-bold text-slate-800 text-sm mb-4">Technologies / Stack Logicielle</h3>
          <div className="space-y-3">
            {stats.stackDistribution.map((item, idx) => (
              <div key={idx} className="space-y-1">
                <div className="flex justify-between text-xs font-semibold text-slate-700">
                  <span>{item.software_name}</span>
                  <span>{item.count} déploiements</span>
                </div>
                <div className="w-full bg-slate-100 h-2 rounded-full overflow-hidden">
                  <div className="bg-slate-800 h-2 rounded-full" style={{ width: `${(item.count / (stats.totalVMs || 1)) * 100}%` }}></div>
                </div>
              </div>
            ))}
          </div>
        </div>
      </div>
    </div>
  );
}