import React, { useState } from 'react';
import Login from './components/Login';
import Dashboard from './components/Dashboard';

export default function App() {
  const [user, setUser] = useState(null);

  if (!user) {
    return <Login onLogin={(userData) => setUser(userData)} />;
  }

  return (
    <div className="min-h-screen bg-slate-50 flex flex-col">
      {/* Barre de navigation principale avec z-index ajusté (z-10) */}
      <header className="bg-white border-b border-slate-200 sticky top-0 z-10 shadow-sm">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
          <div className="flex justify-between items-center h-16">
            
            {/* Logo & Titre */}
            <div className="flex items-center space-x-3">
              <img 
                src="/logo.png" 
                alt="Sonatrach" 
                className="h-10 w-auto object-contain" 
              />
              <div className="h-6 w-px bg-slate-300"></div>
              <div>
                <h1 className="text-base font-bold text-slate-800 leading-tight">Portail de Migration VM</h1>
                <p className="text-[10px] text-amber-700 font-semibold tracking-wider uppercase">Sonatrach TRC</p>
              </div>
            </div>

            {/* Profil Utilisateur & Déconnexion */}
            <div className="flex items-center space-x-4">
              <div className="text-right">
                <p className="text-xs font-bold text-slate-800">{user.username}</p>
                <p className="text-[10px] text-slate-500 capitalize">{user.role || 'Utilisateur'}</p>
              </div>
              <button 
                onClick={() => setUser(null)}
                className="text-xs bg-slate-100 hover:bg-slate-200 text-slate-700 px-3 py-1.5 rounded-lg border border-slate-300 font-medium transition"
              >
                Déconnexion
              </button>
            </div>

          </div>
        </div>
      </header>

      {/* Contenu principal */}
      <main className="flex-1 max-w-7xl w-full mx-auto p-4 sm:p-6 lg:p-8">
        <Dashboard />
      </main>
    </div>
  );
}