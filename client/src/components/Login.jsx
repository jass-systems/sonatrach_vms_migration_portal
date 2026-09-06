import React, { useState } from 'react';

export default function Login({ onLogin }) {
  const [username, setUsername] = useState('');
  const [password, setPassword] = useState('');
  const [error, setError] = useState('');
  const [loading, setLoading] = useState(false);

  const handleSubmit = async (e) => {
    e.preventDefault();
    setError('');
    setLoading(true);

    try {
      const res = await fetch('http://localhost:5000/api/auth/login', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ username, password })
      });
      
      const data = await res.json();

      if (res.ok && data.success) {
        onLogin(data.user);
      } else {
        setError(data.message || 'Identifiants incorrects');
      }
    } catch {
      setError('Impossible de contacter le serveur backend');
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="min-h-screen bg-slate-100 flex items-center justify-center p-4">
      <div className="bg-white p-8 rounded-2xl shadow-xl w-full max-w-md border border-slate-200">
        <div className="flex justify-center mb-6">
          <img 
            src="/logo.png" 
            alt="Sonatrach TRC" 
            className="h-16 object-contain" 
          />
        </div>

        <h2 className="text-xl font-bold text-center text-slate-800 mb-1">Connexion Portail VM</h2>
        <p className="text-xs text-center text-slate-500 mb-6">Sonatrach TRC - Direction Centrale IT</p>

        {error && (
          <div className="mb-4 text-xs bg-red-50 text-red-600 p-3 rounded-lg border border-red-200 text-center font-medium">
            {error}
          </div>
        )}

        <form onSubmit={handleSubmit} className="space-y-4">
          <div>
            <label className="block text-xs font-semibold text-slate-700 mb-1">Nom d'utilisateur</label>
            <input 
              type="text" 
              className="w-full px-3 py-2 border rounded-lg text-sm focus:ring-2 focus:ring-amber-500 outline-none transition" 
              value={username} 
              onChange={(e) => setUsername(e.target.value)}
              placeholder="Nom d'utilisateur"
              required 
            />
          </div>
          <div>
            <label className="block text-xs font-semibold text-slate-700 mb-1">Mot de passe</label>
            <input 
              type="password" 
              className="w-full px-3 py-2 border rounded-lg text-sm focus:ring-2 focus:ring-amber-500 outline-none transition" 
              value={password} 
              onChange={(e) => setPassword(e.target.value)}
              placeholder="••••••••"
              required 
            />
          </div>
          <button 
            type="submit" 
            disabled={loading}
            className="w-full bg-amber-600 hover:bg-amber-700 disabled:bg-amber-400 text-white font-bold py-2.5 rounded-lg text-sm transition shadow-md"
          >
            {loading ? 'Vérification...' : 'Se Connecter'}
          </button>
        </form>
      </div>
    </div>
  );
}