const express = require('express');
const cors = require('cors');

// Importation des sous-modules de routes
const authRoutes = require('./routes/auth');
const migrationRoutes = require('./routes/migrations');
const vmRoutes = require('./routes/vms'); // <-- Prise en charge des actions directes sur les VMs (/api/vms)
const exportRoutes = require('./routes/export');
const statsRoutes = require('./routes/stats');

const app = express();

// Middlewares
app.use(cors());
app.use(express.json({ limit: '10mb' }));
app.use(express.urlencoded({ extended: true, limit: '10mb' }));

// Enregistrement des sous-modules de routes API
app.use('/api/auth', authRoutes);
app.use('/api/migrations', migrationRoutes);
app.use('/api/vms', vmRoutes); // <-- Permet de gérer PUT /api/vms/:id et DELETE /api/vms/:id
app.use('/api', exportRoutes);
app.use('/api/stats', statsRoutes);

// Route de test de santé API
app.get('/health', (req, res) => {
  res.status(200).json({ status: 'OK', message: 'Serveur Sonatrach TRC Migration VM opérationnel.' });
});

// Middleware d'erreur globale
app.use((err, req, res, next) => {
  console.error('Erreur Serveur :', err.stack);
  res.status(500).json({ message: 'Erreur interne du serveur', error: err.message });
});

const PORT = process.env.PORT || 5000;
app.listen(PORT, () => {
  console.log(`Serveur Sonatrach TRC actif et prêt sur http://localhost:${PORT}`);
});