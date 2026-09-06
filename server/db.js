const { Pool } = require('pg');

const pool = new Pool({
  user: 'postgres',
  host: 'localhost',
  database: 'sonatrach_db', // <-- Doit correspondre exactement au nom de la base créée
  password: '72741405',    // Votre mot de passe de conteneur
  port: 5432,
});

module.exports = pool;