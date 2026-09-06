-- schema.sql

-- Table Utilisateurs (Authentification)
CREATE TABLE IF NOT EXISTS users (
    id SERIAL PRIMARY KEY,
    username VARCHAR(50) UNIQUE NOT NULL,
    password VARCHAR(255) NOT NULL,
    role VARCHAR(20) DEFAULT 'user',
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

-- Table Demandes de Migration
CREATE TABLE IF NOT EXISTS migration_requests (
    id SERIAL PRIMARY KEY,
    title VARCHAR(255) NOT NULL,
    pole VARCHAR(100) NOT NULL,
    structure VARCHAR(100) NOT NULL,
    responsable_structure VARCHAR(100) NOT NULL,
    responsable_service VARCHAR(100) NOT NULL,
    contact VARCHAR(100) NOT NULL,
    status VARCHAR(50) DEFAULT 'Brouillon',
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

-- Table Fiches VM (Onglets 1, 2, 3)
CREATE TABLE IF NOT EXISTS vms (
    id SERIAL PRIMARY KEY,
    migration_request_id INTEGER REFERENCES migration_requests(id) ON DELETE CASCADE,
    app_name VARCHAR(150) NOT NULL,
    publication_type VARCHAR(100) NOT NULL,
    target_population TEXT,
    dns_entry VARCHAR(150) NOT NULL,
    ip_address VARCHAR(50) NOT NULL,
    port VARCHAR(20) NOT NULL,
    os_server VARCHAR(100) NOT NULL,
    architecture_desc TEXT,
    public_ip VARCHAR(50),
    f5_virtual_ip VARCHAR(50),
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

-- Table Software Stack (Composants logiciels de la VM)
CREATE TABLE IF NOT EXISTS vm_software_stack (
    id SERIAL PRIMARY KEY,
    vm_id INTEGER REFERENCES vms(id) ON DELETE CASCADE,
    software_name VARCHAR(100) NOT NULL,
    exists BOOLEAN DEFAULT FALSE,
    version VARCHAR(50)
);

-- Table Matrice des Flux Réseau (Onglet 3)
CREATE TABLE IF NOT EXISTS network_flows (
    id SERIAL PRIMARY KEY,
    vm_id INTEGER REFERENCES vms(id) ON DELETE CASCADE,
    source VARCHAR(100) NOT NULL,
    destination VARCHAR(100) NOT NULL,
    service VARCHAR(50) NOT NULL,
    port VARCHAR(20) NOT NULL,
    flow_type VARCHAR(100) NOT NULL,
    description TEXT
);

-- Table Suivi des Non-Conformités (Onglet 4 - Sécurité SI)
CREATE TABLE IF NOT EXISTS security_compliance (
    id SERIAL PRIMARY KEY,
    vm_id INTEGER REFERENCES vms(id) ON DELETE CASCADE,
    control_name VARCHAR(255) NOT NULL,
    status VARCHAR(50) DEFAULT 'Non Conforme',
    comments TEXT
);

-- Utilisateur par défaut pour le test (admin / admin123)
INSERT INTO users (username, password, role) 
VALUES ('admin', 'admin123', 'admin') 
ON CONFLICT (username) DO NOTHING;