# 🛡️ Portail de Publication & Sécurisation de Machines Virtuelles — SONATRACH TRC
Plateforme web full-stack développée pour la **Direction Technologies de l'Information (DTI) de SONATRACH TRC**. Ce système automatise la collecte des fiches techniques de machines virtuelles (VM), génère dynamiquement la matrice de flux réseau, contrôle la conformité aux exigences PSSI (Sécurité SI) et exporte des classeurs `.xlsx` multi-onglets prêts pour l'exploitation.

---

## 📌 Table des Matières
[Fonctionnalités Clés](#-fonctionnalités-clés)
[Prérequis](#-prérequis)
[Installation & Configuration](#-installation--configuration)
   * [1. Base de Données (PostgreSQL / Docker)](#1-base-de-données-postgresql--docker)
   * [2. Serveur Backend (Node.js / Express)](#2-serveur-backend-nodejs--express)
   * [3. Application Frontend (React.js)](#3-application-frontend-reactjs)
[Structure du Projet](#-structure-du-projet)
[Génération du Schéma & Export Excel](#-génération-du-schéma--export-excel)
[Auteur & Contexte Académique](#-auteur--contexte-académique)

---

## ✨ Fonctionnalités Clés

* **Assistant Guidé en 4 Étapes (Wizard Form)** : Saisie structurée des informations demandeur, paramètres réseau (IP, DNS, Port), stack logicielle et contrôles PSSI.
* **Moteur d'Inférence des Flux Réseau** : Génération et recommandation automatique des règles et ports (SSH 22, Oracle 1521, PostgreSQL 5432, HTTPS 443) selon la stack technique choisie.
* **Génération Dynamique du Diagramme d'Architecture** : Rendu HTML5 Canvas haute définition intégrant les données réelles (ports, services, noms de VM) directement converti en image dans le fichier Excel.
* **Exportation Complète `.xlsx` via ExcelJS** : Injection automatique des données dans les 4 onglets réglementaires de SONATRACH (`Principale`, `Publication VM`, `Informations liées au service`, `Suivi des Non conformités`).
* **Module d'Audit Sécurité SI & RBAC** : Espace réservé aux auditeurs pour la validation des 8 points de conformité PSSI (DMZ, Antivirus, TLS, PoLP, Scans de vulnérabilités) avec verrouillage des accès demandeurs.

---

## 🛠️ Prérequis

Assurez-vous d'avoir installé les éléments suivants sur votre machine :

* **Node.js** : Version `>= 18.x`
* **npm** ou **yarn**
* **PostgreSQL** (v14+) OU **Docker Desktop**
* **Git**

---

## 🚀 Installation & Configuration

### 1. Base de Données (PostgreSQL / Docker)

Vous pouvez lancer PostgreSQL rapidement à l'aide de Docker :

```bash
docker run --name sonatrach-db \
  -e POSTGRES_USER=postgres \
  -e POSTGRES_PASSWORD=postgres \
  -e POSTGRES_DB=sonatrach_vm_db \
  -p 5400:5432 \
  -d postgres:15

2. Serveur Backend (Node.js / Express)
     Naviguez dans le répertoire backend : Bash : cd backend
     2.1. Installez les dépendances : Bash : npm install

     2.2. Créez un fichier .env à la racine du dossier backend :
     PORT=5000
     DB_HOST=localhost
     DB_PORT=5400
     DB_USER=postgres
     DB_PASSWORD=postgres
     DB_NAME=sonatrach_vm_db
     JWT_SECRET=sonatrach_trc_secret_key_2026

     2.3. Démarrez le serveur backend :

     Bash
     npm start
     # ou en mode développement :
     npm run dev

3. Application Frontend (React.js)
     3.1. Naviguez dans le répertoire frontend : Bash : cd frontend
     3.2. Installez les dépendances : Bash : npm install
     3.3. Assets Requis pour l'Export Excel :
            Assurez-vous que les images et icônes suivantes sont présentes dans le dossier frontend/public/ pour permettre la génération dynamique des diagrammes et logos :
            logo.png (Logo SONATRACH)
            principale.png (Bannière officielle)
            img_0.png (Icône Cloud / WAN)
            img_1.png (Icône Base de données)
            img_2.png (Icône Administrateur interne)
      3.4. Lancer l'application React : Bash : npm run dev
      # L'application sera accessible sur http://localhost:5173

📂 Structure du Projet
sonatrach_vms_migration_portal/
│
├── client/                     # Application Frontend React (Vite + Tailwind CSS)
│   ├── public/                 # Assets statiques (logos, images du schéma d'architecture)
│   │   ├── logo.png
│   │   ├── principale.png
│   │   ├── img_0.png           # Icône Cloud / WAN
│   │   ├── img_1.png           # Icône Base de données
│   │   └── img_2.png           # Icône Administrateur interne
│   ├── src/
│   │   ├── components/         # Composants IHM (Dashboard, Modaux, Formulaires, Auth)
│   │   │   ├── Dashboard.jsx
│   │   │   ├── Login.jsx
│   │   │   ├── RequestDetailsModal.jsx
│   │   │   ├── StatsDashboard.jsx
│   │   │   └── VMFormModal.jsx
│   │   ├── utils/              # Moteur d'exportation ExcelJS & Générateur Canvas
│   │   │   └── exportVMExcel.js
│   │   ├── App.jsx             # Composant racine de l'application
│   │   ├── index.css           # Styles globaux et Tailwind Directives
│   │   └── main.jsx            # Point d'entrée React
│   ├── index.html              # Document HTML principal
│   ├── package.json
│   ├── postcss.config.js       # Configuration PostCSS pour Tailwind
│   ├── tailwind.config.js      # Configuration Tailwind CSS
│   └── vite.config.js          # Configuration du bundler Vite
│
└── server/                     # Serveur Backend Node.js / Express
    ├── routes/                 # Routes API REST (Auth, Migrations, VMs, Stats)
    ├── db.js                   # Module de connexion PostgreSQL (pg Pool)
    ├── formulaire_création_VM.xlsx  # Modèle officiel de référence SONATRACH
    ├── inspect.js              # Script utilitaire d'inspection
    ├── package.json
    └── server.js               # Point d'entrée principal du serveur API


📊 Génération du Schéma & Export Excel
Le module exportVMExcel.js prend en charge :
L'extraction dynamique des données de la VM et de sa matrice des flux (networkFlows).
Le rendu temps réel d'un schéma d'architecture 2D sur un Canvas HTML5 (WAN, VM, Base de données, Administrateur Interne).
L'injection de l'image générée dans l'onglet Informations liées au service du fichier Excel.
La création d'un document téléchargeable respectant la charte graphique officielle de SONATRACH TRC (couleurs institutionnelles #ED7D31, #4F81BD).

👨‍🎓 Auteur & Contexte Académique
Projet : Projet de Fin d'Études (SPE / Master) — Promotion 2026.
Élève Ingénieur : HEMAIZIA Abderrahmene (Option : Systèmes Intelligents et Données - SD).
Établissement : École Nationale Supérieure d'Informatique (ESI), Alger.
Organisme d'Accueil : SONATRACH TRC — Direction Technologies de l'Information (DTI).
Encadrants : Mme BARKAT Siham (SONATRACH TRC) & Mme OUFAIDA Houda (ESI).






