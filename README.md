# 🔥 JSP St Cyprien — Suivi d'Activité

Application Google Apps Script pour le suivi des présences et absences des Jeunes Sapeurs-Pompiers (JSP) de St Cyprien.

## 📋 Fonctionnalités

### Espace JSP
- Calendrier 3 mois (précédent, courant, suivant) avec navigation
- Marqueurs bleus sur les jours avec événements
- Détail d'événement au clic + tooltip au survol
- Signalement d'absence avec motif
- Bilan personnel : absences totales, signalées, non signalées, taux

### Espace Responsable
- Calendrier mensuel avec navigation
- Création d'événements (nom, horaires, lieu, sections)
- Prise de présences le jour de l'événement ("Cocher tout" + décocher les absents)
- Bilan global par section avec tri par colonne

## 🗂️ Structure du Spreadsheet

| Onglet | Colonnes |
|--------|----------|
| **Liste JSP** | Identité, Login, Mot de passe, Section (dropdown) |
| **Sections** | Section, Référent(s) |
| **Référents** | Identité, Login, Mot de passe |
| **Événements** | ID, Date, Nom, Heure Début, Heure Fin, Lieu, Sections, Créé par |
| **Présences** | EventID, Date, Login JSP, Nom JSP, Section, Présent, Absence Signalée, Motif |

## 🚀 Installation

### 1. Créer le projet Apps Script
```bash
cd jsp-st-cyp
npx @google/clasp create --type webapp --title "JSP St Cyp"
npx @google/clasp push --force
```

### 2. Configurer le Spreadsheet
- Dans l'éditeur Apps Script, exécuter `setupSpreadsheet()`
- Copier l'ID du spreadsheet affiché dans les logs
- Coller dans `Config.js` > `SPREADSHEET_ID`
- `npx @google/clasp push --force`

### 3. Déployer
```bash
npx @google/clasp deploy -d "v1"
```

## 🔑 Identifiants de test

| Rôle | Login | Mot de passe |
|------|-------|-------------|
| Responsable | 66000 | 66000 |
| JSP | lmartin | 66000 |
| JSP | edupont | 66000 |
| JSP | hbernard | 66000 |
| ... | ... | 66000 |
