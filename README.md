# UManager

Base d'application web pour gérer un large répertoire de personnes avec :

- Catégories ajoutables/éditables.
- Champs obligatoires : **Nom de naissance**, **Prénom**, **Date de naissance**.
- Génération d'un identifiant unique : `NOM_PRENOM_JJ_MM_AAAA`.
- Recherche dynamique des contacts.
- Gestion de mots-clés sur chaque contact.
- Persistance en base de données locale SQLite (pas de cookies navigateur).
- Import Excel avec mapping des colonnes vers les catégories du site.

## Lancement

```bash
npm install
npm start
```

Puis ouvrir : `http://localhost:3000`

## Notes techniques

- Base locale : `data/umanager.db`
- API backend : Express + SQLite
- Import Excel : package `xlsx`
