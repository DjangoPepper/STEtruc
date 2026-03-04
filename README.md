# ⚡ STEtruc

Application mobile de nettoyage de fichiers Excel, combinant :
- **Visuel** de [coil-deploy](https://github.com/DjangoPepper/coil-deploy) — thème navy sombre, mobile-first, navigation bottom
- **Fonctions** de [STEpi](https://github.com/DjangoPepper/STEpi) — import Excel multi-sheets, nettoyage, export

## Fonctionnalités

- **Import** : Excel (.xlsx/.xls) et CSV, drag & drop, détection automatique des en-têtes
- **Multi-onglets** : navigation, masquage, suppression ou conservation sélective
- **Tableau** : renommage d'en-têtes au clic, ajout de lignes, grisage des valeurs répétitives (≥35%)
- **Nettoyage** : masquage de colonnes / lignes avec restauration
- **Export** : fichier `.xlsx` propre avec nom personnalisé

## Stack

- React 18 + TypeScript + Vite
- SheetJS / xlsx (parsing & export Excel)
- CSS-in-JS (zéro dépendance UI)

## Démarrage local

```bash
npm install
npm run dev
```

## Déploiement GitHub Pages

### Méthode 1 — script npm (manuel)

```bash
npm run build
npm run deploy
```

> Cela pousse le dossier `dist/` sur la branche `gh-pages`.

### Méthode 2 — GitHub Actions (automatique)

Le workflow `.github/workflows/deploy.yml` se déclenche à chaque push sur `main`.

**Étapes pour activer :**
1. Aller dans **Settings → Pages** du dépôt
2. Source : **GitHub Actions**
3. Pousser sur `main` → déploiement automatique

L'application sera disponible sur :  
`https://<username>.github.io/STEtruc/`

> `vite.config.ts` est déjà configuré avec `base: '/STEtruc/'`

## Structure

```
STEtruc/
├── .github/workflows/deploy.yml   ← CI/CD GitHub Actions
├── public/
├── src/
│   ├── main.tsx
│   ├── index.css
│   └── App.tsx                    ← Application complète (1 fichier)
├── index.html
├── package.json
├── vite.config.ts
└── tsconfig*.json
```
