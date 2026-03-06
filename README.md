# Mini app web — Tableau de bord Expéditions

**But :** déposer un fichier Excel hebdomadaire (.xlsx) et obtenir immédiatement des **KPIs** + **graphiques** interactifs (par semaine, top matériel, top IPO/SO), avec boutons **Réinitialiser** et **Exporter CSV**.

## 🚀 Déploiement sur GitHub Pages
1. Crée un repo public (ex. `expeditions-dashboard`).
2. Ajoute ces 3 fichiers à la racine : `index.html`, `styles.css`, `app.js`.
3. Commit & push.
4. Dans **Settings → Pages** :
   - Source = **Deploy from a branch**
   - Branch = **main / root**
5. L’URL sera du type `https://<ton-user>.github.io/expeditions-dashboard/`.

## 📦 Utilisation
- Clique **Déposer/Choisir un fichier** et sélectionne ton Excel hebdo.
- Les colonnes sont détectées automatiquement (tolérance accents/variantes : *SEMAINE, Date réception, matériel, IPO/SO, Poids Brut réel (kg), Volume (m3)*).
- Utilise les **filtres multi‑sélection** sur *Semaine*, *Matériel*, *IPO/SO*.
- Bouton **Réinitialiser** : nettoie tous les filtres et reconstruit les graphes.
- Bouton **Exporter** : télécharge les données **filtrées** en CSV.

## 🧩 Notes techniques
- **Charting** : [Chart.js 4](https://www.chartjs.org/)
- **Parsing Excel** : [SheetJS / XLSX](https://sheetjs.com/)
- Tout tourne **100% côté client** (pas de serveur). Idéal pour GitHub Pages.

## 🔧 Personnalisation
- Couleurs : édite `:root` dans `styles.css`.
- Champs supplémentaires : adapte la détection dans `app.js` (fonction `findCol`).
