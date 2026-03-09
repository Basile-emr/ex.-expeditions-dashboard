# Mini app web — Dashboard Caisses (Expéditions)

**But :** déposer un fichier Excel hebdomadaire (.xlsx) et obtenir immédiatement des **KPIs** + **graphiques** interactifs (par semaine, répartition des caisses par dimensions), avec boutons **Réinitialiser** et **Exporter CSV**.

## ✅ Indicateurs
- **Nb caisses**
- **Volume total (m³)**
- **Poids brut total (kg)**
- **Densité moyenne (kg/m³)**
- **Poids moyen / caisse (kg)**
- **Volume moyen / caisse (m³)**
- **Score logistique (0–100)**

## 📊 Graphiques
- Nb **caisses** par semaine
- Volume (m³) par semaine
- Densité (kg/m³) par semaine
- **Poids moyen / caisse** par semaine
- **Répartition des caisses par dimensions (L×l×h en cm)** — regroupement par triplet Longueur×Largeur×Hauteur (arrondi au cm)

> Supprimés : répartition par type d’emballage, dimensions moyennes par type, Top 10 Matériel, Top 10 IPO/SO.

## 🚀 Déploiement GitHub Pages
1. Repo public (ex. `expeditions-dashboard`).
2. Mettre `index.html`, `styles.css`, `app.js`, `README.md` à la **racine**.
3. **Settings → Pages** → *Deploy from a branch* → **main / root**.
4. URL : `https://<user>.github.io/expeditions-dashboard/`.

## 🧰 Tech
- **Chart.js 4**
- **SheetJS / XLSX** pour lire l’Excel **dans le navigateur** (aucun serveur)
