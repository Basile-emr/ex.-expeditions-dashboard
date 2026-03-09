# Mini app web — Tableau de bord Expéditions (v2)

**But :** déposer un fichier Excel hebdomadaire (.xlsx) et obtenir immédiatement des **KPIs** + **graphiques** interactifs (par semaine, top matériel, top IPO/SO), avec boutons **Réinitialiser** et **Exporter CSV**.

## ✅ Nouveaux indicateurs ajoutés
- **Poids moyen / colis**
- **Volume moyen / colis**
- **Score logistique (0–100)** basé sur volume relatif, cadence colis, proximité de la densité cible (≈500 kg/m³) et homogénéité des poids.
- **Répartition par type d’emballages**
  - Si la colonne *Colisage/Emballage/Type* existe → utilisée directement.
  - Sinon, fallback par **classe dimensionnelle** calculée (S ≤0,3 m³ · M ≤0,9 m³ · L ≤2,0 m³ · XL >2,0 m³).
- **Dimensions moyennes (L/l/h) par type**.

## 🚀 Déploiement GitHub Pages
1. Repo public (ex. `expeditions-dashboard`).
2. Mettre `index.html`, `styles.css`, `app.js`, `README.md` à la racine.
3. **Settings → Pages** → *Deploy from a branch* → **main/root**.
4. URL : `https://<user>.github.io/expeditions-dashboard/`.

## 🧪 Utilisation
- Cliquer **Déposer/Choisir un fichier** et sélectionner l’Excel hebdo.
- Filtres multi‑sélection : *Semaine*, *Matériel*, *IPO/SO*.
- **Réinitialiser** : remet les filtres à zéro et reconstruit tous les graphes.
- **Exporter** : CSV des données **filtrées** (avec type & dimensions si présents).

## 🧰 Tech
- **Chart.js 4** (+ plugin **chartjs-plugin-datalabels**)
- **SheetJS / XLSX** pour lire l’Excel **dans le navigateur** (aucun serveur)
