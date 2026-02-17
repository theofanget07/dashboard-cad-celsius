# 📊 Dashboard CAD - Groupe E Celsius

![Version](https://img.shields.io/badge/version-1.0.0-blue.svg)
![License](https://img.shields.io/badge/license-MIT-green.svg)

Dashboard web interactif pour le suivi de projets de **Chauffage à Distance (CAD)** chez Groupe E Celsius. Ce outil permet une analyse financière approfondie, le suivi des imputations par métier et la gestion du planning de réalisation.

## 🎯 Objectifs

- **Analyser** les écarts entre budgets prévisionnels et coûts réels
- **Suivre** l'avancement financier avec courbes en S et KPI
- **Identifier** la répartition des honoraires par métier
- **Visualiser** le planning réel vs prévisionnel (Gantt)
- **Exporter** graphiques et tableaux (PNG, Excel)

## ✨ Fonctionnalités

### 📈 Onglet 1 : Rapport Financier

#### KPI en temps réel
- Budget Révisé Total
- Coûts Réels Total
- Écart (avec code couleur)
- % Avancement projet

#### Graphiques
- **Courbes en S** : Budget / Budget Révisé / AT / Réels cumulés
  - Paramètre `k` ajustable (courbure)
  - Vue Globale ou Vue S uniquement
  - Ligne "Aujourd'hui" pour repère temporel
- **2 Camemberts** : Répartition Budget Révisé et AT par projet
- **Graphique barres horizontales** : Comparaison Budget Révisé vs Réel par projet

#### Tableaux
- **Détails mensuels** : Montants S, cumuls, réels, écarts
- **Comparatif par ID Projet** : Budget / Budget Révisé / AT / Réel / Écarts / %

#### Exports
- PNG haute résolution (graphiques + tableaux)
- Excel (tableaux avec colonnes écart)

---

### 💼 Onglet 2 : Imputations

#### KPI Honoraires
- Honoraires Budget Révisé
- Honoraires AT
- Honoraires Réels
- Reste à imputer (avec code couleur)

#### Extraction intelligente des métiers
Analyse automatique du champ `Nom de la tâche` pour extraire le métier :
```
"Honoraires conducteur travaux interne_Réseau CAD_SIA 5"
→ Métier : "conducteur travaux interne"
```

#### Graphiques
- **Honoraires par métier** (horizontal) : Budget Révisé / AT / Réel
- **Synthèse par métier** (horizontal) : Réel uniquement

#### Tableau
- **Détails par ressource** : Ressource / Métier / Projet / Montant Réel
- Tri décroissant par montant
- **Uniquement les réels** (pas de prévisions)

---

### 📅 Onglet 3 : Planning

#### Dates réelles automatiques
Calcul automatique des dates de début/fin réelles depuis les imputations (`couts_reel.xlsx`)

#### Diagramme de Gantt horizontal
- 3 barres par projet : Budget Révisé / AT / Réel
- Couleurs distinctes pour chaque type
- Tooltips interactifs avec dates

#### Tableau consolidé
Colonnes :
- Projet
- Début BR / Fin BR
- Début AT / Fin AT
- Début Réel / Fin Réel
- Écart (jours)
- **Statut automatique** avec badge coloré :
  - 🔵 Prévu
  - 🟡 En cours
  - 🟢 Terminé
  - 🔴 Retard

#### Export
- PNG haute résolution du Gantt

---

## 📁 Structure des Fichiers Excel

### Fichier 1 : `previsions_realisation.xlsx`

Fichier de **prévisions** contenant les données budgétaires et d'atterrissage.

**Colonnes utilisées :**

| Colonne | Description | Exemple |
|---------|-------------|----------|
| `ID projet` | Identifiant projet (minuscule) | `CL150096_11_05_04` |
| `Nom de la tâche` | Description tâche avec métier | `Honoraires conducteur travaux interne_Réseau CAD_SIA 5` |
| `Modèle de prévision` | Type de budget | `Budget`, `Budget_Rev`, `AT` |
| `Catégorie` | Catégorie de coût | `Gestion de projet` |
| `Ressource` | Nom de la ressource | `Dupont Jean` |
| `Montant total du coût` | Montant en CHF | `125000` |
| `Date de début` | Date de début tâche | `01.01.2024` |
| `Date de fin` | Date de fin tâche | `31.12.2024` |

---

### Fichier 2 : `couts_reel.xlsx`

Fichier des **coûts réels** contenant les imputations effectives.

**Colonnes utilisées :**

| Colonne | Description | Exemple |
|---------|-------------|----------|
| `ID Projet` | Identifiant projet (**majuscule**) | `CL150096_11_05_04` |
| `Nom de la tâche` | Description tâche (format identique) | `Honoraires conducteur travaux interne_Réseau CAD_SIA 5` |
| `ID catégorie` | Catégorie de coût | `Ingénierie` |
| `Ressource.Nom de la ressource` | Nom ressource (**avec point**) | `Dupont Jean` |
| `Montant total du coût` | Montant en CHF | `8500` |
| `Date du projet` | Date de l'imputation | `15.03.2024` |

---

### Mapping des ID Projets

| ID Projet | Nom Complet |
|-----------|-------------|
| `CL150096_11_05_01` | Général - Direction |
| `CL150096_11_05_02` | Centrale |
| `CL150096_11_05_03` | Bâtiment |
| `CL150096_11_05_04` | **Réseau CAD** |
| `CL150096_11_05_05` | Sous-stations |
| `CL150096_11_05_06` | Participations clients |

⚠️ **Important** : "Participations clients" apparaît **toujours en dernier** dans les tableaux et graphiques.

---

## 🚀 Installation & Utilisation

### Prérequis
- Navigateur web moderne (Chrome, Firefox, Edge, Safari)
- Fichiers Excel `previsions_realisation.xlsx` et `couts_reel.xlsx`

### Installation

1. **Cloner le repository**
```bash
git clone https://github.com/theofanget07/dashboard-cad-celsius.git
cd dashboard-cad-celsius
```

2. **Ouvrir le Dashboard**
```bash
# Ouvrir directement dans le navigateur
open index.html

# Ou avec un serveur local (recommandé)
python -m http.server 8000
# Puis ouvrir http://localhost:8000
```

### Utilisation

1. **Importer les fichiers Excel**
   - Cliquer sur "📁 Prévisions (Excel)" et sélectionner `previsions_realisation.xlsx`
   - Cliquer sur "📁 Coûts Réels (Excel)" et sélectionner `couts_reel.xlsx`

2. **Ajuster le coefficient k**
   - Utiliser le curseur ou le champ numérique
   - **Recommandations** :
     - `6-8` : Projets linéaires, études
     - `10-14` : Projets CAPEX réseau, génie civil
     - `14-18` : Projets courts, réalisation rapide

3. **Choisir le mode d'affichage**
   - **Vue Globale** : Toutes les données (réels + prévisions)
   - **Vue S uniquement** : Période projet Budget Révisé uniquement

4. **Explorer les onglets**
   - 📈 Rapport Financier : Vue d'ensemble financière
   - 💼 Imputations : Détail honoraires par métier
   - 📅 Planning : Gantt et écarts temporels

5. **Exporter les données**
   - Boutons d'export disponibles dans chaque onglet
   - Formats : PNG (graphiques/tableaux), Excel (tableaux)

---

## 📸 Screenshots

### Onglet 1 : Rapport Financier
![Rapport Financier](./screenshots/rapport-financier.png)
*KPI, courbes en S, camemberts et tableaux comparatifs*

### Onglet 2 : Imputations
![Imputations](./screenshots/imputations.png)
*Honoraires par métier et détails par ressource*

### Onglet 3 : Planning
![Planning](./screenshots/planning.png)
*Gantt interactif et tableau consolidÃ© avec statuts*

---

## 🛠️ Technologies

- **HTML5** : Structure sémantique
- **CSS3** : Design moderne, responsive, animations
- **JavaScript (Vanilla)** : Logique métier sans framework
- **Chart.js v4.4.0** : Graphiques interactifs (courbes, barres, camemberts)
- **SheetJS (XLSX) v0.18.5** : Lecture fichiers Excel
- **html2canvas v1.4.1** : Export tableaux en PNG

---

## 📝 Règles Métier

1. **"Participations clients" toujours en dernier** avant le total (tableaux, graphiques, planning)
2. **Graphiques en barres = horizontaux** (`indexAxis: 'y'`)
3. **Métier extrait de "Nom de la tâche"** pour les honoraires via regex
4. **Planning réel automatisé** (calcul depuis `couts_reel.xlsx`, pas de saisie manuelle)
5. **Design Groupe E Celsius** : Couleur brand `#163a5f`, dégradés, ombres

---

## 📈 Roadmap

### Version 1.1 (Prévue)
- [ ] Filtre par ID Projet dans tous les onglets
- [ ] Export PDF complet (3 onglets)
- [ ] Mode sombre
- [ ] Sauvegarde préférences utilisateur (localStorage)

### Version 1.2
- [ ] Import automatique depuis serveur distant
- [ ] Alertes configurables (seuils d'écart)
- [ ] Prévisions dynamiques (projection coûts finaux)
- [ ] Comparaison multi-projets

### Version 2.0
- [ ] Backend Node.js + base de données
- [ ] Authentification utilisateurs
- [ ] API REST pour intégration externe
- [ ] Notifications email automatiques

---

## 👥 Contexte

**Groupe E Celsius** - Fribourg, Suisse  
Projets CAPEX de chauffage à distance selon phases SIA

**Responsable Projets** : Gestion multi-projets avec suivi budgétaire, honoraires et planning de réalisation.

---

## 🤝 Contribution

Les contributions sont les bienvenues !

1. Fork le projet
2. Créer une branche feature (`git checkout -b feature/amelioration`)
3. Commit les changements (`git commit -m 'Add amelioration'`)
4. Push vers la branche (`git push origin feature/amelioration`)
5. Ouvrir une Pull Request

---

## 📝 Licence

Ce projet est sous licence MIT. Voir le fichier [LICENSE](LICENSE) pour plus de détails.

---

## 📧 Contact

**Théo Fanget**  
Email: theofanget07@gmail.com  
GitHub: [@theofanget07](https://github.com/theofanget07)

---

<div align="center">
  <strong>Développé avec ❤️ pour Groupe E Celsius</strong>
</div>