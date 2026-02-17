# Changelog

Tous les changements notables de ce projet seront documentés dans ce fichier.

Le format est basé sur [Keep a Changelog](https://keepachangelog.com/fr/1.0.0/),
et ce projet adhère à [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.0.0] - 2026-02-17

### Ajouté

#### Onglet 1 : Rapport Financier
- ✅ Import fichiers Excel (prévisions + coûts réels)
- ✅ Calcul courbes en S avec paramètre k ajustable (4-20)
- ✅ Mode d'affichage : Vue Globale / Vue S uniquement
- ✅ 4 KPI : Budget Révisé Total, Coûts Réels Total, Écart, % Avancement
- ✅ Graphique courbes en S (Budget, Budget Révisé, AT, Réels cumulés)
- ✅ Ligne "Aujourd'hui" sur le graphique pour repère temporel
- ✅ 2 camemberts : Répartition Budget Révisé et AT par projet
- ✅ Graphique barres horizontales : Avancement par ID Projet
- ✅ Tableau détails mensuels avec colonne Écart
- ✅ Tableau comparatif par ID Projet avec totaux
- ✅ Export PNG (graphiques haute résolution)
- ✅ Export PNG (tableaux via html2canvas)
- ✅ Export Excel (tableaux avec SheetJS)

#### Onglet 2 : Imputations
- ✅ 4 KPI Honoraires : Budget Révisé, AT, Réel, Reste à imputer
- ✅ Extraction automatique métier depuis "Nom de la tâche" (regex)
- ✅ Graphique horizontal Honoraires par métier (Budget Révisé / AT / Réel)
- ✅ Tableau ressources : UNIQUEMENT Réel, tri décroissant par montant
- ✅ Graphique horizontal Synthèse par métier (Réel uniquement)

#### Onglet 3 : Planning
- ✅ Calcul automatique dates réelles depuis fichier couts_reel.xlsx
- ✅ Diagramme Gantt horizontal (3 barres par projet : BR / AT / Réel)
- ✅ Tableau consolidé planning avec colonnes complètes
- ✅ Statut automatique avec badges colorés (Prévu / En cours / Terminé / Retard)
- ✅ Export PNG du Gantt

#### Général
- ✅ Design moderne avec couleurs Groupe E Celsius (#163a5f)
- ✅ Interface responsive (desktop, tablette, mobile)
- ✅ Navigation par onglets fluide
- ✅ Messages de statut avec codes couleur
- ✅ Animations et effets hover
- ✅ Tooltips interactifs sur tous les graphiques
- ✅ Gestion des erreurs de format Excel
- ✅ Ordre des projets : "Participations clients" toujours en dernier

#### Documentation
- ✅ README.md complet avec instructions détaillées
- ✅ Structure fichiers Excel documentée
- ✅ Mapping ID Projets → Noms lisibles
- ✅ Recommandations coefficient k
- ✅ .gitignore pour fichiers sensibles
- ✅ Roadmap versions futures

### Technique
- ✅ HTML5 sémantique
- ✅ CSS3 avec variables personnalisées
- ✅ JavaScript Vanilla (pas de framework)
- ✅ Chart.js v4.4.0 pour graphiques
- ✅ SheetJS v0.18.5 pour lecture Excel
- ✅ html2canvas v1.4.1 pour exports PNG tableaux
- ✅ Plugin personnalisé "todayLine" pour Chart.js

---

## [En préparation]

### Version 1.1
- [ ] Filtre par ID Projet dans tous les onglets
- [ ] Export PDF complet (3 onglets)
- [ ] Mode sombre avec toggle
- [ ] Sauvegarde préférences utilisateur (localStorage)
- [ ] Amélioration extraction métier (plus de patterns)

### Version 1.2
- [ ] Import automatique depuis serveur distant
- [ ] Alertes configurables (seuils d'écart personnalisés)
- [ ] Prévisions dynamiques (projection coûts finaux)
- [ ] Comparaison multi-projets simultanée
- [ ] Graphique évolution temporelle des écarts

### Version 2.0
- [ ] Backend Node.js + PostgreSQL
- [ ] Authentification utilisateurs (JWT)
- [ ] API REST pour intégration externe
- [ ] Notifications email automatiques (seuils dépassés)
- [ ] Historique des versions de fichiers Excel
- [ ] Audit trail des modifications

---

## Notes de version

### [1.0.0] - Première version stable

Première version complète et fonctionnelle du Dashboard CAD pour Groupe E Celsius.

**Points forts :**
- Interface moderne et responsive
- Analyse financière complète avec courbes S
- Suivi honoraires par métier avec extraction automatique
- Planning Gantt avec calcul automatique dates réelles
- Exports multiples (PNG, Excel)
- Documentation exhaustive

**Limitations connues :**
- Pas de sauvegarde des préférences utilisateur
- Export PDF non disponible (contournement : captures PNG)
- Filtre par projet global non implémenté
- Calcul métier basé uniquement sur "Honoraires" dans le nom

**Navigateurs testés :**
- ✅ Chrome 120+
- ✅ Firefox 121+
- ✅ Edge 120+
- ✅ Safari 17+

---

**Légende :**
- ✅ Fonctionnalité implémentée et testée
- [ ] Fonctionnalité planifiée
- 🐞 Bug corrigé
- ⚠️ Breaking change