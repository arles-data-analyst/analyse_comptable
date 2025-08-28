# Analyse comptable – Python

> Nettoyage, analyse et visualisations comptables à partir d’un journal de factures (achats, ventes, TVA, banque).  
> Objectifs : montrer un workflow data pro (Python + pandas + Matplotlib) et des livrables exploitables (Excel, PDF, images).

##  Résumé
- **Nettoyage** : normalisation des colonnes, gestion des dates et montants.
- **Analyses** : solde cumulé, tendance mensuelle, top comptes, top dépenses par libellé, flux entrées/sorties.
- **Livrables** : `rapport_comptable.xlsx`, `rapport_comptable.pdf`, graphiques `.png`.

##  Données
- Fichier source : `factures_comptables_brutes.xlsx`
- Colonnes attendues : `Date` (date), `Compte` (texte/numérique), `Libellé` (texte), `Montant` (numérique)
- Exemple de comptes : 411 (Clients), 512 (Banque), 601 (Achats), 606 (Fournitures), 44566 (TVA déductible)

## 🔧 Installation rapide
```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
