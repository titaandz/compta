# Macro Excel - Analyse Financière Comparative

## Description

Macro VBA Excel complète pour analyser deux exercices comptables comparés à partir de deux balances générales. Génère automatiquement tous les onglets d'analyse selon les normes **PCG / ANC (Règlement 2014-03)**.

---

## Onglets générés (20 onglets)

### Bilan (onglets bleus)
| Onglet | Contenu |
|--------|---------|
| `Actif immobilise` | Immo. incorporelles, corporelles, financières + amortissements |
| `Stocks` | Classe 3 + variations de stocks + ratios |
| `Creances` | Clients, TVA, fiscal + DSO (jours) |
| `Tresorerie` | Classe 5, disponibilités, VMP, CBC |
| `Capitaux propres` | Capital, réserves, résultat, subventions |
| `Dettes` | LMT, fournisseurs, fiscal/social + DPO (jours) |
| `Provisions` | Tableau de mouvement des provisions 15x |

### Compte de résultat (onglets verts)
| Onglet | Contenu |
|--------|---------|
| `CA` | Chiffre d'affaires 70x détaillé + analyse par nature |
| `Produits exploitation` | 71x à 75x, 781, 791 + production de l'exercice |
| `Charges exploitation` | 60x à 65x + ratios / CA |
| `Personnel` | 64x détaillé + taux de charges, % VA, % CA |
| `Dotations` | 68x + impact CAF |
| `Financier` | 66x et 76x + ratios de couverture |
| `Exceptionnel` | 67x et 77x + plus/moins-values cessions |
| `Impots` | 69x + taux effectif d'IS |

### Récapitulatifs et ratios
| Onglet | Contenu |
|--------|---------|
| `Recap Bilan` | Bilan complet Actif/Passif selon format ANC |
| `Recap PL` | Compte de résultat complet selon format PCG |
| `SIG` | 8 soldes intermédiaires + 10 ratios de rentabilité |
| `BFR-CAF` | FRNG, BFR (exploitation / hors expl.), Trésorerie nette, CAF (méthode additive) + 8 ratios |
| `Annexe ANC` | États financiers normés : règles comptables, tableau immo., amortissements, provisions R&C, état créances/dettes, informations personnel |
| `Synthese PPT` | Dashboard KPIs, cascade de résultat, bilan simplifié, points d'attention automatiques |

---

## Ratios calculés

### SIG (Soldes Intermédiaires de Gestion)
- Marge commerciale
- Production de l'exercice
- Valeur Ajoutée (VA)
- Excédent Brut d'Exploitation (EBE / EBITDA)
- Résultat d'Exploitation (REX / EBIT)
- Résultat Courant Avant IS (RCA)
- Résultat Exceptionnel
- Résultat Net

### BFR (Besoin en Fonds de Roulement)
- FRNG (Fonds de Roulement Net Global)
- BFR total, BFR exploitation, BFR hors exploitation
- Trésorerie Nette
- BFR en jours de CA
- FRNG en jours de CA

### CAF (Capacité d'Autofinancement)
- Méthode additive à partir du résultat net
- Réintégration des charges calculées (dotations)
- Exclusion des produits calculés (reprises)
- Exclusion des éléments exceptionnels non récurrents

### Autres ratios
- DSO (jours de crédit client)
- DPO (jours de crédit fournisseur)
- Taux de charges sociales
- Taux effectif d'IS
- Levier financier (Dette/CP)
- Couverture des intérêts (EBE/Charges fin.)

---

## Installation

### Étape 1 : Ouvrir Excel
Créer un nouveau classeur Excel (`.xlsm` - avec macros).

### Étape 2 : Ouvrir l'éditeur VBA
- Raccourci : `Alt + F11`
- Ou : Onglet Développeur > Visual Basic

### Étape 3 : Importer les modules
1. Dans VBE : **Fichier > Importer un fichier** (ou `Ctrl+M`)
2. Importer dans cet ordre :
   - `Module1_Configuration.bas`
   - `Module2_Bilan.bas`
   - `Module3_PL.bas`
   - `Module4_Recaps_SIG.bas`
   - `Module5_BFR_CAF_ANC.bas`
   - `Module6_PPT.bas`
   - `INSTALLER_MACRO.bas`

### Étape 4 : Activer les macros
- Fichier > Options > Centre de gestion de la confidentialité > Paramètres des macros
- Activer toutes les macros (ou Activer uniquement les macros signées)

### Étape 5 : Préparer les balances
Option A - **Utiliser les données exemple** :
```
Alt + F8 > DemarrerAnalyse > OUI
```
Cela crée des onglets `Balance_N` et `Balance_N-1` avec des données fictives.

Option B - **Utiliser votre propre balance** :
1. Créer un onglet nommé exactement `Balance_N`
2. Créer un onglet nommé exactement `Balance_N-1`
3. Format attendu (ligne 1 = en-têtes) :

| A | B | C | D |
|---|---|---|---|
| Numero compte | Intitule compte | Debit | Credit |
| 411000 | Clients | 850000 | 0 |
| 401000 | Fournisseurs | 0 | 320000 |

### Étape 6 : Lancer l'analyse
```
Alt + F8 > LancerAnalyseFinanciere > Exécuter
```
Ou via `DemarrerAnalyse` pour l'assistant.

---

## Format des balances

### Compatibilité
La macro est compatible avec les exports de :
- **Sage** (Comptabilité 50/100, i7)
- **Cegid** (Y2/Expert)
- **EBP** (Comptabilité Classic/Pro)
- **Quadratus**
- **Loop** (anciennement Fulll)
- **ACD** / **Coala**

### Règles de saisie
1. **Numéros de compte** : Format texte ou numérique (sans espaces). Les 2 premiers chiffres suffisent pour la plupart des calculs (`60` = tous les comptes 60xxxx).
2. **Débit / Crédit** : Valeurs positives uniquement. Pas de soldes nettés.
3. **Solde** : Calculé automatiquement comme `Débit - Crédit`.
   - Solde débiteur positif → actif, charges
   - Solde créditeur négatif → passif, produits
4. **Balance avant affectation** : Le compte 12 (résultat) doit figurer dans la balance.

---

## Normes ANC

La macro implémente les règles du **Règlement ANC 2014-03** (PCG) :
- Tableau des immobilisations et amortissements (Art. R.123-197)
- État des créances et dettes avec échéancier (Art. R.123-198)
- Tableau des provisions pour risques et charges
- Informations sur le personnel (effectifs, masse salariale)
- Règles et méthodes comptables

---

## Personnalisation

### Changer les noms d'onglets source
Dans `Module1_Configuration.bas` :
```vba
Public Const NOM_ONGLET_N As String = "Balance_N"      ' Modifier ici
Public Const NOM_ONGLET_N1 As String = "Balance_N-1"   ' Modifier ici
```

### Changer les libellés d'exercice
```vba
Public Const NOM_EXERCICE_N As String = "2024"    ' Au lieu de "N"
Public Const NOM_EXERCICE_N1 As String = "2023"   ' Au lieu de "N-1"
```

---

## Structure des fichiers

```
compta/
├── Module1_Configuration.bas   # Configuration, chargement balances, utilitaires
├── Module2_Bilan.bas           # Onglets Actif/Passif (classes 1, 2, 3, 4, 5)
├── Module3_PL.bas              # Onglets P&L (classes 6 et 7)
├── Module4_Recaps_SIG.bas      # Recap Bilan, Recap P&L, SIG
├── Module5_BFR_CAF_ANC.bas     # BFR, CAF, Annexe ANC normée
├── Module6_PPT.bas             # Synthèse PPT avec KPIs et alertes automatiques
└── INSTALLER_MACRO.bas         # Assistant démarrage + données exemple
```
