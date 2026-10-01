# Suivi des heures de vol et des cycles de la flotte

Application **Streamlit** pour analyser l'historique d'utilisation d'une flotte d'aéronefs :
nettoyage des données, KPIs, tableau de bord, filtres, authentification et animation
**Bar Chart Race** de la progression des heures de vol et des cycles (atterrissages).

Technologies : Python, Streamlit, pandas, matplotlib, seaborn.

## Fonctionnalités

| Module | Contenu |
|---|---|
| **Connexion** | Page de login (identifiants hachés en SHA-256), déconnexion |
| **Nettoyage** | Détection automatique des colonnes, normalisation du texte, conversion des types, suppression des doublons et des lignes vides, valeurs négatives/manquantes, cohérence des cumuls ; un rapport détaille chaque étape |
| **KPIs** | Aéronefs suivis, heures et cycles cumulés, durée moyenne par cycle, taux de disponibilité, immobilisés, avion le plus utilisé, seuils dépassés |
| **Tableau de bord** | Heures et cycles par avion, répartition par statut, heures vs cycles, tableau des aéronefs à surveiller |
| **Analyses** | Heures par cycle, moyennes par modèle et compagnie, distribution, corrélations, évolution dans le temps |
| **Bar Chart Race** | Animation GIF configurable (indicateur, nombre de barres, vitesse, fluidité) et téléchargeable |
| **Filtres** | Compagnie, modèle, statut, plage d'heures de vol, période (si historique) |
| **Export** | Données nettoyées en CSV et Excel |

## Installation

```bash
python -m venv .venv
source .venv/bin/activate        # Windows : .venv\Scripts\activate
pip install -r requirements.txt
```

## Lancement

Placez `suivi_heures_vol_cycles.xlsx` à côté de `app.py` (il sera chargé par défaut), puis :

```bash
streamlit run app.py
```

Vous pouvez aussi importer votre propre fichier `.xlsx` ou `.csv` depuis le menu de gauche.

## Connexion

Comptes de démonstration (utilisés si aucun secret n'est défini) :

| Identifiant | Mot de passe |
|---|---|
| `admin` | `flotte2026` |
| `technicien` | `avion2026` |

**À changer avant toute mise en ligne.** Créez le fichier `.streamlit/secrets.toml` :

```toml
[users]
admin = "empreinte_sha256_du_mot_de_passe"
marie = "empreinte_sha256_du_mot_de_passe"
```

Pour générer une empreinte :

```bash
python -c "import hashlib; print(hashlib.sha256('MonMotDePasse'.encode()).hexdigest())"
```

N'ajoutez jamais `secrets.toml` à Git (`.gitignore`). Ce système convient pour un usage interne ;
pour de la production, préférez une authentification SSO/OIDC.

## Format des données attendu

Les colonnes sont reconnues automatiquement par mots-clés (accents et casse ignorés).

| Donnée | Obligatoire | Exemples d'intitulés reconnus |
|---|---|---|
| Immatriculation | Oui | `Immatriculation`, `Registration` |
| Heures de vol | Oui | `Heures de Vol (h)`, `Hours` |
| Cycles | Oui | `Cycles (Atterrissages)`, `Landings` |
| Modèle | Non | `Modèle`, `Model`, `Type` |
| Statut | Non | `Statut`, `Status`, `État` |
| Date | Non | `Date`, `Dernier Contrôle` |
| Compagnie | Non | `Compagnie`, `Airline` |

## Important : historique réel ou simulé

Un Bar Chart Race a besoin d'une **série temporelle** (plusieurs dates par avion).

- **Plusieurs lignes par avion avec des dates différentes** : l'application utilise votre historique réel.
- **Une seule ligne par avion** (cas du fichier d'exemple, 10 avions, une date de dernier contrôle chacun) :
  l'application **reconstitue un historique mensuel simulé** (nombre de mois et ancienneté réglables
  dans « Simulation de l'historique »). Les cumuls sont croissants et se terminent exactement sur vos valeurs
  réelles, mais l'évolution passée est fictive. Un avertissement est affiché dans l'interface.

Pour une animation fidèle, exportez de votre système de GMAO un relevé périodique (ex. fin de mois) :

| Immatriculation | Modèle | Heures de Vol (h) | Cycles | Statut | Date |
|---|---|---|---|---|---|
| CN-AOA | Boeing 737-800 | 12100 | 8050 | Opérationnel | 2026-07-31 |
| CN-AOA | Boeing 737-800 | 12450 | 8300 | Opérationnel | 2026-08-31 |

## Choix techniques

- La bibliothèque `bar_chart_race` de l'ancien `app.py` n'est plus maintenue et pose des problèmes avec
  les versions récentes de pandas/matplotlib ; elle exige aussi `ffmpeg`. L'animation est donc dessinée
  directement avec **matplotlib** et exportée en **GIF** via Pillow : aucune dépendance système.
- Les calculs lourds (chargement, nettoyage, animation) sont mis en cache avec `st.cache_data`.

## Structure du projet

```
.
├── app.py
├── requirements.txt
├── README.md
├── suivi_heures_vol_cycles.xlsx      # données d'exemple (optionnel)
└── .streamlit/secrets.toml           # utilisateurs (optionnel, non versionné)
```

## Dépannage

- **« Colonnes obligatoires introuvables »** : vérifiez que le fichier contient immatriculation, heures de vol et cycles.
- **Animation lente** : réduisez la fluidité, le nombre de mois ou le nombre de barres.
- **Erreur sur `width="stretch"`** : mettez à jour Streamlit (`pip install -U streamlit`, version 1.50 ou plus).
