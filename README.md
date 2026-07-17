# 🍽️ Planificateur de pauses déjeuner — WFM

Application Streamlit pour attribuer et optimiser les pauses déjeuner des
agents d'un plateau de centre de contacts.

## Fonctionnalités

- Configuration des plages horaires principales et intermédiaires (30 min)
- Capacité max par tranche + tolérance de dépassement configurable
- Import du planning agents (Excel/CSV) avec validation d'éligibilité
  (règle 8h-10h, dérogation allaitement à 7h)
- Attribution automatique des pauses (règle d'antériorité)
- Ajustement manuel (Human in the Loop) via grille interactive
- Export du planning final et d'un modèle Excel vide

## Installation

```bash
pip install -r requirements.txt
```

## Lancement

```bash
streamlit run lunch_break_scheduler.py
```

L'application s'ouvre dans le navigateur sur `http://localhost:8501`.

## Format du fichier de planning attendu

| Nom | Heure_debut | Heure_fin | Allaitement |
|-----|--------------|-----------|-------------|
| Dupont Jean | 07:00 | 16:00 | Non |
| Martin Alice | 09:00 | 17:00 | Oui |

Un modèle vide est téléchargeable directement depuis la barre latérale
de l'application.

## Structure du dépôt

```
.
├── lunch_break_scheduler.py   # Application Streamlit (logique + UI)
├── requirements.txt           # Dépendances Python
├── .gitignore
└── README.md
```
