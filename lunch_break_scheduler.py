# -*- coding: utf-8 -*-
"""
=====================================================================
 OUTIL D'ATTRIBUTION ET D'OPTIMISATION DES PAUSES DÉJEUNER — WFM
=====================================================================

Application Streamlit permettant à une "Vigie" / planificateur WFM de :
  1. Paramétrer les plages de pause déjeuner (plages principales + plages
     intermédiaires en 30 min, capacité max par tranche).
  2. Importer le planning des agents (Excel/CSV).
  3. Valider automatiquement l'éligibilité de chaque agent selon son
     amplitude horaire (règle standard 8h-10h, règle dérogatoire
     "Allaitement" à 7h).
  4. Générer un planning de pause optimisé (règle d'antériorité,
     respect de la capacité max, sur-planification maîtrisée en cas
     de sureffectif).
  5. Ajuster manuellement (Human in the Loop) le planning via une
     grille interactive, avec recalcul dynamique des compteurs.
  6. Exporter un template vide et le planning final en Excel.

Lancement :
    streamlit run lunch_break_scheduler.py

Dépendances :
    pip install streamlit pandas openpyxl xlsxwriter
=====================================================================
"""

import io
from datetime import datetime, time, timedelta

import pandas as pd
import streamlit as st

# ---------------------------------------------------------------------------
# CONFIGURATION GÉNÉRALE DE LA PAGE
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="WFM - Planificateur de pauses déjeuner",
    page_icon="🍽️",
    layout="wide",
)

REQUIRED_COLUMNS = ["Nom", "Heure_debut", "Heure_fin", "Allaitement"]
DUREE_PAUSE_MIN = 60  # Durée standard de la pause déjeuner (fixée à 1h)


# ---------------------------------------------------------------------------
# FONCTIONS UTILITAIRES — GESTION DES HEURES
# ---------------------------------------------------------------------------
def hhmm_to_minutes(value) -> int:
    """Convertit une heure 'HH:MM', un datetime.time ou un Timestamp
    en nombre de minutes depuis minuit. Lève ValueError si invalide."""
    if isinstance(value, time):
        return value.hour * 60 + value.minute
    if isinstance(value, datetime):
        return value.hour * 60 + value.minute
    if pd.isna(value):
        raise ValueError("Heure manquante")
    text = str(value).strip()
    # Supporte les cas où Excel stocke une heure comme '08:00:00' ou '8:00'
    parts = text.split(":")
    if len(parts) < 2:
        raise ValueError(f"Format d'heure invalide : '{text}'")
    h, m = int(parts[0]), int(parts[1])
    return h * 60 + m


def minutes_to_hhmm(minutes: int) -> str:
    """Convertit un nombre de minutes depuis minuit en chaîne 'HH:MM'."""
    h, m = divmod(int(minutes), 60)
    return f"{h:02d}:{m:02d}"


def make_slot_label(start_min: int, end_min: int) -> str:
    return f"{minutes_to_hhmm(start_min)}-{minutes_to_hhmm(end_min)}"


# ---------------------------------------------------------------------------
# CONSTRUCTION DES CRÉNEAUX DE PAUSE (PRINCIPAUX + INTERMÉDIAIRES)
# ---------------------------------------------------------------------------
def build_main_slots(start_min: int, end_min: int, duree: int = DUREE_PAUSE_MIN):
    """Construit les créneaux horaires pleins (ex: 11:00-12:00, 12:00-13:00...)."""
    slots = []
    cursor = start_min
    while cursor + duree <= end_min:
        slots.append({"start": cursor, "end": cursor + duree, "type": "Principal"})
        cursor += duree
    return slots


def build_intermediate_candidates(start_min: int, end_min: int, duree: int = DUREE_PAUSE_MIN):
    """Construit la liste de TOUS les créneaux intermédiaires possibles
    (décalés de 30 min), que l'utilisateur pourra choisir librement
    d'activer un par un (ex: activer 12:30-13:30 sans activer 11:30-12:30)."""
    slots = []
    cursor = start_min + 30
    while cursor + duree <= end_min:
        slots.append({"start": cursor, "end": cursor + duree, "type": "Intermédiaire"})
        cursor += 30
    return slots


# ---------------------------------------------------------------------------
# TEMPLATE EXCEL VIDE (MODÈLE À TÉLÉCHARGER)
# ---------------------------------------------------------------------------
def generate_template_excel() -> bytes:
    df_template = pd.DataFrame(
        {
            "Nom": ["Dupont Jean", "Martin Alice"],
            "Heure_debut": ["07:00", "09:00"],
            "Heure_fin": ["16:00", "17:00"],
            "Allaitement": ["Non", "Oui"],
        }
    )
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_template.to_excel(writer, index=False, sheet_name="Planning")
    return buffer.getvalue()


# ---------------------------------------------------------------------------
# VALIDATION DU PLANNING IMPORTÉ
# ---------------------------------------------------------------------------
def validate_agents(df_raw: pd.DataFrame):
    """Calcule l'amplitude, applique les règles d'éligibilité et renvoie
    un DataFrame enrichi avec les colonnes de contrôle."""
    df = df_raw.copy()

    missing_cols = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing_cols:
        raise ValueError(
            "Colonnes manquantes dans le fichier importé : " + ", ".join(missing_cols)
        )

    debut_min, fin_min, amplitude_h, erreurs = [], [], [], []
    for _, row in df.iterrows():
        try:
            d = hhmm_to_minutes(row["Heure_debut"])
            f = hhmm_to_minutes(row["Heure_fin"])
            if f <= d:
                raise ValueError("Heure de fin antérieure ou égale à l'heure de début")
            debut_min.append(d)
            fin_min.append(f)
            amplitude_h.append(round((f - d) / 60, 2))
            erreurs.append("")
        except ValueError as e:
            debut_min.append(None)
            fin_min.append(None)
            amplitude_h.append(None)
            erreurs.append(str(e))

    df["Debut_min"] = debut_min
    df["Fin_min"] = fin_min
    df["Amplitude_h"] = amplitude_h
    df["Erreur_ligne"] = erreurs

    def normalize_allaitement(v):
        return str(v).strip().lower() in ("oui", "yes", "true", "1")

    df["Allaitement_bool"] = df["Allaitement"].apply(normalize_allaitement)

    def statut_eligibilite(row):
        if row["Erreur_ligne"]:
            return "Erreur", row["Erreur_ligne"]
        amp = row["Amplitude_h"]
        if 8 <= amp <= 10:
            return "Éligible", "Amplitude standard (8h-10h)"
        if amp == 7 and row["Allaitement_bool"]:
            return "Éligible", "Dérogation allaitement (7h)"
        if amp == 7 and not row["Allaitement_bool"]:
            return "Exclu", "7h de travail sans dérogation allaitement"
        return "Exclu", f"Amplitude hors règles ({amp}h)"

    statuts = df.apply(statut_eligibilite, axis=1)
    df["Statut_eligibilite"] = statuts.apply(lambda t: t[0])
    df["Motif"] = statuts.apply(lambda t: t[1])

    return df


# ---------------------------------------------------------------------------
# ALGORITHME D'ATTRIBUTION INTELLIGENTE DES PAUSES
# ---------------------------------------------------------------------------
def assign_breaks(df_eligible: pd.DataFrame, slots: list, capacite_max: int, tolerance_depassement: int):
    """
    Attribue à chaque agent éligible un créneau de pause en respectant :
      - la règle d'antériorité (les agents qui commencent le plus tôt
        sont traités en premier et obtiennent les créneaux les plus tôt) ;
      - la capacité max par tranche ;
      - un dépassement toléré de 1 à 2 agents si le volume est trop élevé,
        plutôt que de laisser un agent sans pause.

    Renvoie :
      - assignments : dict {nom_agent: label_creneau_ou_None}
      - slot_counts : dict {label_creneau: nb_agents}
      - non_planifiables : liste des agents pour lesquels aucun créneau
        de la configuration actuelle n'est entièrement couvert par leur
        shift (nécessitent un arbitrage manuel).
    """
    slots_sorted = sorted(slots, key=lambda s: s["start"])
    for s in slots_sorted:
        s["label"] = make_slot_label(s["start"], s["end"])

    slot_counts = {s["label"]: 0 for s in slots_sorted}
    assignments = {}
    non_planifiables = []

    # Règle de l'antériorité : tri des agents par heure de début croissante
    agents_tries = df_eligible.sort_values(["Debut_min", "Nom"])

    for _, agent in agents_tries.iterrows():
        nom = agent["Nom"]

        # Créneaux compatibles = entièrement couverts par le shift de l'agent
        creneaux_possibles = [
            s for s in slots_sorted
            if agent["Debut_min"] <= s["start"] and s["end"] <= agent["Fin_min"]
        ]

        if not creneaux_possibles:
            assignments[nom] = None
            non_planifiables.append(nom)
            continue

        # Passe 1 : respect strict de la capacité max
        choix = next((s for s in creneaux_possibles if slot_counts[s["label"]] < capacite_max), None)

        # Passe 2 : souplesse opérationnelle — dépassement toléré (1 à 2 agents)
        if choix is None:
            choix = next(
                (s for s in creneaux_possibles
                 if slot_counts[s["label"]] < capacite_max + tolerance_depassement),
                None,
            )

        # Passe 3 : au-delà de la tolérance, on force l'agent sur son créneau
        # idéal le plus proche (priorité d'antériorité) plutôt que de le laisser sans pause.
        if choix is None:
            choix = creneaux_possibles[0]

        assignments[nom] = choix["label"]
        slot_counts[choix["label"]] += 1

    return assignments, slot_counts, non_planifiables


# ---------------------------------------------------------------------------
# EXPORT DU PLANNING FINAL
# ---------------------------------------------------------------------------
def generate_export_excel(df_final: pd.DataFrame) -> bytes:
    buffer = io.BytesIO()
    export_cols = ["Nom", "Heure_debut", "Heure_fin", "Amplitude_h", "Allaitement",
                   "Statut_eligibilite", "Creneau_pause"]
    export_cols = [c for c in export_cols if c in df_final.columns]
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_final[export_cols].to_excel(writer, index=False, sheet_name="Planning pauses")
    return buffer.getvalue()


# ===========================================================================
# INTERFACE STREAMLIT
# ===========================================================================
st.title("🍽️ Planificateur de pauses déjeuner — Plateau WFM")
st.caption(
    "Attribution automatique + ajustement manuel des pauses déjeuner, "
    "avec respect des capacités par tranche horaire."
)

# ---------------------------------------------------------------------------
# 1. PANNEAU DE CONFIGURATION (SIDEBAR)
# ---------------------------------------------------------------------------
with st.sidebar:
    st.header("⚙️ Configuration")

    st.subheader("Plage horaire principale")
    heure_debut_plage = st.time_input("Début de la plage", value=time(11, 0), step=1800)
    heure_fin_plage = st.time_input("Fin de la plage (modifiable)", value=time(16, 0), step=1800)

    st.subheader("Capacité")
    capacite_max = st.number_input(
        "Nombre max d'agents simultanés en pause par tranche",
        min_value=1, value=5, step=1,
    )
    tolerance_depassement = st.slider(
        "Tolérance de dépassement en cas de sureffectif (agents en plus)",
        min_value=0, max_value=3, value=2,
        help="Nombre d'agents supplémentaires tolérés sur une tranche clé "
             "plutôt que de laisser un agent sans pause ou hors plage.",
    )

    st.subheader("Plages intermédiaires (30 min)")
    activer_intermediaires = st.checkbox("Activer les plages intermédiaires", value=False)

    debut_min_cfg = hhmm_to_minutes(heure_debut_plage)
    fin_min_cfg = hhmm_to_minutes(heure_fin_plage)

    selected_intermediates = []
    if activer_intermediaires:
        candidats = build_intermediate_candidates(debut_min_cfg, fin_min_cfg)
        options = [make_slot_label(c["start"], c["end"]) for c in candidats]
        choix_labels = st.multiselect(
            "Choisissez librement les plages intermédiaires à activer",
            options=options,
            default=[],
            help="Ex : activer uniquement 12:30-13:30 et 13:30-14:30 "
                 "sans activer 11:30-12:30.",
        )
        label_to_slot = {make_slot_label(c["start"], c["end"]): c for c in candidats}
        selected_intermediates = [label_to_slot[l] for l in choix_labels]

    st.divider()
    st.subheader("📄 Modèle de planning")
    st.download_button(
        "⬇️ Télécharger le modèle Excel vide",
        data=generate_template_excel(),
        file_name="modele_planning_agents.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

# Validation cohérence de la plage
if debut_min_cfg >= fin_min_cfg:
    st.error("⚠️ L'heure de fin de la plage doit être postérieure à l'heure de début.")
    st.stop()

main_slots = build_main_slots(debut_min_cfg, fin_min_cfg)
all_slots = main_slots + selected_intermediates
all_slots = sorted(all_slots, key=lambda s: s["start"])

if not all_slots:
    st.error("⚠️ Aucun créneau de pause ne peut être construit avec cette configuration "
              "(la plage est trop courte pour une pause d'1h).")
    st.stop()

slot_labels_ordered = [make_slot_label(s["start"], s["end"]) for s in all_slots]

with st.expander("🕐 Créneaux de pause actuellement configurés", expanded=False):
    st.write(", ".join(slot_labels_ordered))

# ---------------------------------------------------------------------------
# 2. IMPORT DU PLANNING AGENTS
# ---------------------------------------------------------------------------
st.header("1️⃣ Import du planning des agents")
fichier = st.file_uploader(
    "Uploader le planning (Excel .xlsx ou CSV) — colonnes attendues : "
    "Nom, Heure_debut, Heure_fin, Allaitement",
    type=["xlsx", "csv"],
)

if fichier is None:
    st.info("📥 Importez un fichier pour lancer l'analyse, ou téléchargez le modèle "
             "dans le panneau latéral.")
    st.stop()

try:
    if fichier.name.lower().endswith(".csv"):
        df_raw = pd.read_csv(fichier, dtype=str)
    else:
        df_raw = pd.read_excel(fichier, dtype=str)
except Exception as e:
    st.error(f"❌ Impossible de lire le fichier : {e}")
    st.stop()

try:
    df_validated = validate_agents(df_raw)
except ValueError as e:
    st.error(f"❌ {e}")
    st.stop()

# ---------------------------------------------------------------------------
# 3. RÉSULTATS DE VALIDATION / ÉLIGIBILITÉ
# ---------------------------------------------------------------------------
st.header("2️⃣ Contrôle d'éligibilité")

nb_eligibles = (df_validated["Statut_eligibilite"] == "Éligible").sum()
nb_exclus = (df_validated["Statut_eligibilite"] == "Exclu").sum()
nb_erreurs = (df_validated["Statut_eligibilite"] == "Erreur").sum()

c1, c2, c3 = st.columns(3)
c1.metric("Agents éligibles", nb_eligibles)
c2.metric("Agents exclus", nb_exclus)
c3.metric("Lignes en erreur", nb_erreurs)

def highlight_statut(row):
    color = {"Éligible": "#d4f7d4", "Exclu": "#fff3cd", "Erreur": "#f8d7da"}.get(
        row["Statut_eligibilite"], ""
    )
    return [f"background-color: {color}"] * len(row)

st.dataframe(
    df_validated[
        ["Nom", "Heure_debut", "Heure_fin", "Amplitude_h", "Allaitement",
         "Statut_eligibilite", "Motif"]
    ].style.apply(highlight_statut, axis=1),
    use_container_width=True,
)

df_eligible = df_validated[df_validated["Statut_eligibilite"] == "Éligible"].copy()

if df_eligible.empty:
    st.warning("Aucun agent éligible à la planification des pauses avec ce fichier.")
    st.stop()

# ---------------------------------------------------------------------------
# 4. GÉNÉRATION DU PLANNING DE PAUSE (ALGORITHME)
# ---------------------------------------------------------------------------
st.header("3️⃣ Génération du planning de pause")

if "planning_genere" not in st.session_state:
    st.session_state["planning_genere"] = False

if st.button("🚀 Générer le planning de pause automatiquement", type="primary"):
    assignments, slot_counts, non_planifiables = assign_breaks(
        df_eligible, all_slots, capacite_max, tolerance_depassement
    )
    df_eligible["Creneau_pause"] = df_eligible["Nom"].map(assignments).fillna("Non planifié")
    st.session_state["df_planning"] = df_eligible
    st.session_state["planning_genere"] = True
    st.session_state["non_planifiables"] = non_planifiables

    if non_planifiables:
        st.warning(
            "⚠️ Les agents suivants n'ont aucun créneau entièrement couvert par leur "
            "shift dans la configuration actuelle et ont été forcés sur le créneau le "
            "plus proche (à arbitrer manuellement ci-dessous) : "
            + ", ".join(non_planifiables)
        )
    else:
        st.success("✅ Planning généré sans agent hors-plage.")

# ---------------------------------------------------------------------------
# 5. AJUSTEMENT MANUEL (HUMAN IN THE LOOP) + VISUALISATION DYNAMIQUE
# ---------------------------------------------------------------------------
if st.session_state.get("planning_genere"):
    st.header("4️⃣ Ajustement manuel et visualisation")
    st.caption(
        "Modifiez librement le créneau d'un agent via le menu déroulant. "
        "Les compteurs par tranche se recalculent automatiquement."
    )

    df_edit_source = st.session_state["df_planning"][
        ["Nom", "Heure_debut", "Heure_fin", "Amplitude_h", "Allaitement", "Creneau_pause"]
    ].reset_index(drop=True)

    creneau_options = slot_labels_ordered + ["Non planifié"]

    df_edited = st.data_editor(
        df_edit_source,
        column_config={
            "Creneau_pause": st.column_config.SelectboxColumn(
                "Créneau de pause",
                options=creneau_options,
                required=True,
            ),
            "Nom": st.column_config.TextColumn("Nom", disabled=True),
            "Heure_debut": st.column_config.TextColumn("Début shift", disabled=True),
            "Heure_fin": st.column_config.TextColumn("Fin shift", disabled=True),
            "Amplitude_h": st.column_config.NumberColumn("Amplitude (h)", disabled=True),
            "Allaitement": st.column_config.TextColumn("Allaitement", disabled=True),
        },
        hide_index=True,
        use_container_width=True,
        key="editeur_planning",
    )

    # Recalcul dynamique des compteurs après modification manuelle
    compteur = (
        df_edited["Creneau_pause"]
        .value_counts()
        .reindex(slot_labels_ordered, fill_value=0)
    )

    st.subheader("📊 Occupation par tranche horaire")
    df_occupation = pd.DataFrame(
        {
            "Créneau": compteur.index,
            "Agents affectés": compteur.values,
            "Capacité max": capacite_max,
        }
    )
    df_occupation["Dépassement"] = df_occupation["Agents affectés"] - capacite_max
    df_occupation["Dépassement"] = df_occupation["Dépassement"].clip(lower=0)

    def highlight_occupation(row):
        if row["Agents affectés"] > capacite_max + tolerance_depassement:
            return ["background-color: #f8d7da"] * len(row)
        if row["Agents affectés"] > capacite_max:
            return ["background-color: #fff3cd"] * len(row)
        return ["background-color: #d4f7d4"] * len(row)

    col_a, col_b = st.columns([1, 1])
    with col_a:
        st.dataframe(
            df_occupation.style.apply(highlight_occupation, axis=1),
            use_container_width=True, hide_index=True,
        )
    with col_b:
        st.bar_chart(df_occupation.set_index("Créneau")["Agents affectés"])

    depassements_critiques = df_occupation[
        df_occupation["Agents affectés"] > capacite_max + tolerance_depassement
    ]
    if not depassements_critiques.empty:
        st.error(
            "🔴 Capacité dépassée au-delà de la tolérance autorisée sur : "
            + ", ".join(depassements_critiques["Créneau"].tolist())
        )

    # ---------------------------------------------------------------------
    # 6. EXPORT DU PLANNING FINAL
    # ---------------------------------------------------------------------
    st.header("5️⃣ Export du planning final")

    df_export = st.session_state["df_planning"].drop(columns=["Creneau_pause"]).merge(
        df_edited[["Nom", "Creneau_pause"]], on="Nom", how="left"
    )

    st.download_button(
        "⬇️ Télécharger le planning de pauses (Excel)",
        data=generate_export_excel(df_export),
        file_name="planning_pauses_dejeuner.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
