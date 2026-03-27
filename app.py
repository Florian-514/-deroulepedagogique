import streamlit as st
import pandas as pd
from pptx import Presentation
import io
import json
from openai import OpenAI
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import datetime

st.set_page_config(
    page_title="Déroulé Pédagogique Qualiopi",
    page_icon="📋",
    layout="wide"
)

COLS = [
    "Jour",
    "Horaires",
    "Objectifs pédagogiques",
    "Contenu de la séquence",
    "Moyens pédagogiques",
    "Modalités de validation des acquis",
]


# ─────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────

def get_api_key():
    """Clé depuis st.secrets (prod) ou saisie manuelle (local)."""
    try:
        key = st.secrets.get("OPENAI_API_KEY", "")
        if key:
            return key
    except Exception:
        pass
    return st.session_state.get("api_key", "")


def extract_pptx_text(pptx_bytes: bytes) -> str:
    prs = Presentation(io.BytesIO(pptx_bytes))
    slides = []
    for i, slide in enumerate(prs.slides):
        texts = [
            shape.text.strip()
            for shape in slide.shapes
            if hasattr(shape, "text") and shape.text.strip()
        ]
        if texts:
            slides.append(f"=== Slide {i + 1} ===\n" + "\n".join(texts))
    return "\n\n".join(slides)


def analyze_with_gpt(pptx_text: str, sequences_fixes: str, nb_jours: int, api_key: str) -> list[dict]:
    client = OpenAI(api_key=api_key)

    prompt = f"""Tu es un expert en ingénierie de formation et certifications Qualiopi.

Tu dois construire un déroulé pédagogique complet sur {nb_jours} jour(s) en combinant deux sources :

---
SOURCE 1 — SÉQUENCES FIXES (toujours présentes, quelle que soit la formation) :
{sequences_fixes}
---
SOURCE 2 — CONTENU PÉDAGOGIQUE (extrait du support de formation) :
{pptx_text}
---

Ta mission :
1. Intègre TOUTES les séquences fixes aux bons moments de la journée (accueil en début, pauses le matin et l'après-midi, déjeuner à 12h30, ice breaker en début d'après-midi, positionnements en début/fin de journée, satisfaction en fin de journée, etc.).
2. Regroupe le contenu pédagogique du PPTX en séquences logiques et insère-les entre les séquences fixes.
3. Répartis les horaires sur {nb_jours} jour(s), journée 9h00 → 17h00 (1h déjeuner à 12h30, 2 pauses de 15 min).
4. Pour les séquences fixes (pauses, repas, tours de table) : objectifs/moyens/modalités adaptés et courts.
5. Pour les séquences de contenu : objectifs avec verbe d'action, moyens variés, modalités précises.

Réponds UNIQUEMENT avec un tableau JSON valide (sans markdown, sans texte autour) :
[
  {{
    "jour": "J1",
    "horaires": "09h00 - 09h30",
    "objectifs_pedagogiques": "...",
    "contenu_sequence": "...",
    "moyens_pedagogiques": "...",
    "modalites_validation": "..."
  }},
  ...
]"""

    response = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[{"role": "user", "content": prompt}],
        temperature=0.3,
        max_tokens=4000,
    )

    raw = response.choices[0].message.content.strip()
    # Nettoyer les éventuels blocs markdown
    if raw.startswith("```"):
        raw = raw.split("```")[1]
        if raw.startswith("json"):
            raw = raw[4:]
        raw = raw.strip()

    data = json.loads(raw)
    return [
        {
            "Jour": item.get("jour", ""),
            "Horaires": item.get("horaires", ""),
            "Objectifs pédagogiques": item.get("objectifs_pedagogiques", ""),
            "Contenu de la séquence": item.get("contenu_sequence", ""),
            "Moyens pédagogiques": item.get("moyens_pedagogiques", ""),
            "Modalités de validation des acquis": item.get("modalites_validation", ""),
        }
        for item in data
    ]


def generate_excel(meta: dict, df: pd.DataFrame) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Déroulé Pédagogique"

    dark  = "1F3864"
    mid   = "2F5496"
    light = "D6E4F7"

    f_dark  = PatternFill("solid", fgColor=dark)
    f_mid   = PatternFill("solid", fgColor=mid)
    f_light = PatternFill("solid", fgColor=light)
    f_white = PatternFill("solid", fgColor="FFFFFF")

    thin   = Side(border_style="thin",   color="AAAAAA")
    medium = Side(border_style="medium", color="555555")
    bdr    = Border(left=thin,   right=thin,   top=thin,   bottom=thin)
    bdr_h  = Border(left=medium, right=medium, top=medium, bottom=medium)

    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left   = Alignment(horizontal="left",   vertical="top",    wrap_text=True)

    # ── Titre ──────────────────────────────────────────────
    ws.merge_cells("A1:F1")
    c = ws["A1"]
    c.value     = f"DÉROULÉ PÉDAGOGIQUE — {meta['nom_formation'].upper()}"
    c.font      = Font(name="Calibri", bold=True, color="FFFFFF", size=14)
    c.fill      = f_dark
    c.alignment = center
    ws.row_dimensions[1].height = 32

    # ── Métadonnées ────────────────────────────────────────
    meta_items = [
        ("A2:B2", f"Formateur : {meta['formateur']}"),
        ("C2:D2", f"Dates : {meta['dates']}"),
        ("E2:F2", f"Lieu : {meta['lieu']}  |  Stagiaires : {meta['nb_stagiaires']}"),
    ]
    for merge_range, val in meta_items:
        ws.merge_cells(merge_range)
        first_cell = merge_range.split(":")[0]
        c = ws[first_cell]
        c.value     = val
        c.font      = Font(name="Calibri", color="FFFFFF", size=10)
        c.fill      = f_mid
        c.alignment = center
    ws.row_dimensions[2].height = 22

    # ── Séquences fixes (info dans l'export) ──────────────
    next_row = 3
    if meta.get("sequences_fixes"):
        ws.merge_cells("A3:F3")
        c = ws["A3"]
        c.value     = f"Séquences fixes intégrées : {meta['sequences_fixes'].replace(chr(10), ' | ')}"
        c.font      = Font(name="Calibri", italic=True, size=9, color=dark)
        c.alignment = left
        c.border    = bdr
        ws.row_dimensions[3].height = 30
        next_row = 4

    # ── En-têtes colonnes ──────────────────────────────────
    headers = [
        ("JOUR", 8),
        ("HORAIRES", 13),
        ("OBJECTIFS PÉDAGOGIQUES", 32),
        ("CONTENU DE LA SÉQUENCE", 44),
        ("MOYENS PÉDAGOGIQUES\nComment se fait la transmission ?", 30),
        ("MODALITÉS DE VALIDATION DES ACQUIS\nComment évaluer l'atteinte de l'objectif ?", 32),
    ]
    hr = next_row
    for ci, (label, width) in enumerate(headers, 1):
        c = ws.cell(row=hr, column=ci, value=label)
        c.font      = Font(name="Calibri", bold=True, color="FFFFFF", size=10)
        c.fill      = f_mid
        c.alignment = center
        c.border    = bdr_h
        ws.column_dimensions[get_column_letter(ci)].width = width
    ws.row_dimensions[hr].height = 45

    # ── Données ────────────────────────────────────────────
    for ri, (_, row) in enumerate(df.iterrows()):
        er   = hr + 1 + ri
        fill = f_light if ri % 2 == 0 else f_white
        for ci, col in enumerate(COLS, 1):
            c = ws.cell(row=er, column=ci, value=str(row.get(col, "") or ""))
            c.font      = Font(name="Calibri", size=10)
            c.fill      = fill
            c.alignment = left
            c.border    = bdr
        ws.row_dimensions[er].height = 65

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.getvalue()


# ─────────────────────────────────────────────
# UI
# ─────────────────────────────────────────────

st.title("📋 Déroulé Pédagogique Qualiopi")
st.caption("Importe ton PPTX → l'IA remplit le tableau Qualiopi → tu exportes en Excel.")

# ── Sidebar : clé API ──────────────────────────────────────
with st.sidebar:
    st.header("⚙️ Configuration")
    try:
        has_secret = bool(st.secrets.get("OPENAI_API_KEY", ""))
    except Exception:
        has_secret = False

    if not has_secret:
        st.session_state["api_key"] = st.text_input(
            "Clé API OpenAI",
            type="password",
            placeholder="sk-...",
            help="Ta clé reste dans ton navigateur, elle n'est pas enregistrée.",
        )
    else:
        st.success("✅ Clé API configurée")

    st.divider()
    st.caption("💡 Pour déployer sur Streamlit Cloud, ajoute `OPENAI_API_KEY` dans les Secrets du projet.")

# ── Métadonnées ────────────────────────────────────────────
with st.expander("ℹ️ Informations de la formation", expanded=True):
    c1, c2, c3 = st.columns(3)
    with c1:
        nom_formation = st.text_input("Nom de la formation", placeholder="Ex : Management - Édouard")
        formateur     = st.text_input("Formateur", placeholder="Ex : Jean Dupont")
    with c2:
        date_debut = st.date_input("Date de début", value=datetime.date.today())
        date_fin   = st.date_input("Date de fin",   value=datetime.date.today())
        nb_jours   = max(1, (date_fin - date_debut).days + 1)
        st.caption(f"Durée calculée : **{nb_jours} jour(s)**")
    with c3:
        lieu          = st.text_input("Lieu", placeholder="Ex : Paris / Distanciel")
        nb_stagiaires = st.number_input("Nombre de stagiaires", min_value=1, value=10)
    sequences_fixes = st.text_area(
        "Séquences fixes (une par ligne)",
        placeholder="Tour de table de présentation (ouverture)\nPositionnement sur le sujet (début de journée)\nPause de 15 min (matin)\nDéjeuner d'1H\nRupture pédagogique (ice breaker) en début d'am\nPause de 15 min (après-midi)\nPositionnement sur le sujet (fin de journée)\nTour de table de satisfaction (fin de journée)",
        height=160,
        help="Ces séquences seront insérées aux bons moments par l'IA, quelle que soit la formation.",
    )

st.divider()

# ── Étape 1 : Upload PPTX ─────────────────────────────────
st.subheader("1 — Importer le support de formation")
uploaded = st.file_uploader("Glisse ton fichier .pptx ici", type=["pptx"], label_visibility="collapsed")

if "pptx_text" not in st.session_state:
    st.session_state["pptx_text"] = ""
if "df" not in st.session_state:
    st.session_state["df"] = pd.DataFrame(columns=COLS)

if uploaded:
    pptx_text = extract_pptx_text(uploaded.read())
    st.session_state["pptx_text"] = pptx_text
    nb_slides = pptx_text.count("=== Slide ")
    st.success(f"✅ {nb_slides} slides chargées — prêt pour l'analyse IA")

st.divider()

# ── Étape 2 : Analyse IA ──────────────────────────────────
st.subheader("2 — Générer le déroulé avec l'IA")

api_key = get_api_key()
can_generate = bool(api_key) and bool(st.session_state["pptx_text"])

if not api_key:
    st.warning("⚠️ Entre ta clé API OpenAI dans la barre latérale.")
elif not st.session_state["pptx_text"]:
    st.info("⬆️ Importe d'abord un fichier PPTX.")

if st.button("✨ Générer le déroulé pédagogique", type="primary", disabled=not can_generate):
    with st.spinner("L'IA analyse le support et remplit le tableau... (10-20 secondes)"):
        try:
            rows = analyze_with_gpt(st.session_state["pptx_text"], sequences_fixes, nb_jours, api_key)
            st.session_state["df"] = pd.DataFrame(rows)
            st.success(f"✅ {len(rows)} séquences générées !")
        except json.JSONDecodeError as e:
            st.error(f"Erreur de format JSON retourné par GPT : {e}")
        except Exception as e:
            st.error(f"Erreur : {e}")

st.divider()

# ── Étape 3 : Tableau éditable ────────────────────────────
st.subheader("3 — Vérifier et ajuster")
st.caption("Modifie les cellules directement. Tu peux ajouter ou supprimer des lignes.")

edited = st.data_editor(
    st.session_state["df"],
    num_rows="dynamic",
    use_container_width=True,
    height=500,
    column_config={
        "Jour":                               st.column_config.TextColumn(width=70),
        "Horaires":                           st.column_config.TextColumn(width=120),
        "Objectifs pédagogiques":             st.column_config.TextColumn(width=240),
        "Contenu de la séquence":             st.column_config.TextColumn(width=340),
        "Moyens pédagogiques":                st.column_config.TextColumn(width=220),
        "Modalités de validation des acquis": st.column_config.TextColumn(width=220),
    },
)
st.session_state["df"] = edited

st.divider()

# ── Étape 4 : Export Excel ────────────────────────────────
st.subheader("4 — Exporter en Excel")

if st.button("📥 Générer le fichier Excel Qualiopi", type="primary"):
    if not nom_formation.strip():
        st.warning("⚠️ Renseigne le nom de la formation.")
    elif edited.empty:
        st.warning("⚠️ Le tableau est vide.")
    else:
        meta = {
            "nom_formation":  nom_formation,
            "formateur":      formateur,
            "dates":          f"{date_debut.strftime('%d/%m/%Y')} → {date_fin.strftime('%d/%m/%Y')}",
            "lieu":           lieu,
            "nb_stagiaires":  nb_stagiaires,
            "sequences_fixes": sequences_fixes,
        }
        excel_bytes = generate_excel(meta, edited)
        filename = (
            f"Deroulé_Pédagogique_{nom_formation.replace(' ', '_')}"
            f"_{date_debut.strftime('%Y%m%d')}.xlsx"
        )
        st.download_button(
            label="💾 Télécharger l'Excel",
            data=excel_bytes,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )
