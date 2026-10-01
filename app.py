"""
Suivi des heures de vol et des cycles de la flotte aérienne
===========================================================
Application Streamlit : nettoyage des données, KPIs, tableau de bord,
filtres, authentification et animation « Bar Chart Race ».

Lancement :  streamlit run app.py
"""

import hashlib
import hmac
import io
import re
import unicodedata
from pathlib import Path

import matplotlib

matplotlib.use("Agg")  # backend sans interface graphique (indispensable côté serveur)
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import seaborn as sns
import streamlit as st
from matplotlib.patches import Patch
from matplotlib.ticker import FuncFormatter
from PIL import Image

# ──────────────────────────────────────────────────────────────────────────────
# CONSTANTES
# ──────────────────────────────────────────────────────────────────────────────
COL_ID = "Immatriculation"
COL_MODELE = "Modèle"
COL_HEURES = "Heures de vol"
COL_CYCLES = "Cycles"
COL_STATUT = "Statut"
COL_DATE = "Date"
COL_COMPAGNIE = "Compagnie"
COL_RATIO = "Heures par cycle"

SAMPLE_FILE = Path(__file__).parent / "suivi_heures_vol_cycles.xlsx"

MOIS_FR = ["janv.", "févr.", "mars", "avr.", "mai", "juin",
           "juil.", "août", "sept.", "oct.", "nov.", "déc."]

STATUS_COLORS = {
    "Opérationnel": "#16a34a",
    "En maintenance": "#f59e0b",
    "En révision": "#2563eb",
}
FALLBACK_COLOR = "#94a3b8"

# Mots-clés permettant de reconnaître les colonnes quel que soit leur intitulé
COLUMN_KEYWORDS = {
    COL_ID: ["immatriculation", "immat", "registration", "tail"],
    COL_MODELE: ["modele", "model", "typeavion", "type"],
    COL_HEURES: ["heure", "hours"],
    COL_CYCLES: ["cycle", "atterrissage", "landing"],
    COL_STATUT: ["statut", "status", "etat"],
    COL_DATE: ["date", "controle", "control"],
    COL_COMPAGNIE: ["compagnie", "airline", "operateur"],
}
REQUIRED = [COL_ID, COL_HEURES, COL_CYCLES]


# ──────────────────────────────────────────────────────────────────────────────
# OUTILS GÉNÉRAUX
# ──────────────────────────────────────────────────────────────────────────────
def fmt(n) -> str:
    """Formate un nombre avec séparateur de milliers à la française."""
    if n is None or (isinstance(n, float) and np.isnan(n)):
        return "—"
    return f"{n:,.0f}".replace(",", " ")


def date_fr(ts) -> str:
    ts = pd.Timestamp(ts)
    return f"{MOIS_FR[ts.month - 1]} {ts.year}"


def _normalize(text: str) -> str:
    text = unicodedata.normalize("NFKD", str(text))
    text = "".join(c for c in text if not unicodedata.combining(c))
    return re.sub(r"[^a-z0-9]", "", text.lower())


def status_palette(statuts) -> dict:
    return {s: STATUS_COLORS.get(s, FALLBACK_COLOR) for s in statuts}


# ──────────────────────────────────────────────────────────────────────────────
# AUTHENTIFICATION
# ──────────────────────────────────────────────────────────────────────────────
def _hash(password: str) -> str:
    return hashlib.sha256(password.encode("utf-8")).hexdigest()


def get_users() -> dict:
    """
    Utilisateurs autorisés {identifiant: empreinte SHA-256}.
    Définis dans .streamlit/secrets.toml (section [users]) ;
    à défaut, comptes de démonstration (à changer en production).
    """
    try:
        users = {str(k): str(v) for k, v in st.secrets["users"].items()}
        if users:
            return users
    except Exception:
        pass
    return {"admin": _hash("flotte2026"), "technicien": _hash("avion2026")}


def check_credentials(username: str, password: str) -> bool:
    users = get_users()
    expected = users.get(username.strip())
    if expected is None:
        _hash(password)  # temps de traitement comparable
        return False
    return hmac.compare_digest(_hash(password), expected)


def login_page() -> None:
    _, centre, _ = st.columns([1, 1.2, 1])
    with centre:
        st.markdown(
            """<div class="login-card">
            <div class="login-title">Suivi de la flotte</div>
            <div class="login-sub">Connectez-vous pour accéder au tableau de bord</div>
            </div>""",
            unsafe_allow_html=True,
        )
        with st.form("login_form"):
            username = st.text_input("Identifiant")
            password = st.text_input("Mot de passe", type="password")
            submitted = st.form_submit_button("Se connecter", width="stretch")
        if submitted:
            if check_credentials(username, password):
                st.session_state["auth"] = True
                st.session_state["user"] = username.strip()
                st.rerun()
            else:
                st.error("Identifiant ou mot de passe incorrect.")


# ──────────────────────────────────────────────────────────────────────────────
# NETTOYAGE DES DONNÉES
# ──────────────────────────────────────────────────────────────────────────────
def detect_columns(df: pd.DataFrame) -> dict:
    """Associe chaque colonne du fichier à un nom canonique."""
    mapping, used = {}, set()
    normalized = {c: _normalize(c) for c in df.columns}
    for canon, keywords in COLUMN_KEYWORDS.items():
        for col, norm in normalized.items():
            if col in used or norm == "id":
                continue
            if any(k in norm for k in keywords):
                mapping[col] = canon
                used.add(col)
                break
    return mapping


def _to_number(series: pd.Series) -> pd.Series:
    if pd.api.types.is_numeric_dtype(series):
        return series.astype(float)
    cleaned = (series.astype(str)
               .str.replace(r"[\s\u00a0\u202f]", "", regex=True)
               .str.replace(",", ".", regex=False))
    return pd.to_numeric(cleaned, errors="coerce")


def clean_data(raw: pd.DataFrame):
    """
    Nettoie le fichier brut.
    Retourne (df_propre, rapport[list[(étape, résultat)]], colonnes_manquantes).
    """
    report = []
    df = raw.copy()
    report.append(("Lignes dans le fichier", f"{len(df)} lignes, {df.shape[1]} colonnes"))

    # 1. Colonnes
    mapping = detect_columns(df)
    df = df.rename(columns=mapping)
    missing = [c for c in REQUIRED if c not in df.columns]
    if missing:
        return None, report, missing
    recognized = ", ".join(f"{src} → {dst}" for src, dst in mapping.items())
    report.append(("Colonnes reconnues", recognized))
    for optional, default in [(COL_MODELE, "Non renseigné"),
                              (COL_STATUT, "Inconnu"),
                              (COL_COMPAGNIE, "Non renseignée")]:
        if optional not in df.columns:
            df[optional] = default
            report.append((f"Colonne « {optional} » absente", f"valeur « {default} » appliquée"))
    if COL_DATE not in df.columns:
        df[COL_DATE] = pd.NaT
        report.append(("Colonne « Date » absente", "aucun historique temporel détecté"))
    df = df[[COL_ID, COL_MODELE, COL_COMPAGNIE, COL_STATUT, COL_DATE, COL_HEURES, COL_CYCLES]]

    # 2. Lignes vides
    n = len(df)
    df = df.dropna(how="all")
    report.append(("Lignes entièrement vides supprimées", str(n - len(df))))

    # 3. Texte : espaces, casse, valeurs vides
    for col in [COL_ID, COL_MODELE, COL_COMPAGNIE, COL_STATUT]:
        df[col] = df[col].astype("string").str.strip()
        df[col] = df[col].replace("", pd.NA)
    df[COL_ID] = df[COL_ID].str.upper()
    df[COL_STATUT] = df[COL_STATUT].str.capitalize()
    df[COL_MODELE] = df[COL_MODELE].fillna("Non renseigné")
    df[COL_COMPAGNIE] = df[COL_COMPAGNIE].fillna("Non renseignée")
    df[COL_STATUT] = df[COL_STATUT].fillna("Inconnu")
    report.append(("Textes normalisés", "espaces supprimés, immatriculations en majuscules, statuts harmonisés"))

    # 4. Types numériques et dates
    df[COL_HEURES] = _to_number(df[COL_HEURES])
    df[COL_CYCLES] = _to_number(df[COL_CYCLES])
    df[COL_DATE] = pd.to_datetime(df[COL_DATE], errors="coerce")
    report.append(("Conversion des types", "heures et cycles en nombres, date en format date"))

    # 5. Immatriculation manquante
    n = len(df)
    df = df.dropna(subset=[COL_ID])
    report.append(("Lignes sans immatriculation supprimées", str(n - len(df))))

    # 6. Valeurs négatives
    neg = int(((df[COL_HEURES] < 0) | (df[COL_CYCLES] < 0)).sum())
    df.loc[df[COL_HEURES] < 0, COL_HEURES] = np.nan
    df.loc[df[COL_CYCLES] < 0, COL_CYCLES] = np.nan
    report.append(("Valeurs négatives invalidées", str(neg)))

    # 7. Doublons (même avion, même date)
    n = len(df)
    df = df.sort_values([COL_ID, COL_DATE]).drop_duplicates(subset=[COL_ID, COL_DATE], keep="last")
    report.append(("Doublons supprimés (même avion et même date)", str(n - len(df))))

    # 8. Valeurs manquantes : report de la dernière valeur connue de l'avion
    n_na = int(df[[COL_HEURES, COL_CYCLES]].isna().sum().sum())
    df[[COL_HEURES, COL_CYCLES]] = df.groupby(COL_ID)[[COL_HEURES, COL_CYCLES]].ffill()
    n = len(df)
    df = df.dropna(subset=[COL_HEURES, COL_CYCLES])
    report.append(("Valeurs manquantes (heures/cycles)",
                   f"{n_na} détectées, {n - len(df)} ligne(s) irrécupérable(s) supprimée(s)"))

    # 9. Cohérence : les cumuls ne peuvent pas diminuer dans le temps
    before = df[[COL_HEURES, COL_CYCLES]].copy()
    df = df.sort_values([COL_ID, COL_DATE])
    df[[COL_HEURES, COL_CYCLES]] = df.groupby(COL_ID)[[COL_HEURES, COL_CYCLES]].cummax()
    corrected = int((before.loc[df.index] != df[[COL_HEURES, COL_CYCLES]]).any(axis=1).sum())
    report.append(("Cumuls décroissants corrigés", str(corrected)))

    df[COL_HEURES] = df[COL_HEURES].round(0).astype(int)
    df[COL_CYCLES] = df[COL_CYCLES].round(0).astype(int)
    for col in [COL_ID, COL_MODELE, COL_COMPAGNIE, COL_STATUT]:
        df[col] = df[col].astype(str)

    # 10. Alerte (non bloquante) : durée moyenne par cycle aberrante
    with np.errstate(divide="ignore", invalid="ignore"):
        ratio = df[COL_HEURES] / df[COL_CYCLES].replace(0, np.nan)
    suspects = int((ratio > 20).sum())
    report.append(("Alerte : plus de 20 h par cycle (à vérifier)", str(suspects)))

    df = df.reset_index(drop=True)
    report.append(("Lignes après nettoyage", f"{len(df)} lignes, {df[COL_ID].nunique()} aéronefs"))
    return df, report, []


def last_snapshot(df: pd.DataFrame) -> pd.DataFrame:
    """Dernière situation connue de chaque aéronef + ratio heures/cycle."""
    snap = (df.sort_values([COL_ID, COL_DATE], na_position="first")
              .groupby(COL_ID, as_index=False).tail(1)
              .reset_index(drop=True))
    snap[COL_RATIO] = (snap[COL_HEURES] / snap[COL_CYCLES].replace(0, np.nan)).round(2)
    return snap


def simulate_history(snap: pd.DataFrame, months: int, age_months: int, seed: int = 42) -> pd.DataFrame:
    """
    Reconstitue un historique mensuel PLAUSIBLE à partir d'une situation unique.
    Chaque avion cumule en moyenne (total / ancienneté) par mois, avec une variation
    aléatoire reproductible. Les cumuls sont croissants et se terminent sur les
    valeurs réelles du fichier.
    """
    rng = np.random.default_rng(seed)
    ref = snap[COL_DATE].max()
    ref = pd.Timestamp.today() if pd.isna(ref) else ref
    end = pd.Timestamp(ref).normalize().replace(day=1)
    dates = pd.date_range(end=end, periods=months, freq="MS")

    rows = []
    for _, r in snap.iterrows():
        weights = rng.uniform(0.6, 1.4, size=months - 1)
        steps_h = weights * r[COL_HEURES] / age_months
        steps_c = weights * r[COL_CYCLES] / age_months
        # cumul en remontant le temps depuis la valeur réelle
        h = r[COL_HEURES] - np.concatenate([np.cumsum(steps_h[::-1])[::-1], [0]])
        c = r[COL_CYCLES] - np.concatenate([np.cumsum(steps_c[::-1])[::-1], [0]])
        h, c = np.clip(h, 0, None), np.clip(c, 0, None)
        for d, hv, cv in zip(dates, h, c):
            rows.append({COL_ID: r[COL_ID], COL_MODELE: r[COL_MODELE],
                         COL_COMPAGNIE: r[COL_COMPAGNIE], COL_STATUT: r[COL_STATUT],
                         COL_DATE: d, COL_HEURES: int(round(hv)), COL_CYCLES: int(round(cv))})
    return pd.DataFrame(rows)


# ──────────────────────────────────────────────────────────────────────────────
# CHARGEMENT (avec cache)
# ──────────────────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def load_file(content: bytes, filename: str) -> pd.DataFrame:
    if filename.lower().endswith(".csv"):
        return pd.read_csv(io.BytesIO(content), sep=None, engine="python")
    return pd.read_excel(io.BytesIO(content))


@st.cache_data(show_spinner=False)
def cached_clean(raw: pd.DataFrame):
    return clean_data(raw)


# ──────────────────────────────────────────────────────────────────────────────
# GRAPHIQUES (matplotlib + seaborn)
# ──────────────────────────────────────────────────────────────────────────────
def setup_style() -> None:
    sns.set_theme(style="whitegrid", context="notebook", rc={
        "axes.spines.top": False, "axes.spines.right": False,
        "axes.titleweight": "bold", "axes.titlesize": 12, "axes.titlelocation": "left",
        "grid.color": "#e5e7eb", "figure.facecolor": "white",
    })


def show(fig) -> None:
    fig.tight_layout()
    st.pyplot(fig)
    plt.close(fig)


def chart_hours_by_aircraft(snap: pd.DataFrame):
    d = snap.sort_values(COL_HEURES, ascending=False)
    fig, ax = plt.subplots(figsize=(7, max(3.2, 0.42 * len(d) + 1)))
    sns.barplot(data=d, y=COL_ID, x=COL_HEURES, hue=COL_STATUT,
                palette=status_palette(d[COL_STATUT].unique()), dodge=False, ax=ax)
    for p in ax.patches:
        w = p.get_width()
        if np.isfinite(w) and w > 0:
            ax.text(w, p.get_y() + p.get_height() / 2, f"  {fmt(w)} h", va="center", fontsize=9)
    ax.set_title("Heures de vol cumulées par aéronef")
    ax.set_xlabel("Heures de vol"); ax.set_ylabel("")
    ax.set_xlim(0, d[COL_HEURES].max() * 1.18)
    ax.legend(title="Statut", loc="lower right", frameon=False)
    return fig


def chart_cycles_by_aircraft(snap: pd.DataFrame):
    d = snap.sort_values(COL_CYCLES, ascending=False)
    fig, ax = plt.subplots(figsize=(7, max(3.2, 0.42 * len(d) + 1)))
    sns.barplot(data=d, y=COL_ID, x=COL_CYCLES, hue=COL_ID, palette="crest_r", legend=False, ax=ax)
    for p in ax.patches:
        w = p.get_width()
        if np.isfinite(w) and w > 0:
            ax.text(w, p.get_y() + p.get_height() / 2, f"  {fmt(w)}", va="center", fontsize=9)
    ax.set_title("Cycles (atterrissages) cumulés par aéronef")
    ax.set_xlabel("Cycles"); ax.set_ylabel("")
    ax.set_xlim(0, d[COL_CYCLES].max() * 1.15)
    return fig


def chart_status_donut(snap: pd.DataFrame):
    counts = snap[COL_STATUT].value_counts()
    fig, ax = plt.subplots(figsize=(5, 4))
    pal = status_palette(counts.index)
    ax.pie(counts.values, labels=counts.index, colors=[pal[s] for s in counts.index],
           autopct=lambda p: f"{p:.0f} %", startangle=90, pctdistance=0.8,
           wedgeprops=dict(width=0.42, edgecolor="white"), textprops=dict(fontsize=10))
    ax.text(0, 0, f"{len(snap)}\naéronefs", ha="center", va="center", fontsize=13, fontweight="bold")
    ax.set_title("Répartition par statut opérationnel")
    return fig


def chart_scatter(snap: pd.DataFrame):
    fig, ax = plt.subplots(figsize=(7, 4.6))
    sns.scatterplot(data=snap, x=COL_CYCLES, y=COL_HEURES, hue=COL_MODELE, s=150,
                    edgecolor="white", ax=ax)
    for _, r in snap.iterrows():
        ax.annotate(r[COL_ID], (r[COL_CYCLES], r[COL_HEURES]), textcoords="offset points",
                    xytext=(7, 6), fontsize=8, color="#475569")
    total_c = snap[COL_CYCLES].sum()
    if total_c > 0:
        ratio = snap[COL_HEURES].sum() / total_c
        xs = np.array([0, snap[COL_CYCLES].max() * 1.05])
        ax.plot(xs, xs * ratio, "--", color="#94a3b8", lw=1.2,
                label=f"Moyenne flotte ({ratio:.2f} h/cycle)")
    ax.set_title("Heures de vol en fonction des cycles")
    ax.set_xlabel("Cycles (atterrissages)"); ax.set_ylabel("Heures de vol")
    ax.legend(fontsize=8, frameon=False, loc="upper left", bbox_to_anchor=(1, 1))
    return fig


def chart_ratio(snap: pd.DataFrame):
    d = snap.sort_values(COL_RATIO, ascending=False)
    fig, ax = plt.subplots(figsize=(7, max(3.2, 0.42 * len(d) + 1)))
    sns.barplot(data=d, y=COL_ID, x=COL_RATIO, hue=COL_ID, palette="flare", legend=False, ax=ax)
    for p in ax.patches:
        w = p.get_width()
        if np.isfinite(w) and w > 0:
            ax.text(w, p.get_y() + p.get_height() / 2, f"  {w:.2f} h", va="center", fontsize=9)
    ax.set_title("Durée moyenne de vol par cycle")
    ax.set_xlabel("Heures par cycle"); ax.set_ylabel("")
    ax.set_xlim(0, d[COL_RATIO].max() * 1.18)
    return fig


def chart_by_company(snap: pd.DataFrame):
    fig, ax = plt.subplots(figsize=(6, 4))
    sns.barplot(data=snap, x=COL_COMPAGNIE, y=COL_HEURES, hue=COL_COMPAGNIE, estimator="mean",
                errorbar=None, palette="Blues_d", legend=False, ax=ax)
    for p in ax.patches:
        h = p.get_height()
        if np.isfinite(h) and h > 0:
            ax.text(p.get_x() + p.get_width() / 2, h, fmt(h), ha="center", va="bottom", fontsize=9)
    ax.set_title("Heures de vol moyennes par compagnie")
    ax.set_xlabel(""); ax.set_ylabel("Heures de vol (moyenne)")
    return fig


def chart_box_model(snap: pd.DataFrame):
    order = snap.groupby(COL_MODELE)[COL_HEURES].mean().sort_values(ascending=False).index
    fig, ax = plt.subplots(figsize=(7, max(3.2, 0.4 * len(order) + 1)))
    sns.barplot(data=snap, y=COL_MODELE, x=COL_HEURES, hue=COL_MODELE, order=order, estimator="mean",
                errorbar=None, palette="viridis", legend=False, ax=ax)
    ax.set_title("Heures de vol moyennes par modèle")
    ax.set_xlabel("Heures de vol (moyenne)"); ax.set_ylabel("")
    return fig


def chart_heatmap(snap: pd.DataFrame):
    corr = snap[[COL_HEURES, COL_CYCLES, COL_RATIO]].corr()
    fig, ax = plt.subplots(figsize=(5, 4))
    sns.heatmap(corr, annot=True, fmt=".2f", cmap="vlag", vmin=-1, vmax=1, square=True,
                cbar=False, linewidths=2, linecolor="white", ax=ax)
    ax.set_title("Corrélations")
    return fig


def chart_distribution(snap: pd.DataFrame):
    fig, ax = plt.subplots(figsize=(6, 4))
    sns.histplot(snap[COL_HEURES], bins=min(10, max(3, len(snap))), kde=len(snap) > 3,
                 color="#2563eb", ax=ax)
    ax.axvline(snap[COL_HEURES].mean(), color="#ef4444", ls="--", lw=1.3,
               label=f"Moyenne : {fmt(snap[COL_HEURES].mean())} h")
    ax.set_title("Distribution des heures de vol")
    ax.set_xlabel("Heures de vol"); ax.set_ylabel("Nombre d'aéronefs")
    ax.legend(frameon=False)
    return fig


def chart_evolution(hist: pd.DataFrame, metric: str):
    fig, axes = plt.subplots(1, 2, figsize=(12, 4.2), gridspec_kw={"width_ratios": [1, 1.4]})
    total = hist.groupby(COL_DATE)[metric].sum().reset_index()
    sns.lineplot(data=total, x=COL_DATE, y=metric, marker="o", color="#2563eb", ax=axes[0])
    axes[0].fill_between(total[COL_DATE], total[metric], alpha=0.12, color="#2563eb")
    axes[0].set_title(f"{metric} cumulés de la flotte")
    axes[0].set_xlabel(""); axes[0].set_ylabel(metric)
    sns.lineplot(data=hist, x=COL_DATE, y=metric, hue=COL_ID, ax=axes[1], linewidth=1.6)
    axes[1].set_title(f"{metric} cumulés par aéronef")
    axes[1].set_xlabel(""); axes[1].set_ylabel("")
    axes[1].legend(fontsize=7, ncol=2, frameon=False, loc="upper left", bbox_to_anchor=(1, 1))
    for ax in axes:
        ax.tick_params(axis="x", rotation=30)
        ax.yaxis.set_major_formatter(FuncFormatter(lambda x, _: fmt(x)))
    return fig


# ──────────────────────────────────────────────────────────────────────────────
# BAR CHART RACE (animation GIF générée avec matplotlib, sans ffmpeg)
# ──────────────────────────────────────────────────────────────────────────────
def build_color_meta(snap: pd.DataFrame) -> dict:
    """Associe chaque immatriculation à (modèle, couleur) ; couleur = modèle."""
    models = sorted(snap[COL_MODELE].unique())
    colors = sns.color_palette("tab10" if len(models) <= 10 else "husl", len(models))
    model_color = {m: c for m, c in zip(models, colors)}
    return {r[COL_ID]: (r[COL_MODELE], tuple(model_color[r[COL_MODELE]])) for _, r in snap.iterrows()}


@st.cache_data(show_spinner=False)
def render_race(pivot: pd.DataFrame, meta: dict, title: str, n_bars: int,
                steps: int, fps: int) -> bytes:
    """Dessine chaque image de la course et assemble un GIF animé."""
    dates = list(pivot.index)
    n = len(dates)
    x_old = np.arange(n)
    x_new = np.linspace(0, n - 1, (n - 1) * steps + 1)
    arr = np.column_stack([np.interp(x_new, x_old, pivot[c].to_numpy(dtype=float))
                           for c in pivot.columns])
    cols = np.array(pivot.columns)
    n_bars = min(n_bars, len(cols))

    fig, ax = plt.subplots(figsize=(10, 5.6), dpi=80)
    fig.subplots_adjust(left=0.11, right=0.97, top=0.90, bottom=0.15)
    legend_items = {}
    for model, color in meta.values():
        legend_items[model] = color
    fig.legend(handles=[Patch(color=c, label=m) for m, c in sorted(legend_items.items())],
               loc="lower center", ncol=min(5, len(legend_items)), frameon=False, fontsize=9)

    frames = []
    for i, row in enumerate(arr):
        idx = np.argsort(row)[::-1][:n_bars]
        vals, names = row[idx], cols[idx]
        ypos = np.arange(n_bars)[::-1]
        ax.clear()
        ax.barh(ypos, vals, height=0.78, color=[meta[nm][1] for nm in names])
        ax.set_yticks(ypos)
        ax.set_yticklabels(names, fontsize=11, fontweight="bold")
        xmax = max(vals.max() * 1.2, 1)
        ax.set_xlim(0, xmax)
        ax.set_ylim(-0.6, n_bars - 0.4)
        for y, v in zip(ypos, vals):
            ax.text(v + xmax * 0.008, y, fmt(v), va="center", fontsize=10, color="#334155")
        ax.xaxis.set_major_formatter(FuncFormatter(lambda x, _: fmt(x)))
        ax.grid(axis="x", color="#e5e7eb")
        ax.set_axisbelow(True)
        ax.tick_params(length=0)
        for s in ax.spines.values():
            s.set_visible(False)
        label_date = dates[min(int(round(i / steps)), n - 1)]
        ax.text(0.97, 0.10, date_fr(label_date), transform=ax.transAxes, ha="right",
                fontsize=34, fontweight="bold", color="#cbd5e1")
        ax.set_title(title, loc="left", fontsize=14, fontweight="bold", color="#0f172a")
        fig.canvas.draw()
        frames.append(Image.fromarray(np.asarray(fig.canvas.buffer_rgba())[:, :, :3].copy()))
    plt.close(fig)

    frames += [frames[-1]] * int(fps * 1.5)  # pause sur le classement final
    buf = io.BytesIO()
    frames[0].save(buf, format="GIF", save_all=True, append_images=frames[1:],
                   duration=int(1000 / fps), loop=0)
    return buf.getvalue()


def make_race_pivot(hist: pd.DataFrame, metric: str) -> pd.DataFrame:
    pivot = hist.pivot_table(index=COL_DATE, columns=COL_ID, values=metric, aggfunc="last").sort_index()
    if len(pivot) > 60:  # historique très fin : agrégation mensuelle
        pivot = pivot.resample("MS").last()
    return pivot.interpolate(limit_area="inside").ffill().fillna(0)


# ──────────────────────────────────────────────────────────────────────────────
# STYLE DE L'INTERFACE
# ──────────────────────────────────────────────────────────────────────────────
CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
.block-container { padding-top: 1.6rem; max-width: 1400px; }
.hero { background: linear-gradient(120deg, #0f172a 0%, #1e3a8a 60%, #2563eb 100%);
        padding: 1.6rem 2rem; border-radius: 16px; color: white; margin-bottom: 1.2rem; }
.hero h1 { color: white; margin: 0; font-size: 1.8rem; font-weight: 700; }
.hero p { color: #cbd5e1; margin: .3rem 0 0 0; font-size: .95rem; }
.kpi { background: white; border: 1px solid #e5e7eb; border-left: 5px solid #2563eb;
       border-radius: 12px; padding: .9rem 1.1rem; box-shadow: 0 1px 3px rgba(15,23,42,.06);
       height: 100%; }
.kpi.green { border-left-color: #16a34a; } .kpi.orange { border-left-color: #f59e0b; }
.kpi.violet { border-left-color: #7c3aed; } .kpi.red { border-left-color: #ef4444; }
.kpi-label { color: #64748b; font-size: .78rem; font-weight: 600; text-transform: uppercase;
             letter-spacing: .04em; }
.kpi-value { color: #0f172a; font-size: 1.7rem; font-weight: 700; line-height: 1.25; }
.kpi-sub { color: #94a3b8; font-size: .78rem; }
.section-title { font-size: 1.05rem; font-weight: 600; color: #0f172a; margin: 1.2rem 0 .4rem 0; }
.login-card { text-align: center; margin: 4rem 0 1rem 0; }
.login-title { font-size: 1.9rem; font-weight: 700; color: #0f172a; }
.login-sub { color: #64748b; margin-top: .3rem; }
.stTabs [data-baseweb="tab"] { font-weight: 600; padding: .6rem 1.1rem; }
</style>
"""


def kpi(label: str, value: str, sub: str = "", color: str = "") -> str:
    return (f'<div class="kpi {color}"><div class="kpi-label">{label}</div>'
            f'<div class="kpi-value">{value}</div><div class="kpi-sub">{sub}</div></div>')


# ──────────────────────────────────────────────────────────────────────────────
# APPLICATION
# ──────────────────────────────────────────────────────────────────────────────
def sidebar_data():
    """Importation + nettoyage. Retourne (df propre, rapport) ou arrête l'app."""
    st.sidebar.markdown("### Données")
    uploaded = st.sidebar.file_uploader("Fichier Excel ou CSV", type=["xlsx", "xls", "csv"])
    if uploaded is not None:
        raw = load_file(uploaded.getvalue(), uploaded.name)
        source = uploaded.name
    elif SAMPLE_FILE.exists():
        raw = load_file(SAMPLE_FILE.read_bytes(), SAMPLE_FILE.name)
        source = f"{SAMPLE_FILE.name} (exemple)"
    else:
        st.info("Importez un fichier Excel ou CSV dans le menu de gauche pour commencer.")
        st.stop()
    st.sidebar.caption(f"Source : {source}")

    clean, report, missing = cached_clean(raw)
    if clean is None:
        st.error("Colonnes obligatoires introuvables : " + ", ".join(missing) +
                 ". Le fichier doit contenir l'immatriculation, les heures de vol et les cycles.")
        st.dataframe(raw.head(), width="stretch")
        st.stop()
    return clean, report


def sidebar_filters(clean: pd.DataFrame):
    """Filtres. Retourne (historique filtré, situation actuelle filtrée, mode_historique)."""
    is_history = clean.groupby(COL_ID)[COL_DATE].nunique().max() > 1
    st.sidebar.markdown("### Filtres")

    data = clean
    if is_history:
        dmin, dmax = clean[COL_DATE].min().date(), clean[COL_DATE].max().date()
        period = st.sidebar.date_input("Période", (dmin, dmax), min_value=dmin, max_value=dmax)
        if isinstance(period, (tuple, list)) and len(period) == 2:
            data = clean[(clean[COL_DATE] >= pd.Timestamp(period[0])) &
                         (clean[COL_DATE] <= pd.Timestamp(period[1]))]

    snap = last_snapshot(data)
    sel_comp = st.sidebar.multiselect("Compagnie", sorted(snap[COL_COMPAGNIE].unique()))
    sel_mod = st.sidebar.multiselect("Modèle", sorted(snap[COL_MODELE].unique()))
    sel_stat = st.sidebar.multiselect("Statut", sorted(snap[COL_STATUT].unique()))
    hmin, hmax = int(snap[COL_HEURES].min()), int(snap[COL_HEURES].max())
    if hmin < hmax:
        h_range = st.sidebar.slider("Heures de vol", hmin, hmax, (hmin, hmax), step=50)
    else:
        h_range = (hmin, hmax)

    mask = snap[COL_HEURES].between(*h_range)
    if sel_comp:
        mask &= snap[COL_COMPAGNIE].isin(sel_comp)
    if sel_mod:
        mask &= snap[COL_MODELE].isin(sel_mod)
    if sel_stat:
        mask &= snap[COL_STATUT].isin(sel_stat)
    snap = snap[mask].reset_index(drop=True)
    data = data[data[COL_ID].isin(snap[COL_ID])]
    return data, snap, is_history


def tab_dashboard(snap: pd.DataFrame, thr_h: int, thr_c: int) -> None:
    total_h, total_c = snap[COL_HEURES].sum(), snap[COL_CYCLES].sum()
    n = len(snap)
    n_ops = int(snap[COL_STATUT].str.lower().str.startswith("op").sum())
    top = snap.loc[snap[COL_HEURES].idxmax()]
    over = snap[(snap[COL_HEURES] >= thr_h) | (snap[COL_CYCLES] >= thr_c)]

    c = st.columns(4)
    c[0].markdown(kpi("Aéronefs suivis", fmt(n), f"{snap[COL_MODELE].nunique()} modèles"), unsafe_allow_html=True)
    c[1].markdown(kpi("Heures de vol cumulées", fmt(total_h) + " h", f"moyenne {fmt(total_h / n)} h / avion",
                      "violet"), unsafe_allow_html=True)
    c[2].markdown(kpi("Cycles cumulés", fmt(total_c), f"moyenne {fmt(total_c / n)} / avion", "orange"),
                  unsafe_allow_html=True)
    c[3].markdown(kpi("Durée moyenne par cycle", f"{total_h / total_c:.2f} h" if total_c else "—",
                      "heures de vol ÷ atterrissages"), unsafe_allow_html=True)
    st.write("")
    c = st.columns(4)
    c[0].markdown(kpi("Taux de disponibilité", f"{100 * n_ops / n:.0f} %", f"{n_ops} avion(s) opérationnel(s)",
                      "green"), unsafe_allow_html=True)
    c[1].markdown(kpi("Immobilisés", fmt(n - n_ops), "maintenance, révision ou autre", "orange"),
                  unsafe_allow_html=True)
    c[2].markdown(kpi("Avion le plus utilisé", top[COL_ID], f"{fmt(top[COL_HEURES])} h · {top[COL_MODELE]}"),
                  unsafe_allow_html=True)
    c[3].markdown(kpi("Seuils dépassés", fmt(len(over)),
                      f"≥ {fmt(thr_h)} h ou ≥ {fmt(thr_c)} cycles", "red"), unsafe_allow_html=True)

    st.markdown('<div class="section-title">Vue d\'ensemble de la flotte</div>', unsafe_allow_html=True)
    left, right = st.columns([1.4, 1])
    with left:
        show(chart_hours_by_aircraft(snap))
    with right:
        show(chart_status_donut(snap))
    left, right = st.columns(2)
    with left:
        show(chart_cycles_by_aircraft(snap))
    with right:
        show(chart_scatter(snap))

    st.markdown('<div class="section-title">Aéronefs à surveiller</div>', unsafe_allow_html=True)
    if over.empty:
        st.success("Aucun aéronef n'a dépassé les seuils d'alerte définis dans le menu de gauche.")
    else:
        alert = over.copy()
        alert["Motif"] = np.where(
            (alert[COL_HEURES] >= thr_h) & (alert[COL_CYCLES] >= thr_c), "Heures et cycles",
            np.where(alert[COL_HEURES] >= thr_h, "Seuil d'heures", "Seuil de cycles"))
        st.dataframe(alert[[COL_ID, COL_MODELE, COL_COMPAGNIE, COL_STATUT, COL_HEURES, COL_CYCLES, "Motif"]]
                     .sort_values(COL_HEURES, ascending=False), width="stretch", hide_index=True)


def tab_analysis(hist: pd.DataFrame, snap: pd.DataFrame, simulated: bool) -> None:
    left, right = st.columns(2)
    with left:
        show(chart_ratio(snap))
    with right:
        show(chart_box_model(snap))
    left, right, third = st.columns([1, 1, 0.8])
    with left:
        show(chart_by_company(snap))
    with right:
        show(chart_distribution(snap))
    with third:
        if len(snap) >= 3:
            show(chart_heatmap(snap))
        else:
            st.info("Au moins 3 aéronefs sont nécessaires pour la matrice de corrélation.")

    st.markdown('<div class="section-title">Évolution dans le temps</div>', unsafe_allow_html=True)
    if simulated:
        st.warning("Historique simulé : votre fichier ne contient qu'une situation par avion. "
                   "Les courbes sont une reconstitution plausible, pas des mesures réelles.")
    metric = st.radio("Indicateur", [COL_HEURES, COL_CYCLES], horizontal=True, key="evo_metric")
    show(chart_evolution(hist, metric))


def tab_race(hist: pd.DataFrame, snap: pd.DataFrame, simulated: bool) -> None:
    st.markdown("Course animée des aéronefs classés selon leur cumul. "
                "Les couleurs correspondent aux modèles d'avion.")
    if simulated:
        st.warning("Votre fichier ne contient qu'une situation par avion (pas d'historique). "
                   "L'animation s'appuie donc sur un historique simulé, réglable dans le menu de gauche. "
                   "Pour une course fidèle à la réalité, importez un fichier avec plusieurs dates par avion.")

    c1, c2, c3, c4 = st.columns(4)
    metric = c1.selectbox("Indicateur", [COL_HEURES, COL_CYCLES])
    n_bars = c2.slider("Nombre de barres", 3, max(3, min(15, len(snap))), min(10, len(snap)))
    fps = c3.slider("Vitesse (images/s)", 5, 30, 15)
    steps = c4.slider("Fluidité (images par période)", 2, 15, 8)

    pivot = make_race_pivot(hist, metric)
    if len(pivot) < 2:
        st.error("Au moins deux dates sont nécessaires pour animer la course.")
        return
    steps = max(1, min(steps, 250 // (len(pivot) - 1)))  # plafonne le nombre d'images
    meta = build_color_meta(snap)
    sig = (metric, n_bars, fps, steps, pivot.shape, round(float(pivot.to_numpy().sum()), 2))

    if st.button("Générer l'animation", type="primary"):
        with st.spinner(f"Génération de {(len(pivot) - 1) * steps + 1} images en cours..."):
            gif = render_race(pivot, meta, f"Évolution des {metric.lower()} de la flotte",
                              n_bars, steps, fps)
        st.session_state["race"] = {"sig": sig, "gif": gif}

    race = st.session_state.get("race")
    if race:
        if race["sig"] != sig:
            st.info("Les paramètres ou les filtres ont changé : cliquez sur « Générer l'animation » "
                    "pour mettre à jour (animation précédente affichée ci-dessous).")
        st.image(race["gif"])
        st.download_button("Télécharger l'animation (GIF)", race["gif"],
                           file_name="bar_chart_race_flotte.gif", mime="image/gif")


def tab_data(clean: pd.DataFrame, report: list, snap: pd.DataFrame) -> None:
    st.markdown('<div class="section-title">Rapport de nettoyage</div>', unsafe_allow_html=True)
    st.dataframe(pd.DataFrame(report, columns=["Étape", "Résultat"]), width="stretch", hide_index=True)

    st.markdown('<div class="section-title">Situation actuelle de la flotte (après filtres)</div>',
                unsafe_allow_html=True)
    shown = snap.copy()
    shown[COL_DATE] = shown[COL_DATE].dt.date
    st.dataframe(shown, width="stretch", hide_index=True)

    c1, c2, _ = st.columns([1, 1, 3])
    c1.download_button("Télécharger en CSV", shown.to_csv(index=False).encode("utf-8-sig"),
                       file_name="flotte_nettoyee.csv", mime="text/csv", width="stretch")
    xls = io.BytesIO()
    shown.to_excel(xls, index=False, sheet_name="Flotte")
    c2.download_button("Télécharger en Excel", xls.getvalue(), file_name="flotte_nettoyee.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       width="stretch")


def main() -> None:
    st.set_page_config(page_title="Suivi des heures de vol", layout="wide")
    st.markdown(CSS, unsafe_allow_html=True)
    setup_style()

    if not st.session_state.get("auth"):
        login_page()
        return

    st.sidebar.markdown(f"**Connecté : {st.session_state.get('user', '')}**")
    if st.sidebar.button("Se déconnecter", width="stretch"):
        st.session_state.clear()
        st.rerun()
    st.sidebar.divider()

    st.markdown('<div class="hero"><h1>Suivi des heures de vol et des cycles</h1>'
                '<p>Gestion technique et suivi opérationnel de maintenance de la flotte aérienne</p></div>',
                unsafe_allow_html=True)

    clean, report = sidebar_data()
    hist, snap, is_history = sidebar_filters(clean)
    if snap.empty:
        st.warning("Aucun aéronef ne correspond aux filtres sélectionnés.")
        st.stop()

    st.sidebar.markdown("### Seuils d'alerte")
    thr_h = st.sidebar.number_input("Heures de vol", 0, 200000, 20000, step=500)
    thr_c = st.sidebar.number_input("Cycles", 0, 200000, 15000, step=500)

    simulated = not is_history
    if simulated:
        with st.sidebar.expander("Simulation de l'historique"):
            st.caption("Utilisée pour l'évolution et le Bar Chart Race car le fichier ne contient "
                       "qu'une situation par avion.")
            months = st.slider("Nombre de mois", 6, 36, 12)
            age = st.slider("Ancienneté estimée (mois)", 24, 240, 120,
                            help="Sert à estimer l'utilisation mensuelle moyenne : cumul ÷ ancienneté.")
        hist = simulate_history(snap, months, age)

    tabs = st.tabs(["Tableau de bord", "Analyses", "Bar Chart Race", "Nettoyage et données"])
    with tabs[0]:
        tab_dashboard(snap, thr_h, thr_c)
    with tabs[1]:
        tab_analysis(hist, snap, simulated)
    with tabs[2]:
        tab_race(hist, snap, simulated)
    with tabs[3]:
        tab_data(clean, report, snap)


if __name__ == "__main__":
    main()
