import xmlrpc.client
from datetime import datetime, timedelta, date
import calendar
import re
import streamlit as st
import pandas as pd
import plotly.express as px
from streamlit_autorefresh import st_autorefresh

# ---------- CONFIG ODOO ----------
ODOO_URL = "https://olsen-engineering.odoo.com"
DB = "mynalios-olsen-main-7388485"
USERNAME = "f.mordant@olsen-engineering.com"
PASSWORD = "a9a52b95f9ba02f3d813aa02e113d51ffac6de1d"

# ============================================================
# 🔧 ODOO HELPERS
# ============================================================

def connect_odoo():
    common = xmlrpc.client.ServerProxy(f"{ODOO_URL}/xmlrpc/2/common")
    uid = common.authenticate(DB, USERNAME, PASSWORD, {})
    if not uid:
        raise Exception("Échec d'authentification Odoo")
    models = xmlrpc.client.ServerProxy(f"{ODOO_URL}/xmlrpc/2/object")
    return uid, models


def get_top_companies_batch(uid, models, partner_ids):
    clean_ids = []
    for pid in partner_ids:
        if pid and isinstance(pid, (list, tuple)):
            clean_ids.append(pid[0])
        elif pid and isinstance(pid, int):
            clean_ids.append(pid)

    clean_ids = list(set(clean_ids))
    if not clean_ids:
        return {}

    partners = models.execute_kw(
        DB, uid, PASSWORD,
        "res.partner", "read",
        [clean_ids], {"fields": ["id", "name", "parent_id"]}
    )

    parent_ids = list({p["parent_id"][0] for p in partners if p["parent_id"]})
    parent_map = {}
    if parent_ids:
        parents = models.execute_kw(
            DB, uid, PASSWORD,
            "res.partner", "read",
            [parent_ids], {"fields": ["id", "name"]}
        )
        parent_map = {p["id"]: p["name"] for p in parents}

    result = {}
    for p in partners:
        if p["parent_id"]:
            result[p["id"]] = parent_map.get(p["parent_id"][0], p["name"])
        else:
            result[p["id"]] = p["name"]

    return result


def extract_project_code(display_name: str) -> str:
    if not display_name:
        return ""
    m = re.search(r"S\d{2}-\d{5}", display_name)
    return m.group(0) if m else ""


@st.cache_data(ttl=300)
def load_projects(_uid, _models, filter_mode="both"):
    uid, models = _uid, _models

    tag_engineering = models.execute_kw(
        DB, uid, PASSWORD, 'project.tags', 'search',
        [[('name', '=', 'Engineering')]]
    )
    tag_standard = models.execute_kw(
        DB, uid, PASSWORD, 'project.tags', 'search',
        [[('name', '=', 'Standard')]]
    )
    tag_prolig = models.execute_kw(
        DB, uid, PASSWORD, 'project.tags', 'search',
        [[('name', 'ilike', 'PRO (LIG)')]]
    )

    if filter_mode == "engineering":
        domain = [
            ('stage_id.name', 'not in', ['Cloturé', 'Template', 'Annulé']),
            ('tag_ids', 'in', tag_engineering),
            ('tag_ids', 'in', tag_prolig),
        ]
    elif filter_mode == "standard":
        domain = [
            ('stage_id.name', 'not in', ['Cloturé', 'Template', 'Annulé']),
            ('tag_ids', 'in', tag_standard),
            ('tag_ids', 'in', tag_prolig),
        ]
    else:  # "both"
        domain = [
            ('stage_id.name', 'not in', ['Cloturé', 'Template', 'Annulé']),
            '|',
                ('tag_ids', 'in', tag_engineering),
                ('tag_ids', 'in', tag_standard),
            ('tag_ids', 'in', tag_prolig),
        ]

    fields = ['id', 'display_name', 'partner_id', 'name', 'analytic_account_id']
    projects = models.execute_kw(
        DB, uid, PASSWORD, 'project.project', 'search_read',
        [domain], {'fields': fields}
    )

    partner_ids = [p["partner_id"] for p in projects]
    company_map = get_top_companies_batch(uid, models, partner_ids)
    for p in projects:
        pid = p["partner_id"][0] if p["partner_id"] else None
        p["company"] = company_map.get(pid, "N/A")

    project_ids = [p["id"] for p in projects]
    updates = models.execute_kw(
        DB, uid, PASSWORD,
        'project.update', 'search_read',
        [[('project_id', 'in', project_ids)]],
        {'fields': ['project_id', 'status', 'write_date']}
    )
    last_update = {}
    for u in updates:
        pid = u["project_id"][0]
        if pid not in last_update or u["write_date"] > last_update[pid]["write_date"]:
            last_update[pid] = u

    filtered = [p for p in projects if last_update.get(p["id"], {}).get("status") != "done"]
    filtered.sort(key=lambda p: extract_project_code(p['display_name']))
    return filtered


def get_tasks(uid, models, project_ids, start_date, end_date):
    domain = [
        ('project_id', 'in', project_ids),
        ('date_deadline', '!=', False),
        ('tag_ids.name', 'in', ['Engineering', 'PRO (LIG)', 'PRO(LIG)']),
    ]
    fields = ['id', 'name', 'project_id', 'date_deadline']
    tasks = models.execute_kw(
        DB, uid, PASSWORD, 'project.task', 'search_read',
        [domain], {'fields': fields}
    )
    for t in tasks:
        raw = t['date_deadline']
        if raw:
            raw = raw.split(" ")[0]
            t['date_deadline'] = datetime.strptime(raw, '%Y-%m-%d').date()
    return tasks


# ============================================================
# 🔥 PURCHASE LOADER
# ============================================================

@st.cache_data(ttl=300)
def load_purchase_data_all_projects():
    uid, models = connect_odoo()

    po_data = models.execute_kw(
        DB, uid, PASSWORD,
        "purchase.order", "search_read",
        [[("state", "=", "purchase")]],
        {"fields": ["id", "user_id", "name"]}
    )

    buyer_map = {
        po["id"]: (po["user_id"][1] if po["user_id"] else "Unknown")
        for po in po_data
    }
    po_name_map = {po["id"]: po["name"] for po in po_data}
    po_ids = [po["id"] for po in po_data]

    po_lines = models.execute_kw(
        DB, uid, PASSWORD,
        "purchase.order.line", "search_read",
        [[("order_id", "in", po_ids)]],
        {"fields": [
            "name", "product_qty", "qty_received",
            "date_planned", "order_id", "product_id",
            "analytic_distribution"
        ]}
    )

    product_ids = list({l["product_id"][0] for l in po_lines if l.get("product_id")})
    policy_map = {}
    if product_ids:
        products = models.execute_kw(
            DB, uid, PASSWORD,
            "product.product", "read",
            [product_ids],
            {"fields": ["type"]}
        )
        policy_map = {p["id"]: p["type"] for p in products}

    return po_lines, policy_map, buyer_map, po_name_map


def get_purchase_for_project(project, po_lines, policy_map, buyer_map, po_name_map):
    """
    Couleurs des lignes :
      - orange (#FFA000)  : partiellement reçu                   → rank 0
      - grey   (#757575)  : en retard, produit physique           → rank 1
      - white  (#FFFFFF)  : pas encore dû                         → rank 2
      - blue   (#1565C0)  : en retard, service                    → rank 3 (liste détaillée)
      - green  (#2E7D32)  : totalement reçu                       → rank 4

    Progress bar order : orange | grey | white | blue | green
    Rouge vignette : uniquement si grey > 0 (produits physiques en retard).
    Les lignes service (blue) n'influent PAS sur la couleur rouge.
    """
    today = date.today()
    orange = grey = white = green = blue_service = 0
    formatted = []

    analytic_id = project["analytic_account_id"][0] if project.get("analytic_account_id") else None
    if not analytic_id:
        return {"orange": 0, "grey": 0, "white": 0, "green": 0, "blue": 0, "total": 0}, []

    analytic_str = str(analytic_id)

    for l in po_lines:
        dist = l.get("analytic_distribution") or {}
        if analytic_str not in dist:
            continue

        if l["product_qty"] == 0:
            continue

        qty_o = l["product_qty"]
        qty_r = l["qty_received"]

        if l["date_planned"]:
            d = l["date_planned"].split(" ")[0]
            dp = datetime.strptime(d, "%Y-%m-%d").date()
        else:
            dp = None

        pid_prod = l["product_id"][0] if l.get("product_id") else None
        prod_type = policy_map.get(pid_prod, "")
        is_service = (prod_type == "service")

        if qty_r >= qty_o:
            color = "#2E7D32"; rank = 4; green += 1
        elif qty_r > 0:
            color = "#FFA000"; rank = 0; orange += 1
        elif dp and dp < today:
            if is_service:
                color = "#1565C0"; rank = 3; blue_service += 1
            else:
                color = "#757575"; rank = 1; grey += 1
        else:
            color = "#FFFFFF"; rank = 2; white += 1

        po_id = l["order_id"][0]
        formatted.append({
            "PO": po_name_map.get(po_id, str(po_id)),
            "Buyer": buyer_map.get(po_id, "Unknown"),
            "Description": short_desc(l["name"], 50),
            "Ordered": qty_o,
            "Received": qty_r,
            "Planned Date": dp,
            "Color": color,
            "Rank": rank,
            "IsService": is_service,
        })

    formatted.sort(key=lambda x: x["Rank"])

    summary = {
        "orange": orange,
        "grey": grey,
        "white": white,
        "green": green,
        "blue": blue_service,
        "total": orange + grey + white + green + blue_service
    }
    return summary, formatted


# ============================================================
# 📊 ANALYTICS LOADERS
# ============================================================

@st.cache_data(ttl=300)
def load_projects_with_closed(_uid, _models, filter_mode="both"):
    """
    Comme load_projects mais INCLUT les projets Cloturés.
    Retourne aussi is_closed pour colorisation.
    """
    uid, models = _uid, _models

    tag_engineering = models.execute_kw(
        DB, uid, PASSWORD, 'project.tags', 'search',
        [[('name', '=', 'Engineering')]]
    )
    tag_standard = models.execute_kw(
        DB, uid, PASSWORD, 'project.tags', 'search',
        [[('name', '=', 'Standard')]]
    )
    tag_prolig = models.execute_kw(
        DB, uid, PASSWORD, 'project.tags', 'search',
        [[('name', 'ilike', 'PRO (LIG)')]]
    )

    # Pas de filtre sur le stage — on prend tout sauf Template/Annulé
    base_domain = [('stage_id.name', 'not in', ['Template', 'Annulé'])]

    if filter_mode == "engineering":
        domain = base_domain + [
            ('tag_ids', 'in', tag_engineering),
            ('tag_ids', 'in', tag_prolig),
        ]
    elif filter_mode == "standard":
        domain = base_domain + [
            ('tag_ids', 'in', tag_standard),
            ('tag_ids', 'in', tag_prolig),
        ]
    else:
        domain = base_domain + [
            '|',
                ('tag_ids', 'in', tag_engineering),
                ('tag_ids', 'in', tag_standard),
            ('tag_ids', 'in', tag_prolig),
        ]

    fields = ['id', 'display_name', 'partner_id', 'name', 'analytic_account_id', 'stage_id']
    projects = models.execute_kw(
        DB, uid, PASSWORD, 'project.project', 'search_read',
        [domain], {'fields': fields}
    )

    partner_ids = [p["partner_id"] for p in projects]
    company_map = get_top_companies_batch(uid, models, partner_ids)
    for p in projects:
        pid = p["partner_id"][0] if p["partner_id"] else None
        p["company"] = company_map.get(pid, "N/A")
        stage_name = p["stage_id"][1] if p.get("stage_id") else ""
        p["is_closed"] = (stage_name == "Cloturé")

    # Tri par code décroissant
    projects.sort(key=lambda p: extract_project_code(p['display_name']), reverse=True)
    return projects


@st.cache_data(ttl=300)
def load_analytics_for_projects(_uid, _models, project_list):
    """
    Bilan analytique année en cours uniquement.
      - Dépenses   : montants négatifs sur account.analytic.line (année courante)
      - Facturé    : montants positifs sur account.analytic.line (année courante)
      - CA total   : montant total des SO confirmés liés au projet (tous les temps)
      - À facturer : CA total - Facturé (cumulé toutes années)
      - Marge C    : CA total - Dépenses (année courante)
    """
    uid, models = _uid, _models

    analytic_ids = [
        p["analytic_account_id"][0]
        for p in project_list
        if p.get("analytic_account_id")
    ]
    if not analytic_ids:
        return {}

    year_start = f"{date.today().year}-01-01"
    year_end   = f"{date.today().year}-12-31"

    # ── 1. Lignes analytiques — année en cours ────────────────────────────────
    analytic_lines = models.execute_kw(
        DB, uid, PASSWORD,
        "account.analytic.line", "search_read",
        [[
            ("account_id", "in", analytic_ids),
            ("date", ">=", year_start),
            ("date", "<=", year_end),
        ]],
        {"fields": ["account_id", "amount"]}
    )

    depenses_map = {}
    facture_map  = {}
    for line in analytic_lines:
        aid = line["account_id"][0]
        amt = line["amount"]
        if amt < 0:
            depenses_map[aid] = depenses_map.get(aid, 0.0) + abs(amt)
        elif amt > 0:
            facture_map[aid]  = facture_map.get(aid, 0.0) + amt

        # ── NOUVEAU : Dépenses + Facturé toutes années ───────────────────────────────
    all_lines = models.execute_kw(
        DB, uid, PASSWORD,
        "account.analytic.line", "search_read",
        [[("account_id", "in", analytic_ids)]],
        {"fields": ["account_id", "amount"]}
    )
    
    depenses_all = {}
    facture_all  = {}
    
    for line in all_lines:
        aid = line["account_id"][0]
        amt = line["amount"]
        if amt < 0:
            depenses_all[aid] = depenses_all.get(aid, 0.0) + abs(amt)
        elif amt > 0:
            facture_all[aid]  = facture_all.get(aid, 0.0) + amt

    
    # ── 2. Facturé cumulé toutes années (pour calculer reste à facturer) ──────
    all_analytic_lines = models.execute_kw(
        DB, uid, PASSWORD,
        "account.analytic.line", "search_read",
        [[("account_id", "in", analytic_ids)]],
        {"fields": ["account_id", "amount"]}
    )
    facture_total_map = {}
    for line in all_analytic_lines:
        aid = line["account_id"][0]
        amt = line["amount"]
        if amt > 0:
            facture_total_map[aid] = facture_total_map.get(aid, 0.0) + amt

    # ── 3. CA total via sale.order (matching code projet) ─────────────────────
    code_to_proj = {}
    for p in project_list:
        code = extract_project_code(p.get("display_name", ""))
        if code:
            code_to_proj[code] = p

    ca_map = {}
    if code_to_proj:
        all_so = models.execute_kw(
            DB, uid, PASSWORD,
            "sale.order", "search_read",
            [[("state", "in", ["sale", "done"])]],
            {"fields": ["id", "name", "amount_untaxed"]}
        )
        for so in all_so:
            so_code = extract_project_code(so["name"])
            if not so_code:
                continue
            proj = code_to_proj.get(so_code)
            if not proj or not proj.get("analytic_account_id"):
                continue
            aid = proj["analytic_account_id"][0]
            ca_map[aid] = ca_map.get(aid, 0.0) + so["amount_untaxed"]

    # ── 4. Résultat par projet ────────────────────────────────────────────────
    result = {}
    for p in project_list:
        if not p.get("analytic_account_id"):
            result[p["id"]] = None
            continue
        aid = p["analytic_account_id"][0]

        ca_total       = ca_map.get(aid, 0.0)
        depenses       = depenses_map.get(aid, 0.0)        # année en cours
        facture_annee  = facture_map.get(aid, 0.0)         # année en cours
        facture_cumul  = facture_total_map.get(aid, 0.0)   # tout temps
        a_facturer     = max(ca_total - facture_cumul, 0.0)
        marge_c        = ca_total - depenses                # CA total - dépenses année
        marge_pct      = (marge_c / ca_total * 100) if ca_total > 0 else 0.0

        result[p["id"]] = {
            # Résumé (année en cours)
            "depenses_annee": depenses,
            "facture_annee":  facture_annee,
        
            # Tableau (toutes années)
            "depenses_all":   depenses_all.get(aid, 0.0),
            "facture_all":    facture_all.get(aid, 0.0),
        
            # Données globales
            "ca_total":       ca_total,
            "a_facturer":     max(ca_total - facture_all.get(aid, 0.0), 0.0),
            "marge_c":        ca_total - depenses_all.get(aid, 0.0),
            "marge_pct":      ((ca_total - depenses_all.get(aid, 0.0)) / ca_total * 100) if ca_total > 0 else 0.0,
        
            "is_closed":      p.get("is_closed", False),
        }

    return result


@st.cache_data(ttl=300)
def load_analytics_monthly(_uid, _models, project_list):
    """
    Charge les lignes analytiques des 12 derniers mois pour tous les projets
    de la liste, agrégées par mois.
    Retourne un DataFrame avec colonnes : Mois (date), CA, Dépenses
    """
    uid, models = _uid, _models

    analytic_ids = [
        p["analytic_account_id"][0]
        for p in project_list
        if p.get("analytic_account_id")
    ]
    if not analytic_ids:
        return pd.DataFrame()

    date_start = (date.today().replace(day=1) - timedelta(days=365)).strftime("%Y-%m-%d")

    lines = models.execute_kw(
        DB, uid, PASSWORD,
        "account.analytic.line", "search_read",
        [[
            ("account_id", "in", analytic_ids),
            ("date", ">=", date_start),
        ]],
        {"fields": ["account_id", "amount", "date"]}
    )

    if not lines:
        return pd.DataFrame()

    records = []
    for l in lines:
        d = datetime.strptime(l["date"], "%Y-%m-%d").date()
        mois = d.replace(day=1)  # 1er du mois
        amt = l["amount"]
        records.append({
            "Mois": mois,
            "CA":       amt if amt > 0 else 0.0,
            "Dépenses": abs(amt) if amt < 0 else 0.0,
        })

    df = pd.DataFrame(records)
    df_agg = df.groupby("Mois")[["CA", "Dépenses"]].sum().reset_index()
    df_agg = df_agg.sort_values("Mois")

    # Remplir les mois manquants avec 0
    all_months = pd.date_range(
        start=date_start, end=date.today().strftime("%Y-%m-%d"), freq="MS"
    ).date
    df_full = pd.DataFrame({"Mois": all_months})
    df_full = df_full.merge(df_agg, on="Mois", how="left").fillna(0)

    return df_full


# ============================================================
# 🔧 GANTT HELPERS
# ============================================================

def build_weeks_horizon(months=3):
    start = date.today()
    end = start + timedelta(days=30 * months)
    current = start - timedelta(days=start.weekday())
    weeks = []
    while current <= end:
        weeks.append((current.isocalendar()[1], current, current + timedelta(days=6)))
        current += timedelta(days=7)
    return weeks


COLOR_ORDER = [
    "Soudure", "Peinture", "Assemblage", "Câblage",
    "Test", "Montage", "Mise en service", "Réception",
    "Transport", "Etude", "Autres"
]

COLOR_MAP = {
    "Soudure":        "#1E88E5",
    "Peinture":       "#FDD835",
    "Assemblage":     "#43A047",
    "Câblage":        "#8E24AA",
    "Test":           "#FB8C00",
    "Montage":        "#E53935",
    "Mise en service":"#EC407A",
    "Réception":      "#6D4C41",
    "Transport":      "#00ACC1",
    "Etude":          "#34ebc6",
    "Autres":         "#9E9E9E"
}


def classify_task_type(name):
    n = name.lower()
    if "soud" in n: return "Soudure"
    if "peint" in n: return "Peinture"
    if "assembl" in n: return "Assemblage"
    if "cabl" in n or "câbl" in n: return "Câblage"
    if "test" in n: return "Test"
    if "montage" in n or "installation" in n: return "Montage"
    if "mise en service" in n or "mes" in n: return "Mise en service"
    if "récept" in n or "recept" in n or "assistance" in n: return "Réception"
    if "transport" in n: return "Transport"
    if "étude" in n or "etude" in n or "conception" in n or "plan" in n or "calcul" in n: return "Etude"
    return "Autres"


def classify_task_color(name):
    return COLOR_MAP[classify_task_type(name)]


def map_tasks_to_grid(projects, tasks, weeks):
    proj_index = {p['id']: i for i, p in enumerate(projects)}
    grid = {}
    detailed = {}
    for t in tasks:
        pid = t['project_id'][0]
        if pid not in proj_index:
            continue
        row = proj_index[pid]
        color = classify_task_color(t['name'])
        for col, (_, start_w, end_w) in enumerate(weeks):
            if start_w <= t['date_deadline'] <= end_w:
                key = (row, col)
                grid.setdefault(key, []).append(color)
                detailed.setdefault(key, []).append(t)
                break
    return grid, detailed


def clean_description_from_display_name(display_name: str) -> str:
    if not display_name:
        return ""
    if " - " not in display_name:
        return display_name
    parts = display_name.split(" - ")
    if len(parts) >= 2 and parts[-1].strip() == parts[-2].strip():
        parts = parts[:-1]
    return " - ".join(parts)


def short_desc(desc: str, max_len: int) -> str:
    if not desc:
        return ""
    if len(desc) <= max_len:
        return desc
    return desc[:max_len].rstrip() + "…"


def project_label(p):
    client = p.get("company", "N/A")
    display = p.get("display_name") or p.get("name") or "Projet"
    desc_clean = clean_description_from_display_name(display)
    desc_short = short_desc(desc_clean, 20)
    return f"{client} - {desc_short}"


def fmt_eur(val):
    return f"{val:,.0f} €".replace(",", " ")


# ============================================================
# 🔵 STREAMLIT APP
# ============================================================

def main():
    st.set_page_config(page_title="Master Planning Odoo", layout="wide")

    if "filter_standard" not in st.session_state:
        st.session_state["filter_standard"] = False

    st.markdown("""
    <style>
    .block-container { padding-top: 0.5rem !important; }
    div[data-testid="stToggle"] > label { font-size: 13px !important; }
    </style>
    """, unsafe_allow_html=True)

    try:
        uid, models = connect_odoo()
    except Exception as e:
        st.error(f"Connexion Odoo impossible : {e}")
        return

    st_autorefresh(interval=600000, key="refresh_10min")

    if "months" not in st.session_state:
        st.session_state["months"] = 3
    if "selected_purchase_project_id" not in st.session_state:
        st.session_state["selected_purchase_project_id"] = None
    if "filter_engineering" not in st.session_state:
        st.session_state["filter_engineering"] = True
    if "filter_standard" not in st.session_state:
        st.session_state["filter_standard"] = True

    # ---------- BANNIÈRE ----------
    col1, col2, col3 = st.columns([1, 4, 1])
    with col1:
        st.image("https://upload.wikimedia.org/wikipedia/commons/b/ba/Olsen-Logo.png", width=180)
        st.markdown(
            "<div style='text-align:left;color:green;font-weight:bold;margin-top:20px;'>🟢 Connecté à Odoo</div>",
            unsafe_allow_html=True
        )
    with col2:
        st.markdown(
            "<h2 style='text-align:center;margin-top:10px;'>Olsen Dashboard</h2>",
            unsafe_allow_html=True
        )
    with col3:
        filter_engineering = st.toggle(
            "🔵 Engineering (PRO LIG)",
            value=st.session_state["filter_engineering"],
            key="toggle_engineering"
        )
        filter_standard = st.toggle(
            "⚪ Standard (PRO LIG)",
            value=st.session_state["filter_standard"],
            key="toggle_standard"
        )

        if not filter_engineering and not filter_standard:
            st.warning("⚠️ Au moins un filtre doit être actif.")
            filter_engineering = True

        if filter_engineering and filter_standard:
            filter_mode = "both"
        elif filter_engineering:
            filter_mode = "engineering"
        else:
            filter_mode = "standard"

        if (filter_engineering != st.session_state["filter_engineering"] or
                filter_standard != st.session_state["filter_standard"]):
            st.session_state["filter_engineering"] = filter_engineering
            st.session_state["filter_standard"] = filter_standard
            st.rerun()

    tab1, tab2, tab3 = st.tabs(["📅 Planning", "📦 Purchases", "📊 Analytique"])

    # ============================================================
    # 🟦 ONGLET 1 — MASTER PLANNING
    # ============================================================
    with tab1:
        projects = load_projects(uid, models, filter_mode)
        project_ids = [p['id'] for p in projects]

        months = st.session_state["months"]
        weeks = build_weeks_horizon(months)
        tasks = get_tasks(uid, models, project_ids, weeks[0][1], weeks[-1][2])
        grid, detailed = map_tasks_to_grid(projects, tasks, weeks)

        st.subheader("📊 Gantt")

        gantt_data = []
        overlap_counter = {}

        for t in tasks:
            proj = next((p for p in projects if p['id'] == t['project_id'][0]), None)
            if not proj:
                continue
            label = project_label(proj)
            deadline = t["date_deadline"]

            week_key = (label, deadline.isocalendar()[1], deadline.year)
            count = overlap_counter.get(week_key, 0)
            overlap_counter[week_key] = count + 1
            offset_days = count * 1

            gantt_data.append({
                "Tâche": t["name"],
                "Projet": label,
                "Début": deadline - timedelta(days=3) + timedelta(days=offset_days),
                "Fin": deadline + timedelta(days=3) + timedelta(days=offset_days),
                "Type": classify_task_type(t["name"])
            })

        if gantt_data:
            df_gantt = pd.DataFrame(gantt_data)
            df_gantt["code"] = df_gantt["Projet"].apply(extract_project_code)
            df_gantt = df_gantt.sort_values(by="code")
            df_gantt["Type détaillé"] = pd.Categorical(
                df_gantt["Type"],
                categories=COLOR_ORDER,
                ordered=True
            )

            fig = px.timeline(
                df_gantt,
                x_start="Début",
                x_end="Fin",
                y="Projet",
                color="Type détaillé",
                color_discrete_map=COLOR_MAP,
                hover_name="Tâche",
                hover_data={"Début": True, "Fin": True, "Type détaillé": True, "Projet": False}
            )

            fig.update_yaxes(autorange="reversed")

            n_proj = len(df_gantt["Projet"].unique())
            chart_height = max(500, n_proj * 18 + 140)

            today = date.today()
            start_view = today
            end_view = today + timedelta(days=30 * months)

            fig.update_layout(
                dragmode="pan",
                height=chart_height,
                bargap=0.3,
                bargroupgap=0.1,
                margin=dict(l=20, r=20, t=40, b=20),
                yaxis=dict(tickfont=dict(size=12), showgrid=True,
                           gridcolor="rgba(180,180,180,0.18)", gridwidth=1),
                xaxis=dict(showgrid=False),
                plot_bgcolor="rgba(0,0,0,0)",
                legend=dict(orientation="h", yanchor="bottom", y=1.02,
                            xanchor="center", x=0.5, font=dict(size=10))
            )

            fig.add_vline(x=today, line_width=2, line_color="white", opacity=0.9)

            cur = date(today.year, today.month, 1)
            while True:
                cur = date(cur.year + 1, 1, 1) if cur.month == 12 else date(cur.year, cur.month + 1, 1)
                if cur > end_view:
                    break
                fig.add_vline(x=cur, line_width=1, line_dash="dot",
                              line_color="rgba(200,200,200,0.35)")

            fig.update_xaxes(range=[start_view, end_view])

            st.plotly_chart(fig, use_container_width=True, config={"displaylogo": False})
        else:
            st.info("Aucune tâche à afficher dans le Gantt.")

        col_s1, col_s2 = st.columns([3, 1])
        with col_s1:
            new_months = st.slider("", 1, 6, months)
        with col_s2:
            st.markdown(
                f"<div style='margin-top:10px;font-size:14px;'>Projets : <b>{len(projects)}</b></div>",
                unsafe_allow_html=True
            )

        if new_months != months:
            st.session_state["months"] = new_months
            st.rerun()

        st.subheader("🔍 Tâches du projet")
        project_labels = {project_label(p): p["id"] for p in projects}
        selected_label = st.selectbox(
            "Sélectionne un projet pour afficher ses tâches",
            ["— Aucun —"] + list(project_labels.keys()),
            index=0
        )

        if selected_label != "— Aucun —":
            proj_id = project_labels[selected_label]
            tasks_for_project = sorted(
                [t for t in tasks if t["project_id"][0] == proj_id],
                key=lambda x: x["date_deadline"]
            )
            if tasks_for_project:
                for t in tasks_for_project:
                    date_str = t["date_deadline"].strftime("%d-%m-%Y")
                    st.write(f"- **{t['name']}** — deadline : {date_str}")
            else:
                st.info("Aucune tâche pour ce projet.")
        else:
            st.info("Sélectionne un projet pour afficher ses tâches.")

    # ============================================================
    # 🟩 ONGLET 2 — PURCHASE TRACKING
    # ============================================================
    with tab2:
        st.markdown("### 📦 Purchases par projet")

        projects_all = load_projects(uid, models, filter_mode)
        po_lines, policy_map, buyer_map, po_name_map = load_purchase_data_all_projects()

        purchase_data = {}
        for p in projects_all:
            summary, lines = get_purchase_for_project(
                p, po_lines, policy_map, buyer_map, po_name_map
            )
            purchase_data[p['id']] = {"summary": summary, "lines": lines}

        cols_per_row = 6
        for i in range(0, len(projects_all), cols_per_row):
            cols = st.columns(cols_per_row)
            for col, p in zip(cols, projects_all[i:i + cols_per_row]):
                with col:
                    client = p["company"]
                    desc_clean = clean_description_from_display_name(p['display_name'])
                    desc_short_25 = short_desc(desc_clean, 25)

                    summary = purchase_data[p['id']]["summary"]
                    total = summary["total"]
                    total_safe = max(total, 1)

                    # Progress bar : orange | grey | white | blue | green
                    pct_orange = 100 * summary["orange"] / total_safe
                    pct_grey   = 100 * summary["grey"]   / total_safe
                    pct_white  = 100 * summary["white"]  / total_safe
                    pct_blue   = 100 * summary["blue"]   / total_safe
                    pct_green  = 100 * summary["green"]  / total_safe

                    # Couleur du compteur : rouge si produit physique en retard (grey > 0)
                    # Les services bleus n'influent PAS sur le rouge
                    if summary["grey"] > 0:
                        text_color = "red"
                    elif summary["orange"] > 0:
                        text_color = "#FFA000"
                    else:
                        text_color = "white"

                    if st.button(
                        f"{client}\n {desc_short_25}",
                        key=f"proj_btn_{p['id']}"
                    ):
                        st.session_state["selected_purchase_project_id"] = p['id']

                    st.markdown(
                        f"""
                        <div style="
                            width:100%;height:12px;border-radius:6px;
                            overflow:hidden;display:flex;margin-top:4px;border:1px solid #444;
                        ">
                            <div style="width:{pct_orange}%;background:#FFA000;"></div>
                            <div style="width:{pct_grey}%;background:#757575;"></div>
                            <div style="width:{pct_white}%;background:#FFFFFF;"></div>
                            <div style="width:{pct_blue}%;background:#1565C0;"></div>
                            <div style="width:{pct_green}%;background:#2E7D32;"></div>
                        </div>
                        <div style="text-align:right;font-size:12px;color:{text_color};margin-top:2px;">
                            {summary["green"]} / {total} lignes
                        </div>
                        """,
                        unsafe_allow_html=True
                    )

        st.markdown("---")
        st.subheader("📋 Détail des lignes d'achat du projet sélectionné")

        selected_purchase_project_id = st.session_state.get("selected_purchase_project_id", None)
        if selected_purchase_project_id is None:
            st.info("Clique sur une vignette projet pour voir le détail des lignes d'achat.")
        else:
            p = next((p for p in projects_all if p['id'] == selected_purchase_project_id), None)
            if p is None:
                st.warning("Projet introuvable.")
            else:
                st.markdown(
                    f"<div style='font-size:15px;'><b>Projet sélectionné :</b> "
                    f"{p['company']} - {p.get('name') or p['display_name']}</div>",
                    unsafe_allow_html=True
                )

                lines = purchase_data[p['id']]["lines"]

                if not lines:
                    st.info("Aucune ligne d'achat trouvée pour ce projet.")
                else:
                    st.markdown(f"**Total lignes : {len(lines)}**")
                    for row in lines:
                        date_display = row['Planned Date'].strftime("%d-%m-%Y") if row['Planned Date'] else "—"
                        text_color_line = "white" if row['Color'] in ("#1565C0", "#2E7D32", "#757575") else "black"
                        st.markdown(
                            f"""
                            <div style="
                                background:{row['Color']};padding:8px 12px;border-radius:4px;
                                margin-bottom:5px;border:1px solid #555;font-size:14px;
                                color:{text_color_line};display:grid;
                                grid-template-columns: 90px 190px 1fr 80px 90px 110px;
                                column-gap:12px;text-align:left;align-items:center;line-height:1.3;
                            ">
                                <div style="white-space:nowrap;"><b>PO:</b> {row['PO']}</div>
                                <div style="white-space:nowrap;"><b>Buyer:</b> {row['Buyer']}</div>
                                <div><b>Description:</b> {row['Description']}</div>
                                <div style="white-space:nowrap;"><b>Ord.:</b> {row['Ordered']}</div>
                                <div style="white-space:nowrap;"><b>Reçu:</b> {row['Received']}</div>
                                <div style="white-space:nowrap;"><b>Date:</b> {date_display}</div>
                            </div>
                            """,
                            unsafe_allow_html=True
                        )

    # ============================================================
    # 🟨 ONGLET 3 — ANALYTIQUE
    # ============================================================
    with tab3:
        st.markdown(
            f"### 📊 Bilan analytique par projet "
            f"<span style='font-size:14px;color:#aaa;'>— Dépenses & Facturé : année {date.today().year}</span>",
            unsafe_allow_html=True
        )

        # Projets avec Cloturés inclus, triés par code décroissant
        projects_ana = load_projects_with_closed(uid, models, filter_mode)
        # ❌ Exclure les projets dont le compte analytique est "PROJETS (LIG)"
        projects_ana = [
            p for p in projects_ana
            if not (
                p.get("analytic_account_id")
                and p["analytic_account_id"][1] in ["Dépannage (Liège)", "Projets (LIG)"]
                ):
        ]
        with st.spinner("Chargement des données analytiques…"):
            analytics    = load_analytics_for_projects(uid, models, projects_ana)
            df_monthly   = load_analytics_monthly(uid, models, projects_ana)

        if not analytics:
            st.info("Aucune donnée analytique disponible.")
        else:
            rows = []
            for p in projects_ana:
                ana = analytics.get(p["id"])
                if ana is None:
                    continue
                code = extract_project_code(p["display_name"])
                rows.append({
                    "_closed":        p.get("is_closed", False),
                    "Code":           code or p.get("name", "—"),
                    "Client":         p["company"],
                    "Projet":         short_desc(clean_description_from_display_name(p["display_name"]), 40),
                    "CA total (€)":   ana["ca_total"],
                    "Dépenses (€)":   ana["depenses_all"],
                    "Facturé (€)":    ana["facture_all"],
                    "À facturer (€)": ana["a_facturer"],
                    "Marge C (€)":    ana["marge_c"],
                    "Marge C (%)":    ana["marge_pct"],
                })

            if not rows:
                st.info("Aucune ligne analytique trouvée pour ces projets.")
            else:
                df_ana = pd.DataFrame(rows)

                # ── Métriques résumé (projets actifs uniquement pour le bilan) ──
                df_actifs = df_ana[~df_ana["_closed"]]
                total_ca      = df_actifs["CA total (€)"].sum()
                total_dep     = df_actifs["Dépenses (€)"].sum()
                total_afac    = df_actifs["À facturer (€)"].sum()
                total_marge   = df_actifs["Marge C (€)"].sum()
                total_marge_p = (total_marge / total_ca * 100) if total_ca > 0 else 0.0

                m1, m2, m3, m4 = st.columns(4)
                m1.metric("CA Total (actifs)",    fmt_eur(total_ca))
                m2.metric(f"Dépenses {date.today().year}", fmt_eur(total_dep))
                m3.metric("À facturer",           fmt_eur(total_afac))
                m4.metric(
                    "Marge C globale",
                    fmt_eur(total_marge),
                    delta=f"{total_marge_p:.1f} %"
                )

                st.markdown("---")

                # ── Tableau : rendu HTML pour coloriser les lignes fermées ──────
                st.markdown("#### Détail par projet")
                search = st.text_input("🔎 Recherche", "", placeholder="Code, client ou projet...")

                if search:
                    s = search.lower()
                    df_ana = df_ana[
                        df_ana["Code"].str.lower().str.contains(s)
                        | df_ana["Client"].str.lower().str.contains(s)
                        | df_ana["Projet"].str.lower().str.contains(s)
                    ]
                # En-tête
                cols_def = "90px 130px 1fr 100px 110px 100px 110px 100px 80px"
                header_html = f"""
                <div style="display:grid;grid-template-columns:{cols_def};
                    column-gap:10px;padding:6px 10px;font-weight:bold;
                    font-size:12px;color:#aaa;border-bottom:1px solid #444;
                    margin-bottom:4px;">
                    <div>Code</div><div>Client</div><div>Projet</div>
                    <div style="text-align:right;">CA Total</div>
                    <div style="text-align:right;">Dépenses</div>
                    <div style="text-align:right;">Facturé</div>
                    <div style="text-align:right;">À facturer</div>
                    <div style="text-align:right;">Marge C (€)</div>
                    <div style="text-align:right;">Marge C (%)</div>
                </div>"""
                st.markdown(header_html, unsafe_allow_html=True)

                for _, row in df_ana.iterrows():
                    is_closed = row["_closed"]
                    bg    = "#0d2a4a" if is_closed else "rgba(255,255,255,0.03)"
                    border = "1px solid #1565C0" if is_closed else "1px solid #333"
                    marge_color = "#e53935" if row["Marge C (€)"] < 0 else (
                        "#43a047" if row["Marge C (%)"] >= 20 else "#FB8C00"
                    )
                    closed_badge = " <span style='font-size:10px;background:#1565C0;color:white;padding:1px 5px;border-radius:3px;'>Cloturé</span>" if is_closed else ""

                    def fe(v): return f"{v:,.0f} €".replace(",", " ")
                    def fp(v): return f"{v:.1f} %"

                    row_html = f"""
                    <div style="display:grid;grid-template-columns:{cols_def};
                        column-gap:10px;padding:7px 10px;font-size:13px;
                        background:{bg};border:{border};border-radius:4px;margin-bottom:3px;
                        align-items:center;">
                        <div style="white-space:nowrap;font-weight:600;">{row['Code']}</div>
                        <div style="white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">{row['Client']}</div>
                        <div style="overflow:hidden;text-overflow:ellipsis;">{row['Projet']}{closed_badge}</div>
                        <div style="text-align:right;">{fe(row['CA total (€)'])}</div>
                        <div style="text-align:right;">{fe(row['Dépenses (€)'])}</div>
                        <div style="text-align:right;">{fe(row['Facturé (€)'])}</div>
                        <div style="text-align:right;color:#00ACC1;font-weight:600;">{fe(row['À facturer (€)'])}</div>
                        <div style="text-align:right;color:{marge_color};font-weight:600;">{fe(row['Marge C (€)'])}</div>
                        <div style="text-align:right;color:{marge_color};">{fp(row['Marge C (%)'])}</div>
                    </div>"""
                    st.markdown(row_html, unsafe_allow_html=True)

                st.markdown("---")

                # ── Graphique CA & Dépenses sur 12 mois ──────────────────────
                st.markdown("#### Évolution CA & Dépenses — 12 derniers mois")

                if df_monthly.empty:
                    st.info("Pas de données mensuelles disponibles.")
                else:
                    df_m = df_monthly.copy()
                    df_m["Mois_label"] = pd.to_datetime(df_m["Mois"]).dt.strftime("%b %Y")

                    # Deux séries : CA et Dépenses
                    df_ca  = df_m[["Mois", "Mois_label", "CA"]].rename(columns={"CA": "Montant"})
                    df_ca["Série"] = "CA facturé"
                    df_dep = df_m[["Mois", "Mois_label", "Dépenses"]].rename(columns={"Dépenses": "Montant"})
                    df_dep["Série"] = "Dépenses"
                    df_plot = pd.concat([df_ca, df_dep])

                    fig_trend = px.bar(
                        df_plot,
                        x="Mois_label",
                        y="Montant",
                        color="Série",
                        barmode="group",
                        color_discrete_map={
                            "CA facturé": "#43a047",
                            "Dépenses":   "#e53935",
                        },
                        height=380,
                        labels={"Mois_label": "", "Montant": "€"},
                    )
                    fig_trend.update_layout(
                        margin=dict(l=10, r=10, t=20, b=20),
                        plot_bgcolor="rgba(0,0,0,0)",
                        xaxis=dict(tickfont=dict(size=11), showgrid=False),
                        yaxis=dict(
                            tickfont=dict(size=11),
                            showgrid=True,
                            gridcolor="rgba(180,180,180,0.12)",
                            tickformat=",.0f",
                        ),
                        legend=dict(orientation="h", y=1.05, x=0),
                        bargap=0.2,
                        bargroupgap=0.05,
                    )
                    st.plotly_chart(fig_trend, use_container_width=True,
                                    config={"displaylogo": False})

    # ---------- FOOTER ----------
    st.markdown("""
    <style>
    .footer {
        position: fixed; left: 0; bottom: 0; width: 100%;
        background-color: rgba(240,240,240,0.85); color: #333;
        text-align: center; padding: 6px 0; font-size: 14px;
        border-top: 1px solid #ccc; z-index: 9999;
    }
    </style>
    <div class="footer">C Flow - Powered by Olsen-Engineering</div>
    """, unsafe_allow_html=True)


if __name__ == "__main__":
    main()
