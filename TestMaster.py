import xmlrpc.client
from datetime import datetime, timedelta, date
import calendar
import re
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
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
    else:
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
    # Récupère aussi date_assign (date de début) et stage pour savoir si terminée
    fields = ['id', 'name', 'project_id', 'date_deadline', 'stage_id', 'date_last_stage_update']
    tasks = models.execute_kw(
        DB, uid, PASSWORD, 'project.task', 'search_read',
        [domain], {'fields': fields}
    )

    # Récupère les noms des étapes "done" pour ce modèle
    done_stage_ids = models.execute_kw(
        DB, uid, PASSWORD, 'project.task.type', 'search',
        [[('name', 'ilike', 'done')]],
    )
    # Fallback : on cherche aussi "terminé", "fini", "closed"
    done_stage_ids += models.execute_kw(
        DB, uid, PASSWORD, 'project.task.type', 'search',
        [[('name', 'in', ['Terminé', 'Terminée', 'Fini', 'Finie', 'Done', 'Closed', 'Cloturé'])]],
    )
    done_stage_set = set(done_stage_ids)

    for t in tasks:
        raw = t['date_deadline']
        if raw:
            raw = raw.split(" ")[0]
            t['date_deadline'] = datetime.strptime(raw, '%Y-%m-%d').date()
        stage = t.get('stage_id')
        t['is_done'] = bool(stage and stage[0] in done_stage_set)
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
            if is_service:
                color = "#1565C0"; rank = 3; blue_service += 1
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
# 📊 ANALYTICS LOADERS — VERSION OPTIMISÉE (1 appel analytique)
# ============================================================

@st.cache_data(ttl=300)
def load_projects_with_closed(_uid, _models, filter_mode="both"):
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

    projects.sort(key=lambda p: (
        1 if p["is_closed"] else 0,
        tuple(-ord(c) for c in extract_project_code(p['display_name']))
    ))
    return projects


@st.cache_data(ttl=300)
def load_all_analytics(_uid, _models, project_list):
    """
    UN SEUL appel account.analytic.line + UN SEUL appel sale.order.
    Retourne :
      - analytics_summary  : dict[project_id] → métriques (bilan + année)
      - df_monthly         : DataFrame mensuel 12 mois
    
    Fix notes de crédit :
      - Classe 6 (charges) : montant négatif = dépense réelle, positif = avoir fournisseur (réduit dépenses)
      - Classe 7 (produits) : montant positif = revenu réel, négatif = avoir client (réduit revenus)
      → On prend le montant TEL QUEL avec son signe, c'est Odoo qui gère déjà les NC en négatif.
      → On NE classe PAS les comptes 3/4 (acomptes) : ils sont transitoires et ne reflètent
        pas encore le CA ou la dépense définitive. Les ignorer évite les doubles comptes.
    """
    uid, models = _uid, _models

    analytic_ids = [
        p["analytic_account_id"][0]
        for p in project_list
        if p.get("analytic_account_id")
    ]
    if not analytic_ids:
        return {}, pd.DataFrame()

    year_now   = date.today().year
    year_start = f"{year_now}-01-01"
    year_end   = f"{year_now}-12-31"
    date_12m   = (date.today().replace(day=1) - timedelta(days=365)).strftime("%Y-%m-%d")

    # ── 1) UN SEUL appel analytique (sans filtre date pour le bilan complet) ──
    all_lines = models.execute_kw(
        DB, uid, PASSWORD,
        "account.analytic.line", "search_read",
        [[("account_id", "in", analytic_ids)]],
        {"fields": ["account_id", "amount", "general_account_id", "date"], "limit": 0}
    )

    # ── 2) Chargement batch des comptes financiers ──
    account_ids_set = {
        l["general_account_id"][0]
        for l in all_lines
        if l.get("general_account_id")
    }
    account_code_map = {}
    account_ids_list = list(account_ids_set)
    batch_size = 200
    for i in range(0, len(account_ids_list), batch_size):
        sub = account_ids_list[i:i + batch_size]
        accounts = models.execute_kw(
            DB, uid, PASSWORD,
            "account.account", "read",
            [sub],
            {"fields": ["id", "code"]}
        )
        for a in accounts:
            account_code_map[a["id"]] = a["code"]

    # ── 3) Classification & agrégation ──
    # Par projet : all-time
    depenses_map  = {}   # charges (classe 6) — montant négatif en Odoo
    revenus_map   = {}   # produits (classe 7) — montant positif en Odoo

    # Par projet : année courante
    depenses_annee_map = {}
    revenus_annee_map  = {}

    # Mensuel (12 derniers mois) : toutes classes 6 + 7
    monthly_records = []

    for line in all_lines:
        if not line.get("account_id") or not line.get("general_account_id"):
            continue

        aid  = line["account_id"][0]
        amt  = line["amount"]
        code = account_code_map.get(line["general_account_id"][0], "")
        d    = line.get("date", "")

        # ── Classification stricte : 6 = charge, 7 = produit, reste ignoré ──
        if code.startswith("6"):
            # En Odoo BE, les charges arrivent en négatif sur le compte analytique.
            # On prend -amt pour avoir un nombre positif = montant dépensé.
            # Les notes de crédit fournisseur arrivent en positif → -amt sera négatif
            # → elles réduisent bien les dépenses. ✓
            depenses_map[aid]  = depenses_map.get(aid, 0.0)  + (-amt)
            if year_start <= d <= year_end:
                depenses_annee_map[aid] = depenses_annee_map.get(aid, 0.0) + (-amt)
            if d >= date_12m:
                monthly_records.append({"aid": aid, "date": d, "type": "depense", "val": -amt})

        elif code.startswith("7"):
            # Produits : positif = revenu. NC client = négatif → réduit le revenu. ✓
            revenus_map[aid]  = revenus_map.get(aid, 0.0)  + amt
            if year_start <= d <= year_end:
                revenus_annee_map[aid] = revenus_annee_map.get(aid, 0.0) + amt
            if d >= date_12m:
                monthly_records.append({"aid": aid, "date": d, "type": "revenu", "val": amt})

        # Comptes 3, 4, 4xx transitoires → IGNORÉS volontairement
        # (acomptes : pas encore du CA ni de la dépense définitive)

    # ── 4) CA via sale.order ──
    code_to_proj = {}
    for p in project_list:
        code = extract_project_code(p.get("display_name", ""))
        if code:
            code_to_proj[code] = p

    ca_all_map   = {}
    ca_annee_map = {}

    if code_to_proj:
        all_so = models.execute_kw(
            DB, uid, PASSWORD,
            "sale.order", "search_read",
            [[("state", "in", ["sale", "done"])]],
            {"fields": ["id", "name", "amount_untaxed", "date_order"], "limit": 0}
        )
        for so in all_so:
            so_code = extract_project_code(so["name"])
            if not so_code:
                continue
            proj = code_to_proj.get(so_code)
            if not proj or not proj.get("analytic_account_id"):
                continue
            aid = proj["analytic_account_id"][0]
            amt = so["amount_untaxed"]
            ca_all_map[aid] = ca_all_map.get(aid, 0.0) + amt
            do = (so.get("date_order") or "")[:10]
            if year_start <= do <= year_end:
                ca_annee_map[aid] = ca_annee_map.get(aid, 0.0) + amt

    # ── 5) Synthèse par projet ──
    analytics_summary = {}
    for p in project_list:
        if not p.get("analytic_account_id"):
            analytics_summary[p["id"]] = None
            continue

        aid = p["analytic_account_id"][0]

        ca_total  = ca_all_map.get(aid, 0.0)
        ca_annee  = ca_annee_map.get(aid, 0.0)
        dep_all   = depenses_map.get(aid, 0.0)
        rev_all   = revenus_map.get(aid, 0.0)   # = ce qui est facturé (classe 7)

        # À facturer = CA commandé – déjà facturé (ne peut pas être négatif logiquement,
        # mais on laisse passer pour détecter les sur-facturations)
        a_facturer      = ca_total - rev_all
        marge_c         = ca_total - dep_all
        marge_pct       = (marge_c / ca_total * 100) if ca_total > 0 else 0.0

        dep_annee       = depenses_annee_map.get(aid, 0.0)
        rev_annee       = revenus_annee_map.get(aid, 0.0)
        marge_attendue  = ca_annee - dep_annee
        marge_att_pct   = (marge_attendue / ca_annee * 100) if ca_annee > 0 else 0.0
        a_fac_annee     = ca_annee - rev_annee

        analytics_summary[p["id"]] = {
            "ca_annee":           ca_annee,
            "depenses_annee":     dep_annee,
            "marge_attendue":     marge_attendue,
            "marge_attendue_pct": marge_att_pct,
            "a_facturer_annee":   a_fac_annee,

            "ca_total":    ca_total,
            "facture_all": rev_all,
            "a_facturer":  a_facturer,
            "depenses_all": dep_all,
            "marge_c":     marge_c,
            "marge_pct":   marge_pct,

            "is_closed":   p.get("is_closed", False),
        }

    # ── 6) DataFrame mensuel ──
    if not monthly_records:
        df_monthly = pd.DataFrame()
    else:
        df_m = pd.DataFrame(monthly_records)
        df_m["Mois"] = pd.to_datetime(df_m["date"]).dt.to_period("M").dt.to_timestamp()

        dep_agg = df_m[df_m["type"] == "depense"].groupby("Mois")["val"].sum().rename("Dépenses")
        rev_agg = df_m[df_m["type"] == "revenu"].groupby("Mois")["val"].sum().rename("CA")

        all_months = pd.date_range(start=date_12m, end=date.today().strftime("%Y-%m-%d"), freq="MS")
        df_monthly = pd.DataFrame({"Mois": all_months})
        df_monthly = (df_monthly
                      .merge(dep_agg.reset_index(), on="Mois", how="left")
                      .merge(rev_agg.reset_index(), on="Mois", how="left")
                      .fillna(0))

    return analytics_summary, df_monthly


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
    "Soudure":         "#1E88E5",
    "Peinture":        "#FDD835",
    "Assemblage":      "#43A047",
    "Câblage":         "#8E24AA",
    "Test":            "#FB8C00",
    "Montage":         "#E53935",
    "Mise en service": "#EC407A",
    "Réception":       "#6D4C41",
    "Transport":       "#00ACC1",
    "Etude":           "#34ebc6",
    "Autres":          "#9E9E9E"
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

        today = date.today()
        start_view = today
        end_view = today + timedelta(days=30 * months)

        # ── Construire le Gantt avec gestion des overlaps par quinconce ──
        gantt_data = []

        # Grouper par (projet, semaine_ISO) pour détecter les overlaps
        from collections import defaultdict
        week_task_groups = defaultdict(list)
        for t in tasks:
            proj = next((p for p in projects if p['id'] == t['project_id'][0]), None)
            if not proj:
                continue
            label = project_label(proj)
            deadline = t["date_deadline"]
            week_num = deadline.isocalendar()[1]
            week_year = deadline.year
            week_task_groups[(label, week_year, week_num)].append(t)

        for (label, week_year, week_num), group_tasks in week_task_groups.items():
            n = len(group_tasks)
            for idx, t in enumerate(group_tasks):
                deadline = t["date_deadline"]
                task_type = classify_task_type(t["name"])

                if n == 1:
                    # Tâche seule : barre normale centrée
                    y_base  = label
                    y_label = label
                    start   = deadline - timedelta(days=3)
                    end_d   = deadline + timedelta(days=3)
                else:
                    # Overlaps : on crée des sous-lignes alternées
                    # On utilise un suffixe pour créer des "sous-lignes" distinctes
                    # Pair → "haut", Impair → "bas"
                    sub = "▲" if idx % 2 == 0 else "▽"
                    y_label = f"{label} {sub}"
                    # Décalage horizontal léger pour quinconce visuel
                    offset = timedelta(days=1) if idx % 2 == 1 else timedelta(days=0)
                    start   = deadline - timedelta(days=2) + offset
                    end_d   = deadline + timedelta(days=2) + offset

                gantt_data.append({
                    "Tâche":    t["name"],
                    "Projet":   label,
                    "Y":        y_label,
                    "Début":    start,
                    "Fin":      end_d,
                    "Type":     task_type,
                    "is_done":  t.get("is_done", False),
                    "deadline": deadline,
                })

        if gantt_data:
            df_gantt = pd.DataFrame(gantt_data)
            df_gantt["code"] = df_gantt["Projet"].apply(extract_project_code)
            df_gantt = df_gantt.sort_values(by=["code", "Y"])

            # Ordre des Y : projets triés par code, sous-lignes groupées
            y_order = df_gantt.drop_duplicates("Y").sort_values(["code", "Y"])["Y"].tolist()
            y_order_rev = list(reversed(y_order))

            df_gantt["Type détaillé"] = pd.Categorical(
                df_gantt["Type"],
                categories=COLOR_ORDER,
                ordered=True
            )

            # ── Séparer tâches normales et terminées ──
            df_normal = df_gantt[~df_gantt["is_done"]].copy()
            df_done   = df_gantt[df_gantt["is_done"]].copy()

            fig = go.Figure()

            # Tâches normales par type (pour la légende)
            for task_type in COLOR_ORDER:
                df_t = df_normal[df_normal["Type"] == task_type]
                if df_t.empty:
                    continue
                fig.add_trace(go.Bar(
                    name=task_type,
                    x=[(row["Fin"] - row["Début"]).days for _, row in df_t.iterrows()],
                    y=df_t["Y"],
                    base=[row["Début"].strftime("%Y-%m-%d") for _, row in df_t.iterrows()],
                    orientation="h",
                    marker_color=COLOR_MAP[task_type],
                    marker_line_width=0,
                    width=0.5,
                    customdata=df_t[["Tâche", "Projet", "deadline"]].values,
                    hovertemplate=(
                        "<b>%{customdata[0]}</b><br>"
                        "Projet: %{customdata[1]}<br>"
                        "Deadline: %{customdata[2]}<br>"
                        "<extra></extra>"
                    ),
                    legendgroup=task_type,
                    showlegend=True,
                ))

            # Tâches terminées : hachurées (pattern)
            for task_type in COLOR_ORDER:
                df_t = df_done[df_done["Type"] == task_type]
                if df_t.empty:
                    continue
                fig.add_trace(go.Bar(
                    name=f"{task_type} ✓",
                    x=[(row["Fin"] - row["Début"]).days for _, row in df_t.iterrows()],
                    y=df_t["Y"],
                    base=[row["Début"].strftime("%Y-%m-%d") for _, row in df_t.iterrows()],
                    orientation="h",
                    marker=dict(
                        color=COLOR_MAP[task_type],
                        opacity=0.4,
                        pattern=dict(shape="/", fgcolor="white", size=4, solidity=0.5),
                        line=dict(color="white", width=1),
                    ),
                    width=0.5,
                    customdata=df_t[["Tâche", "Projet", "deadline"]].values,
                    hovertemplate=(
                        "<b>%{customdata[0]}</b> ✓ Terminée<br>"
                        "Projet: %{customdata[1]}<br>"
                        "Deadline: %{customdata[2]}<br>"
                        "<extra></extra>"
                    ),
                    legendgroup=f"{task_type}_done",
                    showlegend=False,
                ))

            n_proj = len(y_order)
            chart_height = max(500, n_proj * 22 + 160)

            fig.update_layout(
                barmode="overlay",
                dragmode="pan",
                height=chart_height,
                bargap=0.25,
                bargroupgap=0.05,
                margin=dict(l=20, r=20, t=40, b=20),
                yaxis=dict(
                    categoryorder="array",
                    categoryarray=y_order_rev,
                    tickfont=dict(size=11),
                    showgrid=True,
                    gridcolor="rgba(180,180,180,0.18)",
                    gridwidth=1,
                ),
                xaxis=dict(
                    type="date",
                    range=[start_view.strftime("%Y-%m-%d"), end_view.strftime("%Y-%m-%d")],
                    showgrid=False,
                ),
                plot_bgcolor="rgba(0,0,0,0)",
                legend=dict(
                    orientation="h", yanchor="bottom", y=1.02,
                    xanchor="center", x=0.5, font=dict(size=10)
                ),
            )

            # ── Ligne aujourd'hui ──
            fig.add_vline(x=today.strftime("%Y-%m-%d"), line_width=2,
                          line_color="white", opacity=0.9)

            # ── Séparateurs de mois (tirets) ──
            cur = date(today.year, today.month, 1)
            while True:
                cur = (date(cur.year + 1, 1, 1) if cur.month == 12
                       else date(cur.year, cur.month + 1, 1))
                if cur > end_view:
                    break
                fig.add_vline(x=cur.strftime("%Y-%m-%d"), line_width=1,
                              line_dash="dot", line_color="rgba(200,200,200,0.35)")

            # ── Séparateurs week-end (samedi) — très subtils ──
            cur_day = today - timedelta(days=today.weekday())  # lundi courant
            while cur_day <= end_view:
                saturday = cur_day + timedelta(days=5)
                sunday   = cur_day + timedelta(days=6)
                if today <= saturday <= end_view:
                    # Zone weekend colorée légèrement
                    fig.add_vrect(
                        x0=saturday.strftime("%Y-%m-%d"),
                        x1=(sunday + timedelta(days=1)).strftime("%Y-%m-%d"),
                        fillcolor="rgba(255,255,255,0.04)",
                        layer="below",
                        line_width=0,
                    )
                    # Ligne samedi
                    fig.add_vline(
                        x=saturday.strftime("%Y-%m-%d"),
                        line_width=1,
                        line_dash="dot",
                        line_color="rgba(150,150,150,0.18)",
                    )
                cur_day += timedelta(days=7)

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
        project_labels_map = {project_label(p): p["id"] for p in projects}
        selected_label = st.selectbox(
            "Sélectionne un projet pour afficher ses tâches",
            ["— Aucun —"] + list(project_labels_map.keys()),
            index=0
        )

        if selected_label != "— Aucun —":
            proj_id = project_labels_map[selected_label]
            tasks_for_project = sorted(
                [t for t in tasks if t["project_id"][0] == proj_id],
                key=lambda x: x["date_deadline"]
            )
            if tasks_for_project:
                for t in tasks_for_project:
                    date_str = t["date_deadline"].strftime("%d-%m-%Y")
                    done_str = " ✅" if t.get("is_done") else ""
                    st.write(f"- **{t['name']}**{done_str} — deadline : {date_str}")
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

                    pct_orange = 100 * summary["orange"] / total_safe
                    pct_grey   = 100 * summary["grey"]   / total_safe
                    pct_white  = 100 * summary["white"]  / total_safe
                    pct_blue   = 100 * summary["blue"]   / total_safe
                    pct_green  = 100 * summary["green"]  / total_safe

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
        year_now = date.today().year
        st.markdown("### 📊 Bilan analytique", unsafe_allow_html=True)

        projects_ana = load_projects_with_closed(uid, models, filter_mode)

        # Exclure comptes analytiques génériques
        bad_accounts = ["dépannage (liège)", "projets (lig)"]
        projects_ana = [
            p for p in projects_ana
            if not (
                p.get("analytic_account_id")
                and p["analytic_account_id"][1].lower() in bad_accounts
            )
        ]

        with st.spinner("Chargement des données analytiques…"):
            # ── UN SEUL appel pour tout ──
            analytics, df_monthly = load_all_analytics(uid, models, projects_ana)

        if not analytics:
            st.info("Aucune donnée analytique disponible.")
        else:
            # ── MÉTRIQUES RÉSUMÉ ──
            actifs_avec_ca = [
                p for p in projects_ana
                if not p.get("is_closed")
                and analytics.get(p["id"])
                and analytics[p["id"]]["ca_annee"] > 0
            ]

            sum_ca_annee   = sum(analytics[p["id"]]["ca_annee"]       for p in actifs_avec_ca)
            sum_dep_annee  = sum(analytics[p["id"]]["depenses_annee"]  for p in actifs_avec_ca)
            sum_marge_att  = sum(analytics[p["id"]]["marge_attendue"]  for p in actifs_avec_ca)
            sum_a_fac      = sum(analytics[p["id"]]["a_facturer_annee"] for p in actifs_avec_ca)
            marge_att_pct  = (sum_marge_att / sum_ca_annee * 100) if sum_ca_annee > 0 else 0.0

            st.markdown(
                f"<div style='font-size:13px;color:#aaa;margin-bottom:6px;'>"
                f"Projets confirmés en {year_now} (actifs, selon filtre)</div>",
                unsafe_allow_html=True
            )
            m1, m2, m3, m4 = st.columns(4)
            m1.metric(f"CA {year_now}",    fmt_eur(sum_ca_annee))
            m2.metric("Dépenses associées", fmt_eur(sum_dep_annee))
            m3.metric("Marge attendue",     fmt_eur(sum_marge_att),
                      delta=f"{marge_att_pct:.1f} %")
            m4.metric("Reste à facturer",   fmt_eur(sum_a_fac))

            st.markdown("---")

            # ── TABLEAU ──
            st.markdown("#### Détail par projet — toutes années")

            rows = []
            for p in projects_ana:
                ana = analytics.get(p["id"])
                if ana is None:
                    continue
                code      = extract_project_code(p["display_name"])
                proj_name = short_desc(clean_description_from_display_name(p["display_name"]), 45)
                rows.append({
                    "_closed":        p.get("is_closed", False),
                    "_sort":          (1 if p["is_closed"] else 0, code),
                    "Projet":         proj_name,
                    "Client":         p["company"],
                    "CA total (€)":   ana["ca_total"],
                    "Dépenses (€)":   ana["depenses_all"],
                    "Facturé (€)":    ana["facture_all"],
                    "À facturer (€)": ana["a_facturer"],
                    "Marge C (€)":    ana["marge_c"],
                    "Marge C (%)":    ana["marge_pct"],
                })

            if not rows:
                st.info("Aucune donnée trouvée.")
            else:
                df_ana = pd.DataFrame(rows)

                search = st.text_input("🔎 Recherche", "", placeholder="Projet ou client…", key="ana_search")
                if search:
                    s = search.lower()
                    df_ana = df_ana[
                        df_ana["Projet"].str.lower().str.contains(s)
                        | df_ana["Client"].str.lower().str.contains(s)
                    ]

                cols_def = "2fr 1.5fr 100px 110px 100px 110px 100px 80px"

                header_html = f"""
                <div style="display:grid;grid-template-columns:{cols_def};
                    column-gap:10px;padding:6px 12px;font-weight:bold;
                    font-size:12px;color:#aaa;border-bottom:2px solid #555;
                    margin-bottom:2px;position:sticky;top:0;background:#0e1117;z-index:10;">
                    <div>Projet</div>
                    <div>Client</div>
                    <div style="text-align:right;">CA Total</div>
                    <div style="text-align:right;">Dépenses</div>
                    <div style="text-align:right;">Facturé</div>
                    <div style="text-align:right;">À facturer</div>
                    <div style="text-align:right;">Marge C (€)</div>
                    <div style="text-align:right;">Marge C (%)</div>
                </div>"""

                rows_html = ""
                for _, row in df_ana.iterrows():
                    is_closed   = row["_closed"]
                    bg          = "#0d2a4a" if is_closed else "rgba(255,255,255,0.03)"
                    border      = "1px solid #1a4a7a" if is_closed else "1px solid #2a2a2a"
                    marge_color = (
                        "#e53935" if row["Marge C (€)"] < 0
                        else "#43a047" if row["Marge C (%)"] >= 20
                        else "#FB8C00"
                    )
                    # À facturer : négatif = sur-facturé (alerte orange)
                    afac_val   = row["À facturer (€)"]
                    afac_color = "#e53935" if afac_val < 0 else "#00ACC1"
                    closed_badge = (
                        " <span style='font-size:9px;background:#1565C0;color:white;"
                        "padding:1px 4px;border-radius:3px;vertical-align:middle;'>✓ Cloturé</span>"
                        if is_closed else ""
                    )

                    def fe(v): return f"{v:,.0f} €".replace(",", " ")
                    def fp(v): return f"{v:.1f} %"

                    rows_html += f"""
                    <div style="display:grid;grid-template-columns:{cols_def};
                        column-gap:10px;padding:6px 12px;font-size:13px;
                        background:{bg};border-bottom:{border};
                        align-items:center;min-height:32px;">
                        <div style="overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">
                            {row['Projet']}{closed_badge}
                        </div>
                        <div style="overflow:hidden;text-overflow:ellipsis;white-space:nowrap;color:#ccc;">
                            {row['Client']}
                        </div>
                        <div style="text-align:right;">{fe(row['CA total (€)'])}</div>
                        <div style="text-align:right;">{fe(row['Dépenses (€)'])}</div>
                        <div style="text-align:right;">{fe(row['Facturé (€)'])}</div>
                        <div style="text-align:right;color:{afac_color};font-weight:600;">
                            {fe(afac_val)}
                        </div>
                        <div style="text-align:right;color:{marge_color};font-weight:600;">
                            {fe(row['Marge C (€)'])}
                        </div>
                        <div style="text-align:right;color:{marge_color};">
                            {fp(row['Marge C (%)'])}
                        </div>
                    </div>"""

                st.markdown(
                    f"""
                    <div style="
                        border:1px solid #333;border-radius:6px;overflow:hidden;
                        max-height:420px;overflow-y:auto;background:#0e1117;
                    ">
                        {header_html}
                        <div>{rows_html}</div>
                    </div>
                    """,
                    unsafe_allow_html=True
                )

                st.markdown("---")

                # ── Graphique mensuel ──
                st.markdown("#### Évolution CA facturé & Dépenses — 12 derniers mois")

                if df_monthly.empty:
                    st.info("Pas de données mensuelles disponibles.")
                else:
                    df_m = df_monthly.copy()
                    df_m["Mois_label"] = pd.to_datetime(df_m["Mois"]).dt.strftime("%b %Y")

                    df_ca  = df_m[["Mois", "Mois_label", "CA"]].rename(columns={"CA": "Montant"})
                    df_ca["Série"] = "CA facturé (cl.7)"
                    df_dep = df_m[["Mois", "Mois_label", "Dépenses"]].rename(columns={"Dépenses": "Montant"})
                    df_dep["Série"] = "Dépenses (cl.6)"
                    df_plot = pd.concat([df_ca, df_dep])

                    fig_trend = px.bar(
                        df_plot,
                        x="Mois_label",
                        y="Montant",
                        color="Série",
                        barmode="group",
                        color_discrete_map={
                            "CA facturé (cl.7)": "#43a047",
                            "Dépenses (cl.6)":   "#e53935",
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
