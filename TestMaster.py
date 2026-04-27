import xmlrpc.client
from datetime import datetime, timedelta, date
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


def connect_odoo():
    common = xmlrpc.client.ServerProxy(f"{ODOO_URL}/xmlrpc/2/common")
    uid = common.authenticate(DB, USERNAME, PASSWORD, {})
    if not uid:
        raise Exception("Echec authentification Odoo")
    return uid, xmlrpc.client.ServerProxy(f"{ODOO_URL}/xmlrpc/2/object")


def get_top_companies_batch(uid, models, partner_ids):
    clean_ids = list(set(
        pid[0] if isinstance(pid, (list, tuple)) else pid
        for pid in partner_ids if pid
    ))
    if not clean_ids:
        return {}
    partners = models.execute_kw(DB, uid, PASSWORD, "res.partner", "read",
        [clean_ids], {"fields": ["id", "name", "parent_id"]})
    parent_ids = list({p["parent_id"][0] for p in partners if p["parent_id"]})
    parent_map = {}
    if parent_ids:
        parents = models.execute_kw(DB, uid, PASSWORD, "res.partner", "read",
            [parent_ids], {"fields": ["id", "name"]})
        parent_map = {p["id"]: p["name"] for p in parents}
    return {p["id"]: (parent_map[p["parent_id"][0]] if p["parent_id"] else p["name"])
            for p in partners}


def extract_project_code(display_name):
    if not display_name:
        return ""
    m = re.search(r"S\d{2}-\d{5}", display_name)
    return m.group(0) if m else ""


def clean_description_from_display_name(display_name):
    if not display_name or " - " not in display_name:
        return display_name or ""
    parts = display_name.split(" - ")
    if len(parts) >= 2 and parts[-1].strip() == parts[-2].strip():
        parts = parts[:-1]
    return " - ".join(parts)


def short_desc(desc, max_len):
    if not desc or len(desc) <= max_len:
        return desc or ""
    return desc[:max_len].rstrip() + "..."


def project_label(p):
    display = p.get("display_name") or p.get("name") or "Projet"
    return f"{p.get('company', 'N/A')} - {short_desc(clean_description_from_display_name(display), 20)}"


def fmt_eur(val):
    return f"{val:,.0f} EUR".replace(",", " ")


# ============================================================
# LOADERS
# ============================================================

def _get_tags(uid, models):
    eng  = models.execute_kw(DB, uid, PASSWORD, 'project.tags', 'search', [[('name', '=', 'Engineering')]])
    std  = models.execute_kw(DB, uid, PASSWORD, 'project.tags', 'search', [[('name', '=', 'Standard')]])
    prol = models.execute_kw(DB, uid, PASSWORD, 'project.tags', 'search', [[('name', 'ilike', 'PRO (LIG)')]])
    return eng, std, prol


@st.cache_data(ttl=300)
def load_projects(_uid, _models, filter_mode="both"):
    uid, models = _uid, _models
    eng, std, prol = _get_tags(uid, models)

    base = [('stage_id.name', 'not in', ['Cloture', 'Cloture', 'Template', 'Annule', 'Annule'])]
    if filter_mode == "engineering":
        domain = base + [('tag_ids', 'in', eng), ('tag_ids', 'in', prol)]
    elif filter_mode == "standard":
        domain = base + [('tag_ids', 'in', std), ('tag_ids', 'in', prol)]
    else:
        domain = base + ['|', ('tag_ids', 'in', eng), ('tag_ids', 'in', std), ('tag_ids', 'in', prol)]

    projects = models.execute_kw(DB, uid, PASSWORD, 'project.project', 'search_read',
        [domain], {'fields': ['id', 'display_name', 'partner_id', 'name', 'analytic_account_id']})

    company_map = get_top_companies_batch(uid, models, [p["partner_id"] for p in projects])
    for p in projects:
        pid = p["partner_id"][0] if p["partner_id"] else None
        p["company"] = company_map.get(pid, "N/A")

    ids = [p["id"] for p in projects]
    updates = models.execute_kw(DB, uid, PASSWORD, 'project.update', 'search_read',
        [[('project_id', 'in', ids)]], {'fields': ['project_id', 'status', 'write_date']})
    last_update = {}
    for u in updates:
        pid = u["project_id"][0]
        if pid not in last_update or u["write_date"] > last_update[pid]["write_date"]:
            last_update[pid] = u

    filtered = [p for p in projects if last_update.get(p["id"], {}).get("status") != "done"]
    filtered.sort(key=lambda p: extract_project_code(p['display_name']))
    return filtered


@st.cache_data(ttl=300)
def load_projects_with_closed(_uid, _models, filter_mode="both"):
    uid, models = _uid, _models
    eng, std, prol = _get_tags(uid, models)

    base = [('stage_id.name', 'not in', ['Template', 'Annule', 'Annule'])]
    if filter_mode == "engineering":
        domain = base + [('tag_ids', 'in', eng), ('tag_ids', 'in', prol)]
    elif filter_mode == "standard":
        domain = base + [('tag_ids', 'in', std), ('tag_ids', 'in', prol)]
    else:
        domain = base + ['|', ('tag_ids', 'in', eng), ('tag_ids', 'in', std), ('tag_ids', 'in', prol)]

    projects = models.execute_kw(DB, uid, PASSWORD, 'project.project', 'search_read',
        [domain], {'fields': ['id', 'display_name', 'partner_id', 'name', 'analytic_account_id', 'stage_id']})

    company_map = get_top_companies_batch(uid, models, [p["partner_id"] for p in projects])
    for p in projects:
        pid = p["partner_id"][0] if p["partner_id"] else None
        p["company"] = company_map.get(pid, "N/A")
        stage_name = p["stage_id"][1] if p.get("stage_id") else ""
        # Accepte les variantes avec/sans accent
        p["is_closed"] = "clotu" in stage_name.lower()

    projects.sort(key=lambda p: (
        1 if p["is_closed"] else 0,
        tuple(-ord(c) for c in extract_project_code(p['display_name']))
    ))
    return projects


def get_tasks(uid, models, project_ids, start_date, end_date):
    tasks = models.execute_kw(DB, uid, PASSWORD, 'project.task', 'search_read',
        [[('project_id', 'in', project_ids), ('date_deadline', '!=', False),
          ('tag_ids.name', 'in', ['Engineering', 'PRO (LIG)', 'PRO(LIG)'])]],
        {'fields': ['id', 'name', 'project_id', 'date_deadline', 'state', 'stage_id']})

    all_stage_ids = list({t['stage_id'][0] for t in tasks if t.get('stage_id')})
    closed_stages = set()
    if all_stage_ids:
        stages = models.execute_kw(DB, uid, PASSWORD, 'project.task.type', 'read',
            [all_stage_ids], {'fields': ['id', 'name']})
        name_done = {'done', 'termine', 'terminee', 'fini', 'finie', 'closed', 'annule', 'cancelled'}
        for s in stages:
            if s.get('is_closed') or s.get('name', '').lower().strip() in name_done:
                closed_stages.add(s['id'])

    for t in tasks:
        raw = t['date_deadline']
        if raw:
            t['date_deadline'] = datetime.strptime(raw.split(" ")[0], '%Y-%m-%d').date()
        state = str(t.get('state') or '').lower()
        stage_id = t['stage_id'][0] if t.get('stage_id') else None
        t['is_done'] = (any(kw in state for kw in ('done', 'cancel', 'annul', 'clos'))
                        or stage_id in closed_stages)
    return tasks


@st.cache_data(ttl=300)
def load_purchase_data_all_projects():
    uid, models = connect_odoo()
    po_data = models.execute_kw(DB, uid, PASSWORD, "purchase.order", "search_read",
        [[("state", "=", "purchase")]], {"fields": ["id", "user_id", "name"]})
    buyer_map   = {po["id"]: (po["user_id"][1] if po["user_id"] else "Unknown") for po in po_data}
    po_name_map = {po["id"]: po["name"] for po in po_data}
    po_ids = [po["id"] for po in po_data]

    po_lines = models.execute_kw(DB, uid, PASSWORD, "purchase.order.line", "search_read",
        [[("order_id", "in", po_ids)]],
        {"fields": ["name", "product_qty", "qty_received", "date_planned",
                    "order_id", "product_id", "analytic_distribution"]})

    product_ids = list({l["product_id"][0] for l in po_lines if l.get("product_id")})
    policy_map = {}
    if product_ids:
        products = models.execute_kw(DB, uid, PASSWORD, "product.product", "read",
            [product_ids], {"fields": ["type"]})
        policy_map = {p["id"]: p["type"] for p in products}

    return po_lines, policy_map, buyer_map, po_name_map


def get_purchase_for_project(project, po_lines, policy_map, buyer_map, po_name_map):
    today = date.today()
    counts = {"orange": 0, "grey": 0, "white": 0, "green": 0, "blue": 0}
    formatted = []

    analytic_id = project["analytic_account_id"][0] if project.get("analytic_account_id") else None
    if not analytic_id:
        counts["total"] = 0
        return counts, []

    for l in po_lines:
        dist = l.get("analytic_distribution") or {}
        if str(analytic_id) not in dist or l["product_qty"] == 0:
            continue

        qty_o = l["product_qty"]
        qty_r = l["qty_received"]
        dp = datetime.strptime(l["date_planned"].split(" ")[0], "%Y-%m-%d").date() if l["date_planned"] else None
        is_service = policy_map.get(l["product_id"][0] if l.get("product_id") else None, "") == "service"

        if qty_r >= qty_o:
            color, rank, key = "#2E7D32", 4, "green"
        elif qty_r > 0:
            color, rank, key = "#FFA000", 0, "orange"
        elif dp and dp < today:
            if is_service:
                color, rank, key = "#1565C0", 3, "blue"
            else:
                color, rank, key = "#757575", 1, "grey"
        else:
            if is_service:
                color, rank, key = "#1565C0", 3, "blue"
            else:
                color, rank, key = "#FFFFFF", 2, "white"

        counts[key] += 1
        formatted.append({
            "PO": po_name_map.get(l["order_id"][0], str(l["order_id"][0])),
            "Buyer": buyer_map.get(l["order_id"][0], "Unknown"),
            "Description": short_desc(l["name"], 50),
            "Ordered": qty_o, "Received": qty_r, "Planned Date": dp,
            "Color": color, "Rank": rank,
        })

    formatted.sort(key=lambda x: x["Rank"])
    counts["total"] = sum(counts[k] for k in ("orange", "grey", "white", "green", "blue"))
    return counts, formatted


@st.cache_data(ttl=300)
def load_all_analytics(_uid, _models, project_list):
    uid, models = _uid, _models

    analytic_ids = [p["analytic_account_id"][0] for p in project_list if p.get("analytic_account_id")]
    if not analytic_ids:
        return {}, pd.DataFrame(), 0.0

    year_now   = date.today().year
    year_start = f"{year_now}-01-01"
    year_end   = f"{year_now}-12-31"
    date_12m   = (date.today().replace(day=1) - timedelta(days=365)).strftime("%Y-%m-%d")

    # ── 1) Lignes analytiques (depenses : classe 6 + timesheets) ──
    all_lines = models.execute_kw(DB, uid, PASSWORD, "account.analytic.line", "search_read",
        [[("account_id", "in", analytic_ids)]],
        {"fields": ["account_id", "amount", "general_account_id", "date"], "limit": 0})

    acc_ids_list = list({l["general_account_id"][0] for l in all_lines if l.get("general_account_id")})
    account_code_map = {}
    for i in range(0, len(acc_ids_list), 200):
        for a in models.execute_kw(DB, uid, PASSWORD, "account.account", "read",
                [acc_ids_list[i:i+200]], {"fields": ["id", "code"]}):
            account_code_map[a["id"]] = a["code"]

    dep_map = {}
    dep_yr  = {}
    mo_dep  = []

    for line in all_lines:
        if not line.get("account_id"):
            continue
        aid = line["account_id"][0]
        amt = line["amount"]
        d   = line.get("date", "")

        if not line.get("general_account_id"):
            # Timesheet : montant negatif = cout
            if amt < 0:
                v = -amt
                dep_map[aid] = dep_map.get(aid, 0.0) + v
                if year_start <= d <= year_end:
                    dep_yr[aid] = dep_yr.get(aid, 0.0) + v
                if d >= date_12m:
                    mo_dep.append({"aid": aid, "date": d, "val": v})
            continue

        code = account_code_map.get(line["general_account_id"][0], "")
        if code.startswith("6"):
            # Odoo BE : facture fourn = negatif → -amt positif ; NC fourn = positif → -amt negatif
            v = -amt
            dep_map[aid] = dep_map.get(aid, 0.0) + v
            if year_start <= d <= year_end:
                dep_yr[aid] = dep_yr.get(aid, 0.0) + v
            if d >= date_12m:
                mo_dep.append({"aid": aid, "date": d, "val": v})

    # ── 2) CA via sale.order ──
    code_to_proj = {extract_project_code(p.get("display_name", "")): p
                    for p in project_list if extract_project_code(p.get("display_name", ""))}

    ca_all  = {}
    ca_yr   = {}
    inv_by_aid = {}   # aid -> [invoice_ids]

    all_so = models.execute_kw(DB, uid, PASSWORD, "sale.order", "search_read",
        [[("state", "in", ["sale", "done"])]],
        {"fields": ["name", "amount_untaxed", "date_order", "invoice_ids"], "limit": 0})

    for so in all_so:
        so_code = extract_project_code(so["name"])
        proj = code_to_proj.get(so_code)
        if not proj or not proj.get("analytic_account_id"):
            continue
        aid = proj["analytic_account_id"][0]
        amt = so["amount_untaxed"]
        ca_all[aid] = ca_all.get(aid, 0.0) + amt
        do = (so.get("date_order") or "")[:10]
        if year_start <= do <= year_end:
            ca_yr[aid] = ca_yr.get(aid, 0.0) + amt
        for inv_id in (so.get("invoice_ids") or []):
            inv_by_aid.setdefault(aid, []).append(inv_id)

    # ── 3) Facture via account.move (toutes SO, Engineering ET Standard) ──
    fact_all = {}
    fact_yr  = {}
    mo_rev   = []

    all_inv_ids = list({inv_id for ids in inv_by_aid.values() for inv_id in ids})
    inv_to_aid  = {inv_id: aid for aid, ids in inv_by_aid.items() for inv_id in ids}

    if all_inv_ids:
        all_moves = []
        for i in range(0, len(all_inv_ids), 200):
            all_moves.extend(models.execute_kw(DB, uid, PASSWORD, "account.move", "read",
                [all_inv_ids[i:i+200]],
                {"fields": ["id", "move_type", "state", "amount_untaxed", "invoice_date"]}))

        for move in all_moves:
            if move["state"] != "posted":
                continue
            aid = inv_to_aid.get(move["id"])
            if not aid:
                continue
            sign = +1 if move["move_type"] == "out_invoice" else (
                   -1 if move["move_type"] == "out_refund" else None)
            if sign is None:
                continue
            amt   = move["amount_untaxed"]
            inv_d = (move.get("invoice_date") or "")[:10]
            fact_all[aid] = fact_all.get(aid, 0.0) + sign * amt
            if year_start <= inv_d <= year_end:
                fact_yr[aid] = fact_yr.get(aid, 0.0) + sign * amt
            if inv_d >= date_12m:
                mo_rev.append({"date": inv_d, "val": sign * amt})

    # ── 4) DataFrame mensuel ──
    records = (
        [{"date": r["date"], "type": "dep", "val": r["val"]} for r in mo_dep] +
        [{"date": r["date"], "type": "rev", "val": r["val"]} for r in mo_rev]
    )
    if not records:
        df_monthly = pd.DataFrame()
    else:
        df_m = pd.DataFrame(records)
        df_m["Mois"] = pd.to_datetime(df_m["date"]).dt.to_period("M").dt.to_timestamp()
        d_agg = df_m[df_m["type"] == "dep"].groupby("Mois")["val"].sum().rename("Depenses")
        r_agg = df_m[df_m["type"] == "rev"].groupby("Mois")["val"].sum().rename("CA")
        months = pd.date_range(start=date_12m, end=date.today().strftime("%Y-%m-%d"), freq="MS")
        df_monthly = (pd.DataFrame({"Mois": months})
                      .merge(d_agg.reset_index(), on="Mois", how="left")
                      .merge(r_agg.reset_index(), on="Mois", how="left")
                      .fillna(0))

    # ── 5) Synthese par projet ──
    summary = {}
    for p in project_list:
        if not p.get("analytic_account_id"):
            summary[p["id"]] = None
            continue
        aid = p["analytic_account_id"][0]
        ca_t  = ca_all.get(aid, 0.0)
        ca_a  = ca_yr.get(aid, 0.0)
        dep_t = dep_map.get(aid, 0.0)
        dep_a = dep_yr.get(aid, 0.0)
        fac_t = fact_all.get(aid, 0.0)
        fac_a = fact_yr.get(aid, 0.0)
        marge = ca_t - dep_t
        summary[p["id"]] = {
            "ca_annee": ca_a, "depenses_annee": dep_a,
            "marge_attendue": ca_a - dep_a,
            "marge_attendue_pct": ((ca_a - dep_a) / ca_a * 100) if ca_a > 0 else 0.0,
            "a_facturer_annee": ca_a - fac_a,
            "ca_total": ca_t, "facture_all": fac_t,
            "a_facturer": ca_t - fac_t,
            "depenses_all": dep_t, "marge_c": marge,
            "marge_pct": (marge / ca_t * 100) if ca_t > 0 else 0.0,
            "is_closed": p.get("is_closed", False),
        }

    # ── 6) Marge ponderee projets clotures : sum(benefices) / sum(CA) ──
    sum_bene = sum_ca = 0.0
    for p in project_list:
        if not p.get("is_closed"):
            continue
        d = summary.get(p["id"])
        if not d or d["ca_total"] <= 0 or d["marge_pct"] > 70 or d["marge_pct"] < -100:
            continue
        sum_bene += d["marge_c"]
        sum_ca   += d["ca_total"]
    marge_pond = (sum_bene / sum_ca * 100) if sum_ca > 0 else 0.0

    return summary, df_monthly, marge_pond


# ============================================================
# GANTT
# ============================================================

COLOR_ORDER = ["Soudure", "Peinture", "Assemblage", "Cablage", "Test",
               "Montage", "Mise en service", "Reception", "Transport", "Etude", "Autres"]

COLOR_MAP = {
    "Soudure": "#1E88E5", "Peinture": "#FDD835", "Assemblage": "#43A047",
    "Cablage": "#8E24AA", "Test": "#FB8C00", "Montage": "#E53935",
    "Mise en service": "#EC407A", "Reception": "#6D4C41",
    "Transport": "#00ACC1", "Etude": "#34ebc6", "Autres": "#9E9E9E"
}

COLOR_DISPLAY = {
    "Soudure": "Soudure", "Peinture": "Peinture", "Assemblage": "Assemblage",
    "Cablage": "Cablage", "Test": "Test", "Montage": "Montage",
    "Mise en service": "Mise en service", "Reception": "Reception",
    "Transport": "Transport", "Etude": "Etude", "Autres": "Autres"
}


def classify_task_type(name):
    n = name.lower()
    if "soud" in n: return "Soudure"
    if "peint" in n: return "Peinture"
    if "assembl" in n: return "Assemblage"
    if "cabl" in n: return "Cablage"
    if "test" in n: return "Test"
    if "montage" in n or "installation" in n: return "Montage"
    if "mise en service" in n or " mes" in n: return "Mise en service"
    if "recept" in n or "assistance" in n: return "Reception"
    if "transport" in n: return "Transport"
    if "etude" in n or "conception" in n or "plan" in n or "calcul" in n: return "Etude"
    return "Autres"


def build_weeks_horizon(months=3):
    start = date.today()
    end = start + timedelta(days=30 * months)
    current = start - timedelta(days=start.weekday())
    weeks = []
    while current <= end:
        weeks.append((current.isocalendar()[1], current, current + timedelta(days=6)))
        current += timedelta(days=7)
    return weeks


def map_tasks_to_grid(projects, tasks, weeks):
    proj_index = {p['id']: i for i, p in enumerate(projects)}
    grid, detailed = {}, {}
    for t in tasks:
        pid = t['project_id'][0]
        if pid not in proj_index:
            continue
        row = proj_index[pid]
        for col, (_, sw, ew) in enumerate(weeks):
            if sw <= t['date_deadline'] <= ew:
                key = (row, col)
                grid.setdefault(key, []).append(COLOR_MAP[classify_task_type(t['name'])])
                detailed.setdefault(key, []).append(t)
                break
    return grid, detailed


# ============================================================
# MAIN APP
# ============================================================

def main():
    st.set_page_config(page_title="Master Planning Odoo", layout="wide")
    st.markdown("""<style>
    .block-container{padding-top:0.5rem!important;}
    div[data-testid="stToggle"]>label{font-size:13px!important;}
    </style>""", unsafe_allow_html=True)

    try:
        uid, models = connect_odoo()
    except Exception as e:
        st.error(f"Connexion Odoo impossible : {e}")
        return

    st_autorefresh(interval=600000, key="refresh_10min")

    for k, v in [("months", 3), ("selected_purchase_project_id", None),
                 ("filter_engineering", True), ("filter_standard", False)]:
        if k not in st.session_state:
            st.session_state[k] = v

    # Banniere
    c1, c2, c3 = st.columns([1, 4, 1])
    with c1:
        st.image("https://upload.wikimedia.org/wikipedia/commons/b/ba/Olsen-Logo.png", width=180)
        st.markdown("<div style='color:green;font-weight:bold;margin-top:20px;'>Connecte Odoo</div>",
                    unsafe_allow_html=True)
    with c2:
        st.markdown("<h2 style='text-align:center;margin-top:10px;'>Olsen Dashboard</h2>",
                    unsafe_allow_html=True)
    with c3:
        fe = st.toggle("Engineering (PRO LIG)", value=st.session_state["filter_engineering"],
                       key="toggle_engineering")
        fs = st.toggle("Standard (PRO LIG)", value=st.session_state["filter_standard"],
                       key="toggle_standard")
        if not fe and not fs:
            st.warning("Au moins un filtre actif.")
            fe = True
        fm = "both" if fe and fs else "engineering" if fe else "standard"
        if fe != st.session_state["filter_engineering"] or fs != st.session_state["filter_standard"]:
            st.session_state["filter_engineering"] = fe
            st.session_state["filter_standard"] = fs
            st.rerun()

    tab1, tab2, tab3 = st.tabs(["Planning", "Purchases", "Analytique"])

    # ── ONGLET 1 : PLANNING ──────────────────────────────────────
    with tab1:
        projects = load_projects(uid, models, fm)
        months   = st.session_state["months"]
        weeks    = build_weeks_horizon(months)
        tasks    = get_tasks(uid, models, [p['id'] for p in projects], weeks[0][1], weeks[-1][2])
        grid, detailed = map_tasks_to_grid(projects, tasks, weeks)

        st.subheader("Gantt")
        today      = date.today()
        start_view = today
        end_view   = today + timedelta(days=30 * months)

        gantt_data = []
        overlap_counter = {}
        for t in tasks:
            proj = next((p for p in projects if p['id'] == t['project_id'][0]), None)
            if not proj:
                continue
            label    = project_label(proj)
            deadline = t["date_deadline"]
            wk       = (label, deadline.isocalendar()[1], deadline.year)
            cnt      = overlap_counter.get(wk, 0)
            overlap_counter[wk] = cnt + 1
            gantt_data.append({
                "Tache":   t["name"],
                "Projet":  label,
                "Debut":   deadline - timedelta(days=3) + timedelta(days=cnt),
                "Fin":     deadline + timedelta(days=3) + timedelta(days=cnt),
                "Type":    classify_task_type(t["name"]),
                "is_done": t.get("is_done", False),
                "deadline_str": str(deadline),
            })

        if gantt_data:
            df_gantt = pd.DataFrame(gantt_data)
            df_gantt["code"]   = df_gantt["Projet"].apply(extract_project_code)
            df_gantt           = df_gantt.sort_values("code")
            df_gantt["TypeCat"] = pd.Categorical(df_gantt["Type"], categories=COLOR_ORDER, ordered=True)

            # Gantt de base avec les couleurs normales (legende integre)
            fig = px.timeline(df_gantt, x_start="Debut", x_end="Fin", y="Projet",
                              color="TypeCat", color_discrete_map=COLOR_MAP,
                              hover_name="Tache",
                              hover_data={"Debut": True, "Fin": True, "TypeCat": True, "Projet": False})

            # Renommer les traces pour la legende
            for trace in fig.data:
                trace.name = COLOR_DISPLAY.get(trace.name, trace.name)

            # Masque semi-transparent noir sur les taches terminees
            # → la couleur d'origine reste visible, la legende n'est pas touchee
            df_done = df_gantt[df_gantt["is_done"]].copy()
            if not df_done.empty:
                fig.add_trace(go.Bar(
                    name="Termine",
                    x=[(row["Fin"] - row["Debut"]).days for _, row in df_done.iterrows()],
                    y=df_done["Projet"],
                    base=[str(row["Debut"]) for _, row in df_done.iterrows()],
                    orientation="h",
                    marker=dict(color="rgba(0,0,0,0.50)", line=dict(width=0)),
                    width=0.85,
                    text=df_done["Tache"],
                    hovertemplate="<b>%{text}</b> (Termine)<extra></extra>",
                    showlegend=True,
                ))

            n_proj = len(df_gantt["Projet"].unique())
            fig.update_layout(
                barmode="overlay",
                dragmode="pan",
                height=max(500, n_proj * 18 + 140),
                bargap=0.3, bargroupgap=0.1,
                margin=dict(l=20, r=20, t=40, b=20),
                yaxis=dict(autorange="reversed", tickfont=dict(size=12),
                           showgrid=True, gridcolor="rgba(180,180,180,0.18)"),
                xaxis=dict(showgrid=False),
                plot_bgcolor="rgba(0,0,0,0)",
                legend=dict(orientation="h", yanchor="bottom", y=1.02,
                            xanchor="center", x=0.5, font=dict(size=10))
            )
            fig.update_xaxes(range=[start_view, end_view])

            # Ligne aujourd'hui
            fig.add_vline(x=today, line_width=2, line_color="white", opacity=0.9)

            # Separateurs mois
            cur = date(today.year, today.month, 1)
            while True:
                cur = date(cur.year + 1, 1, 1) if cur.month == 12 else date(cur.year, cur.month + 1, 1)
                if cur > end_view:
                    break
                fig.add_vline(x=cur, line_width=1, line_dash="dot", line_color="rgba(200,200,200,0.35)")

            # Separateurs week-end
            cur_day = today - timedelta(days=today.weekday())
            while cur_day <= end_view:
                sat = cur_day + timedelta(days=5)
                mon = cur_day + timedelta(days=7)
                if sat <= end_view:
                    fig.add_vrect(x0=sat, x1=mon,
                        fillcolor="rgba(255,255,255,0.04)", layer="below", line_width=0)
                    fig.add_vline(x=sat, line_width=1, line_dash="dot",
                                  line_color="rgba(160,160,160,0.20)")
                cur_day += timedelta(days=7)

            st.plotly_chart(fig, use_container_width=True, config={"displaylogo": False})
        else:
            st.info("Aucune tache a afficher dans le Gantt.")

        c_sl, c_nb = st.columns([3, 1])
        with c_sl:
            new_months = st.slider("", 1, 6, months)
        with c_nb:
            st.markdown(f"<div style='margin-top:10px;font-size:14px;'>Projets : <b>{len(projects)}</b></div>",
                        unsafe_allow_html=True)
        if new_months != months:
            st.session_state["months"] = new_months
            st.rerun()

        st.subheader("Taches du projet")
        proj_map = {project_label(p): p["id"] for p in projects}
        sel = st.selectbox("Projet", ["Aucun"] + list(proj_map.keys()), index=0)
        if sel != "Aucun":
            tid = proj_map[sel]
            tlist = sorted([t for t in tasks if t["project_id"][0] == tid],
                           key=lambda x: x["date_deadline"])
            if tlist:
                for t in tlist:
                    wd      = t["date_deadline"].weekday()
                    we_flag = " **[WE]**" if wd >= 5 else ""
                    done    = " (Termine)" if t.get("is_done") else ""
                    st.write(f"- **{t['name']}**{done}{we_flag} — {t['date_deadline'].strftime('%d-%m-%Y')}")
            else:
                st.info("Aucune tache pour ce projet.")
        else:
            st.info("Selectionne un projet.")

    # ── ONGLET 2 : PURCHASES ─────────────────────────────────────
    with tab2:
        st.markdown("### Purchases par projet")
        projects_all = load_projects(uid, models, fm)
        po_lines, policy_map, buyer_map, po_name_map = load_purchase_data_all_projects()

        purchase_data = {p['id']: get_purchase_for_project(p, po_lines, policy_map, buyer_map, po_name_map)
                         for p in projects_all}

        for i in range(0, len(projects_all), 6):
            cols = st.columns(6)
            for col, p in zip(cols, projects_all[i:i+6]):
                with col:
                    sm, _ = purchase_data[p['id']]
                    tot   = max(sm["total"], 1)
                    tc    = "red" if sm["grey"] > 0 else "#FFA000" if sm["orange"] > 0 else "white"
                    if st.button(f"{p['company']}\n {short_desc(clean_description_from_display_name(p['display_name']), 25)}",
                                 key=f"proj_btn_{p['id']}"):
                        st.session_state["selected_purchase_project_id"] = p['id']
                    st.markdown(f"""
                        <div style="width:100%;height:12px;border-radius:6px;overflow:hidden;
                            display:flex;margin-top:4px;border:1px solid #444;">
                            <div style="width:{100*sm['orange']//tot}%;background:#FFA000;"></div>
                            <div style="width:{100*sm['grey']//tot}%;background:#757575;"></div>
                            <div style="width:{100*sm['white']//tot}%;background:#FFFFFF;"></div>
                            <div style="width:{100*sm['blue']//tot}%;background:#1565C0;"></div>
                            <div style="width:{100*sm['green']//tot}%;background:#2E7D32;"></div>
                        </div>
                        <div style="text-align:right;font-size:12px;color:{tc};margin-top:2px;">
                            {sm['green']} / {sm['total']} lignes</div>""",
                        unsafe_allow_html=True)

        st.markdown("---")
        st.subheader("Detail lignes d'achat")
        sel_id = st.session_state.get("selected_purchase_project_id")
        if sel_id is None:
            st.info("Clique sur une vignette pour voir le detail.")
        else:
            p = next((p for p in projects_all if p['id'] == sel_id), None)
            if p:
                st.markdown(f"**{p['company']} - {p.get('name') or p['display_name']}**")
                _, lines = purchase_data[p['id']]
                if not lines:
                    st.info("Aucune ligne.")
                else:
                    st.markdown(f"**{len(lines)} lignes**")
                    for row in lines:
                        dd = row['Planned Date'].strftime("%d-%m-%Y") if row['Planned Date'] else "-"
                        tc = "white" if row['Color'] in ("#1565C0", "#2E7D32", "#757575") else "black"
                        st.markdown(f"""<div style="background:{row['Color']};padding:8px 12px;
                            border-radius:4px;margin-bottom:5px;border:1px solid #555;font-size:14px;
                            color:{tc};display:grid;
                            grid-template-columns:90px 190px 1fr 80px 90px 110px;
                            column-gap:12px;align-items:center;">
                            <div><b>PO:</b> {row['PO']}</div>
                            <div><b>Buyer:</b> {row['Buyer']}</div>
                            <div><b>Desc:</b> {row['Description']}</div>
                            <div><b>Ord.:</b> {row['Ordered']}</div>
                            <div><b>Recu:</b> {row['Received']}</div>
                            <div><b>Date:</b> {dd}</div>
                        </div>""", unsafe_allow_html=True)

    # ── ONGLET 3 : ANALYTIQUE ────────────────────────────────────
    with tab3:
        year_now = date.today().year
        st.markdown("### Bilan analytique")

        projects_ana = load_projects_with_closed(uid, models, fm)
        bad_accs = ["depannage (liege)", "projets (lig)", "vente pure (lig)"]
        projects_ana = [p for p in projects_ana
                        if not (p.get("analytic_account_id")
                                and p["analytic_account_id"][1].lower() in bad_accs)]

        with st.spinner("Chargement analytiques..."):
            analytics, df_monthly, marge_pond = load_all_analytics(uid, models, projects_ana)

        if not analytics:
            st.info("Aucune donnee disponible.")
            return

        actifs = [p for p in projects_ana
                  if not p.get("is_closed") and analytics.get(p["id"])
                  and analytics[p["id"]]["ca_annee"] > 0]

        s_ca  = sum(analytics[p["id"]]["ca_annee"]        for p in actifs)
        s_dep = sum(analytics[p["id"]]["depenses_annee"]   for p in actifs)
        s_ma  = sum(analytics[p["id"]]["marge_attendue"]   for p in actifs)
        s_fac = sum(analytics[p["id"]]["a_facturer_annee"] for p in actifs)

        st.markdown(f"<div style='font-size:13px;color:#aaa;margin-bottom:6px;'>"
                    f"Projets confirmes {year_now} (actifs)</div>", unsafe_allow_html=True)

        m1, m2, m3, m4 = st.columns(4)
        m1.metric(f"Sales {year_now}", fmt_eur(s_ca))
        m2.metric(f"Achats+Timesheets {year_now}", fmt_eur(s_dep))
        m3.metric("Marge reelle (clotures)", f"{marge_pond:.1f} %",
                  help="Somme benefices / Somme CA, projets clotures")
        m4.metric("A facturer (annee)", fmt_eur(s_fac))

        st.markdown("---")
        st.markdown("#### Detail par projet")

        rows = []
        for p in projects_ana:
            a = analytics.get(p["id"])
            if a is None:
                continue
            rows.append({
                "_closed":    p.get("is_closed", False),
                "Projet":     short_desc(clean_description_from_display_name(p["display_name"]), 45),
                "Client":     p["company"],
                "CA":         a["ca_total"],
                "Depenses":   a["depenses_all"],
                "Facture":    a["facture_all"],
                "A_fac":      a["a_facturer"],
                "Marge_EUR":  a["marge_c"],
                "Marge_PCT":  a["marge_pct"],
            })

        if not rows:
            st.info("Aucune donnee.")
        else:
            df_ana = pd.DataFrame(rows)
            search = st.text_input("Recherche", "", placeholder="Projet ou client...", key="ana_search")
            if search:
                s = search.lower()
                df_ana = df_ana[df_ana["Projet"].str.lower().str.contains(s)
                                | df_ana["Client"].str.lower().str.contains(s)]

            cd = "2fr 1.5fr 100px 110px 100px 110px 100px 80px"
            hdr = f"""<div style="display:grid;grid-template-columns:{cd};column-gap:10px;
                padding:6px 12px;font-weight:bold;font-size:12px;color:#aaa;
                border-bottom:2px solid #555;position:sticky;top:0;background:#0e1117;z-index:10;">
                <div>Projet</div><div>Client</div>
                <div style="text-align:right;">CA Total</div>
                <div style="text-align:right;">Depenses</div>
                <div style="text-align:right;">Facture</div>
                <div style="text-align:right;">A facturer</div>
                <div style="text-align:right;">Marge EUR</div>
                <div style="text-align:right;">Marge %</div>
            </div>"""

            body = ""
            for _, row in df_ana.iterrows():
                cl   = row["_closed"]
                bg   = "#0d2a4a" if cl else "rgba(255,255,255,0.03)"
                bdr  = "1px solid #1a4a7a" if cl else "1px solid #2a2a2a"
                mc   = "#e53935" if row["Marge_EUR"] < 0 else "#43a047" if row["Marge_PCT"] >= 20 else "#FB8C00"
                afc  = "#e53935" if row["A_fac"] < 0 else "#00ACC1"
                bdg  = (" <span style='font-size:9px;background:#1565C0;color:white;"
                        "padding:1px 4px;border-radius:3px;'>Cloture</span>" if cl else "")

                def fe(v): return f"{v:,.0f}".replace(",", " ") + " EUR"
                def fp(v): return f"{v:.1f} %"

                body += f"""<div style="display:grid;grid-template-columns:{cd};column-gap:10px;
                    padding:6px 12px;font-size:13px;background:{bg};border-bottom:{bdr};
                    align-items:center;min-height:32px;">
                    <div style="overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">
                        {row['Projet']}{bdg}</div>
                    <div style="overflow:hidden;text-overflow:ellipsis;white-space:nowrap;color:#ccc;">
                        {row['Client']}</div>
                    <div style="text-align:right;">{fe(row['CA'])}</div>
                    <div style="text-align:right;">{fe(row['Depenses'])}</div>
                    <div style="text-align:right;">{fe(row['Facture'])}</div>
                    <div style="text-align:right;color:{afc};font-weight:600;">{fe(row['A_fac'])}</div>
                    <div style="text-align:right;color:{mc};font-weight:600;">{fe(row['Marge_EUR'])}</div>
                    <div style="text-align:right;color:{mc};">{fp(row['Marge_PCT'])}</div>
                </div>"""

            st.markdown(f"""<div style="border:1px solid #333;border-radius:6px;overflow:hidden;
                max-height:420px;overflow-y:auto;background:#0e1117;">
                {hdr}<div>{body}</div></div>""", unsafe_allow_html=True)

            st.markdown("---")
            st.markdown("#### Evolution CA facture & Depenses — 12 derniers mois")

            if df_monthly.empty:
                st.info("Pas de donnees mensuelles.")
            else:
                dm = df_monthly.copy()
                dm["Mois_label"] = pd.to_datetime(dm["Mois"]).dt.strftime("%b %Y")
                dep_col = "Depenses" if "Depenses" in dm.columns else "Depenses"

                df_p = pd.concat([
                    dm[["Mois_label", "CA"]].rename(columns={"CA": "M"}).assign(S="CA facture"),
                    dm[["Mois_label", dep_col]].rename(columns={dep_col: "M"}).assign(S="Depenses"),
                ])
                fig2 = px.bar(df_p, x="Mois_label", y="M", color="S", barmode="group",
                              color_discrete_map={"CA facture": "#43a047", "Depenses": "#e53935"},
                              height=380, labels={"Mois_label": "", "M": "EUR"})
                fig2.update_layout(margin=dict(l=10, r=10, t=20, b=20),
                                   plot_bgcolor="rgba(0,0,0,0)",
                                   xaxis=dict(tickfont=dict(size=11), showgrid=False),
                                   yaxis=dict(tickfont=dict(size=11), showgrid=True,
                                              gridcolor="rgba(180,180,180,0.12)", tickformat=",.0f"),
                                   legend=dict(orientation="h", y=1.05, x=0),
                                   bargap=0.2, bargroupgap=0.05)
                st.plotly_chart(fig2, use_container_width=True, config={"displaylogo": False})

    st.markdown("""<style>.footer{position:fixed;left:0;bottom:0;width:100%;
        background:rgba(240,240,240,0.85);color:#333;text-align:center;
        padding:6px 0;font-size:14px;border-top:1px solid #ccc;z-index:9999;}</style>
        <div class="footer">C Flow - Powered by Olsen-Engineering</div>""",
        unsafe_allow_html=True)


if __name__ == "__main__":
    main()
