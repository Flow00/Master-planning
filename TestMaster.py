import xmlrpc.client
from datetime import datetime, timedelta, date
import re
import json
import html
from pathlib import Path
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from streamlit_autorefresh import st_autorefresh

# ---------- CONFIG ODOO ----------
from cryptography.fernet import Fernet as _F

_ODOO_URL = b"gAAAAABqE2o7n0-46Mrq_zGgEImejUrqFcUfa2KK6mOb6DziDNKwkdMAc4elmHcK5QIRBKW1Fv7nZADDeyOoW1ZwC6bwIqwTwHtsBF9m1p6m8K92k3pG0aWYOHRLUh7xd01QS5dr4ruf"
_DB       = b"gAAAAABqE2o7uI4dQh6jpjUu3vJxnAT69g8bnDPyExgVcoLFHkPQ9Gu6awqPiGpIBJcnyMTawHeLp9u3LUIxgiZ-2eQPjk3d_37HueqIKz6kd-muNKHQpMA="
_USERNAME = b"gAAAAABqE2o7kHXd143Tp1dLyfoJeyfL9x9ec2WhX_7-SQSDnmXo7r0sLJtJ_g8aBhdt60SmpsX1VoJIk-GqMQYfvjwKTrSJkgmy0ith4q0FJAAB8auYSfs="
_PASSWORD = b"gAAAAABqE2o7E-OjKOiLbiaT5ao7M4c8gF8Vmg88jM8aWPe0HdunsMyLHf44NgmedtAnUmPpv43hG0JBmr1BVbXXxIboULh4wKD47KCzbWURt0WZOm6SYpXmXXhtXf08xAt-as0a7GnS"

# Variables remplies au premier appel (voir _load_credentials)
ODOO_URL = DB = USERNAME = PASSWORD = None


def _load_credentials():
    global ODOO_URL, DB, USERNAME, PASSWORD
    if ODOO_URL is not None:        # déjà chargé
        return
    _f = _F(_get_key())             # _get_key existe maintenant (fichier déjà lu)
    ODOO_URL = _f.decrypt(_ODOO_URL).decode()
    DB       = _f.decrypt(_DB).decode()
    USERNAME = _f.decrypt(_USERNAME).decode()
    PASSWORD = _f.decrypt(_PASSWORD).decode()


@st.cache_data(ttl=3600)
def _odoo_uid():
    """Authentifie sur Odoo et renvoie uid. Mis en cache 1h car auth = round-trip réseau."""
    _load_credentials()
    common = xmlrpc.client.ServerProxy(f"{ODOO_URL}/xmlrpc/2/common")
    uid = common.authenticate(DB, USERNAME, PASSWORD, {})
    if not uid:
        raise Exception("Échec authentification Odoo")
    return uid


def connect_odoo():
    """Renvoie (uid, models). uid est caché, models est recréé à chaque appel
    car ServerProxy n'est pas thread-safe et garde un état HTTP interne :
    le partager entre reruns provoque http.client.CannotSendRequest."""
    _load_credentials()
    uid = _odoo_uid()
    models = xmlrpc.client.ServerProxy(f"{ODOO_URL}/xmlrpc/2/object")
    return uid, models


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

@st.cache_data(ttl=600)
def _get_tags(_uid, _models):
    uid, models = _uid, _models
    eng  = models.execute_kw(DB, uid, PASSWORD, 'project.tags', 'search', [[('name', '=', 'Engineering')]])
    std  = models.execute_kw(DB, uid, PASSWORD, 'project.tags', 'search', [[('name', '=', 'Standard')]])
    prol = models.execute_kw(DB, uid, PASSWORD, 'project.tags', 'search', [[('name', 'ilike', 'PRO (LIG)')]])
    return eng, std, prol


@st.cache_data(ttl=300)
def load_projects(_uid, _models, filter_mode="both"):
    uid, models = _uid, _models
    eng, std, prol = _get_tags(uid, models)

    # `active=True` exclut les projets archivés.
    # On exclut aussi les stages clôturés / template / annulés (variantes).
    base = [
        ('active', '=', True),
        ('stage_id.name', 'not in', ['Cloturé', 'Cloture', 'Template', 'Annulé', 'Annule', 'Annulée', 'Annulee', 'Cancelled', 'Canceled', 'Cancel']),
    ]
    if filter_mode == "engineering":
        domain = base + [('tag_ids', 'in', eng), ('tag_ids', 'in', prol)]
    elif filter_mode == "standard":
        domain = base + [('tag_ids', 'in', std), ('tag_ids', 'in', prol)]
    else:
        domain = base + ['|', ('tag_ids', 'in', eng), ('tag_ids', 'in', std), ('tag_ids', 'in', prol)]

    projects = models.execute_kw(DB, uid, PASSWORD, 'project.project', 'search_read',
        [domain], {'fields': ['id', 'display_name', 'partner_id', 'name', 'account_id', 'stage_id', 'date']})
    # AJOUT : alias account_id -> analytic_account_id pour le reste du script
    for p in projects:
        p['analytic_account_id'] = p.pop('account_id', None)

    # Filet de sécurité Python : exclure tout stage contenant "annul" ou "cancel"
    # (couvre les libellés exotiques non listés ci-dessus).
    def _is_cancelled_stage(p):
        name = (p.get("stage_id")[1] if p.get("stage_id") else "") or ""
        n = name.lower()
        return "annul" in n or "cancel" in n
    projects = [p for p in projects if not _is_cancelled_stage(p)]

    company_map = get_top_companies_batch(uid, models, [p["partner_id"] for p in projects])
    for p in projects:
        pid = p["partner_id"][0] if p["partner_id"] else None
        p["company"] = company_map.get(pid, "N/A")
        p["stage"]   = p["stage_id"][1] if p.get("stage_id") else "—"
        # date de fin du projet : champ "date" sur project.project (peut être False)
        raw = p.get("date")
        if raw:
            try:
                p["date_end"] = datetime.strptime(str(raw).split(" ")[0], "%Y-%m-%d").date()
            except Exception:
                p["date_end"] = None
        else:
            p["date_end"] = None

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

    base = [
        ('active', '=', True),
        ('stage_id.name', 'not in', ['Template', 'Annulé', 'Annule', 'Annulée', 'Annulee', 'Cancelled', 'Canceled', 'Cancel']),
    ]
    if filter_mode == "engineering":
        domain = base + [('tag_ids', 'in', eng), ('tag_ids', 'in', prol)]
    elif filter_mode == "standard":
        domain = base + [('tag_ids', 'in', std), ('tag_ids', 'in', prol)]
    else:
        domain = base + ['|', ('tag_ids', 'in', eng), ('tag_ids', 'in', std), ('tag_ids', 'in', prol)]

    projects = models.execute_kw(DB, uid, PASSWORD, 'project.project', 'search_read',
        [domain], {'fields': ['id', 'display_name', 'partner_id', 'name', 'account_id', 'stage_id', 'date']})
     # AJOUT : alias account_id -> analytic_account_id pour le reste du script
    for p in projects:
        p['analytic_account_id'] = p.pop('account_id', None)
    # Filet de sécurité Python : exclure les libellés exotiques "annul"/"cancel".
    def _is_cancelled_stage(p):
        name = (p.get("stage_id")[1] if p.get("stage_id") else "") or ""
        n = name.lower()
        return "annul" in n or "cancel" in n
    projects = [p for p in projects if not _is_cancelled_stage(p)]

    company_map = get_top_companies_batch(uid, models, [p["partner_id"] for p in projects])
    for p in projects:
        pid = p["partner_id"][0] if p["partner_id"] else None
        p["company"] = company_map.get(pid, "N/A")
        stage_name = p["stage_id"][1] if p.get("stage_id") else ""
        p["is_closed"] = "clotu" in stage_name.lower()
        p["stage"]     = stage_name or "—"
        raw = p.get("date")
        if raw:
            try:
                p["date_end"] = datetime.strptime(str(raw).split(" ")[0], "%Y-%m-%d").date()
            except Exception:
                p["date_end"] = None
        else:
            p["date_end"] = None

    projects.sort(key=lambda p: (
        1 if p["is_closed"] else 0,
        tuple(-ord(c) for c in extract_project_code(p['display_name']))
    ))
    return projects


@st.cache_data(ttl=300)
def get_tasks(_uid, _models, project_ids, start_date, end_date):
    uid, models = _uid, _models
    # Détection du champ date de début selon la version Odoo
    # On essaie planned_date_begin (Odoo 16/17) puis date_start (14/15)
    start_field = None
    for candidate in ('planned_date_start', 'planned_date_begin'):
        try:
            models.execute_kw(DB, uid, PASSWORD, 'project.task', 'search_read',
                [[('id', '=', 0)]], {'fields': [candidate], 'limit': 1})
            start_field = candidate
            break
        except Exception:
            pass

    fields_to_fetch = ['id', 'name', 'project_id', 'date_deadline', 'state', 'stage_id']
    if start_field:
        fields_to_fetch.append(start_field)

    tasks = models.execute_kw(
    DB, uid, PASSWORD, 'project.task', 'search_read',
    [[
        ('project_id', 'in', project_ids),
        ('date_deadline', '!=', False)
    ]],
    {'fields': fields_to_fetch}
)

    all_stage_ids = list({t['stage_id'][0] for t in tasks if t.get('stage_id')})
    closed_stages = set()
    if all_stage_ids:
        stages = models.execute_kw(DB, uid, PASSWORD, 'project.task.type', 'read',
            [all_stage_ids], {'fields': ['id', 'name']})
        name_done = {'done', 'terminé', 'terminée', 'fini', 'finie', 'closed', 'annulé', 'cancelled'}
        for s in stages:
            if s.get('is_closed') or s.get('name', '').lower().strip() in name_done:
                closed_stages.add(s['id'])

    for t in tasks:
        # date_deadline
        raw = t['date_deadline']
        if raw:
            t['date_deadline'] = datetime.strptime(raw.split(" ")[0], '%Y-%m-%d').date()

        # date_start : normalisé sous la clé 'date_start' quelle que soit la version Odoo
        raw_start = t.get(start_field) if start_field else None
        try:
            if raw_start:
                parsed = datetime.strptime(raw_start.split(" ")[0], '%Y-%m-%d').date()
                t['date_start'] = min(parsed, t['date_deadline'])  # date_start jamais après deadline
            else:
                t['date_start'] = t['date_deadline']
        except Exception:
            t['date_start'] = t['date_deadline']

                # AJOUTER CES LIGNES après :
        if t['date_start'] > t['date_deadline']:
            t['date_start'] = t['date_deadline']

        state = str(t.get('state') or '').lower()
        stage_id = t['stage_id'][0] if t.get('stage_id') else None
        t['is_done'] = (any(kw in state for kw in ('done', 'cancel', 'termi', 'close'))
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
                color, rank, key = "#156500", 3, "blue"
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
def compute_all_purchase_data(_uid, _models, filter_mode):
    """Pré-calcule purchase_data pour TOUS les projets actifs (non filtrés).
    Mis en cache pour que le filtre projet global ne déclenche pas de recalcul."""
    uid, models = _uid, _models
    projects = load_projects(uid, models, filter_mode)
    po_lines, policy_map, buyer_map, po_name_map = load_purchase_data_all_projects()
    purchase_data = {p["id"]: get_purchase_for_project(p, po_lines, policy_map, buyer_map, po_name_map)
                     for p in projects}
    return purchase_data, projects


@st.cache_data(ttl=300)
def load_all_analytics(_uid, _models, filter_mode):
    """Charge tout l'analytique. Prend filter_mode (string hashable) au lieu
    d'une liste de dicts coûteuse à hasher → cache stable entre reruns."""
    uid, models = _uid, _models

    project_list = load_projects_with_closed(uid, models, filter_mode)
    # Exclure les comptes "fourre-tout" (dépannage, vente pure, etc.)
    bad_accs = ["dépannage (liège)", "projets (lig)", "vente pure (lig)"]
    project_list = [p for p in project_list
                    if not (p.get("analytic_account_id")
                            and p["analytic_account_id"][1].lower() in bad_accs)]

    analytic_ids = [p["analytic_account_id"][0] for p in project_list if p.get("analytic_account_id")]
    if not analytic_ids:
        return {}, pd.DataFrame(), 0.0, project_list

    year_now   = date.today().year
    year_start = f"{year_now}-01-01"
    year_end   = f"{year_now}-12-31"
    date_12m   = (date.today().replace(day=1) - timedelta(days=365)).strftime("%Y-%m-%d")

    # ── 1) Lignes analytiques (dépenses : classe 6 + timesheets) ──
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
            # Timesheet : montant négatif = coût
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
            # Odoo BE : facture fourn = négatif → -amt positif ; NC fourn = positif → -amt négatif
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
    inv_by_aid = {}

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

    # ── 3) Factures via account.move ──
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
        d_agg = df_m[df_m["type"] == "dep"].groupby("Mois")["val"].sum().rename("Dépenses")
        r_agg = df_m[df_m["type"] == "rev"].groupby("Mois")["val"].sum().rename("CA")
        months = pd.date_range(start=date_12m, end=date.today().strftime("%Y-%m-%d"), freq="MS")
        df_monthly = (pd.DataFrame({"Mois": months})
                      .merge(d_agg.reset_index(), on="Mois", how="left")
                      .merge(r_agg.reset_index(), on="Mois", how="left")
                      .fillna(0))

    # ── 5) Synthèse par projet ──
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

    # ── 6) Marge pondérée projets clôturés ──
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

    return summary, df_monthly, marge_pond, project_list


# ============================================================
# GANTT
# ============================================================

# Ordre voulu pour le tri "Par étape" du Gantt.
# Les libellés Odoo sont normalisés en minuscules avant comparaison,
# donc l'ordre est insensible à la casse et aux accents partiels.
STAGE_ORDER = ["kick-off", "technique / étude", "approvisionnement", "atelier",
               "livraison et montage", "récepton et ce", "facture finale"]

def _stage_rank(stage_name):
    """Renvoie le rang de l'étape selon STAGE_ORDER ; inconnues placées à la fin."""
    s = (stage_name or "").lower().strip()
    try:
        return STAGE_ORDER.index(s)
    except ValueError:
        return len(STAGE_ORDER)


COLOR_ORDER = ["Soudure", "Peinture", "Assemblage", "Câblage", "Test",
               "Montage", "Mise en service", "Réception", "Transport", "Étude", "Autres"]

COLOR_MAP = {
    "Soudure": "#1E88E5", "Peinture": "#FDD835", "Assemblage": "#43A047",
    "Câblage": "#8E24AA", "Test": "#FB8C00", "Montage": "#E53935",
    "Mise en service": "#EC407A", "Réception": "#6D4C41",
    "Transport": "#00ACC1", "Étude": "#34ebc6", "Autres": "#9E9E9E"
}

# Couleurs assombries (opacity ~60%) pour les tâches terminées
COLOR_MAP_DONE = {
    "Soudure": "#0d3a6e", "Peinture": "#8a7a00", "Assemblage": "#1a4a1e",
    "Câblage": "#3d0a5a", "Test": "#7a4400", "Montage": "#6b0f0f",
    "Mise en service": "#7a1040", "Réception": "#2e1f18",
    "Transport": "#004a52", "Étude": "#0f5a4e", "Autres": "#3a3a3a"
}


def classify_task_type(name):
    n = name.lower()
    if "soud" in n: return "Soudure"
    if "peint" in n: return "Peinture"
    if "assembl" in n: return "Assemblage"
    if "cabl" in n or "câbl" in n: return "Câblage"
    if "test" in n: return "Test"
    if "montage" in n or "installation" in n: return "Montage"
    if "mise en service" in n or " mes" in n: return "Mise en service"
    if "recept" in n or "réception" in n or "assistance" in n: return "Réception"
    if "transport" in n: return "Transport"
    if "etude" in n or "étude" in n or "conception" in n or "plan" in n or "calcul" in n: return "Étude"
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
# MODE D'AFFICHAGE (toggle bas) + MODE 2 "ÉCRAN"
# ============================================================

# Hauteur réservée au header + footer en mode 2 (px).
# Si les cadres dépassent en bas : augmente. S'il reste du vide : diminue.
MODE2_OFFSET_PX = 56

# Types de tâches qui font "entrer" un projet Engineering dans le Gantt atelier
WORKSHOP_TYPES = {"Soudure", "Peinture", "Câblage", "Assemblage", "Test"}

# Paramètres du mode 2 (employés, fournisseurs, nb semaines), partagés par tous
# les écrans. ⚠ Sur Streamlit Cloud, ce fichier est remis à zéro à chaque
# redéploiement / reboot de l'app : il faut alors refaire les réglages.
MODE2_SETTINGS_FILE = Path(__file__).with_name("mode2_settings.json")

# Mots-clés utilisés tant qu'aucun fournisseur n'a été choisi dans les paramètres
DEFAULT_SUPPLIER_KEYWORDS = ["xometry", "kerschgens", "kershgens", "lasersteel",
                             "cerfontaine", "mottard", "abus"]

JOURS_FR = ["Lun", "Mar", "Mer", "Jeu", "Ven", "Sam", "Dim"]


def render_display_mode_toggle():
    """Toggle fixé en bas de page. Renvoie True si mode 2.
    L'état est aussi mis dans l'URL (?mode=2) pour survivre à un F5."""
    if "display_mode_2" not in st.session_state:
        st.session_state["display_mode_2"] = st.query_params.get("mode") == "2"

    st.markdown("""<style>
    .st-key-display_mode_toggle{
        /* right:160px → à gauche du bouton "Manage app" de Streamlit Cloud */
        position:fixed; right:160px; bottom:3px; z-index:10001;
        width:auto!important; background:transparent; padding:0 10px;
    }
    .st-key-display_mode_toggle label p{font-size:12px!important;color:#111!important;}
    </style>""", unsafe_allow_html=True)

    with st.container(key="display_mode_toggle"):
        mode2 = st.toggle("Mode écran", key="display_mode_2")

    if mode2:
        st.query_params["mode"] = "2"
    elif "mode" in st.query_params:
        del st.query_params["mode"]
    return mode2


def render_footer():
    st.markdown("""
    <style>
    .footer {
        position: fixed; left: 0; bottom: 0; width: 100%;
        background-color: rgba(240,240,240,0.85); color: #333;
        text-align: center; padding: 6px 0; font-size: 14px;
        border-top: 1px solid #ccc; z-index: 9999;
    }
    </style>
    <div class="footer">Flow - Powered by Olsen-Engineering</div>
    """, unsafe_allow_html=True)


# ---------- Paramètres mode 2 (fichier JSON) ----------

def load_mode2_settings():
    try:
        data = json.loads(MODE2_SETTINGS_FILE.read_text(encoding="utf-8"))
    except Exception:
        data = {}
    data.setdefault("atelier_user_ids", [])
    data.setdefault("montage_user_ids", [])
    data.setdefault("supplier_ids", None)      # None = auto via DEFAULT_SUPPLIER_KEYWORDS
    data.setdefault("gantt_weeks", 4)
    data.setdefault("show_atelier", True)
    data.setdefault("show_montage", True)
    return data


def save_mode2_settings(data):
    try:
        MODE2_SETTINGS_FILE.write_text(json.dumps(data, indent=2), encoding="utf-8")
    except Exception as e:
        st.warning(f"Impossible d'enregistrer les paramètres : {e}")


# ---------- Loaders mode 2 ----------

@st.cache_data(ttl=3600)
def load_internal_users(_uid, _models):
    users = _models.execute_kw(DB, _uid, PASSWORD, "res.users", "search_read",
        [[("share", "=", False), ("active", "=", True)]],
        {"fields": ["id", "name"], "order": "name"})
    return [(u["id"], u["name"]) for u in users]


@st.cache_data(ttl=3600)
def load_suppliers(_uid, _models):
    parts = _models.execute_kw(DB, _uid, PASSWORD, "res.partner", "search_read",
        [[("supplier_rank", ">", 0), ("is_company", "=", True)]],
        {"fields": ["id", "name"], "order": "name"})
    return [(p["id"], p["name"]) for p in parts]


def _norm(s):
    return re.sub(r"[^a-z0-9]", "", (s or "").lower())


def resolve_supplier_ids(settings, suppliers):
    """IDs fournisseurs choisis, ou détection auto par mots-clés si jamais réglé."""
    if settings.get("supplier_ids") is not None:
        return list(settings["supplier_ids"])
    ids = []
    for pid, name in suppliers:
        full = _norm(name)
        tokens = {_norm(t) for t in re.split(r"[\s\-_.,/&]+", name or "")}
        for kw in DEFAULT_SUPPLIER_KEYWORDS:
            # mots courts ("abus") : mot entier ; mots longs : contenu dans le nom
            if (len(kw) <= 5 and kw in tokens) or (len(kw) > 5 and kw in full):
                ids.append(pid)
                break
    return ids


@st.cache_data(ttl=3600)
def _task_start_field(_uid, _models):
    for candidate in ("planned_date_start", "planned_date_begin"):
        try:
            _models.execute_kw(DB, _uid, PASSWORD, "project.task", "search_read",
                [[("id", "=", 0)]], {"fields": [candidate], "limit": 1})
            return candidate
        except Exception:
            pass
    return None


def _to_date(raw):
    if not raw:
        return None
    try:
        return datetime.strptime(str(raw).split(" ")[0], "%Y-%m-%d").date()
    except Exception:
        return None


@st.cache_data(ttl=300)
def load_week_tasks_for_users(_uid, _models, user_ids, monday):
    """Tâches assignées aux utilisateurs donnés qui touchent la semaine (lun→ven)."""
    if not user_ids:
        return []
    friday = monday + timedelta(days=4)
    start_field = _task_start_field(_uid, _models)
    fields = ["id", "name", "project_id", "date_deadline", "state", "user_ids"]
    if start_field:
        fields.append(start_field)
    tasks = _models.execute_kw(DB, _uid, PASSWORD, "project.task", "search_read",
        [[("user_ids", "in", list(user_ids)),
          ("date_deadline", ">=", monday.strftime("%Y-%m-%d"))]],
        {"fields": fields})
    out = []
    for t in tasks:
        dl = _to_date(t.get("date_deadline"))
        if not dl:
            continue
        ds = _to_date(t.get(start_field)) if start_field else None
        ds = min(ds or dl, dl)
        if ds > friday:
            continue
        state = str(t.get("state") or "").lower()
        out.append({
            "name": t["name"],
            "project": t["project_id"][1] if t.get("project_id") else "",
            "user_ids": t.get("user_ids") or [],
            "date_start": ds, "date_deadline": dl,
            "is_done": any(k in state for k in ("done", "cancel", "termi", "close")),
        })
    return out


@st.cache_data(ttl=300)
def load_incoming_po_lines(_uid, _models, supplier_ids):
    """Lignes d'achat confirmées, pas encore totalement reçues.
    supplier_ids = None → tous les fournisseurs ; tuple → uniquement ceux-là."""
    domain = [("order_id.state", "in", ["purchase", "done"]), ("product_qty", ">", 0)]
    if supplier_ids is not None:
        if not supplier_ids:
            return []
        domain.append(("partner_id", "child_of", list(supplier_ids)))
    lines = _models.execute_kw(DB, _uid, PASSWORD, "purchase.order.line", "search_read",
        [domain],
        {"fields": ["name", "product_qty", "qty_received", "date_planned",
                    "partner_id", "order_id", "analytic_distribution"]})
    return [l for l in lines if (l.get("qty_received") or 0) < l["product_qty"]]


def load_workshop_projects(uid, models, monday, weeks):
    """Projets Engineering ayant au moins une tâche Soudure/Peinture/Câblage/
    Assemblage/Test dans la fenêtre affichée. Renvoie (projets, toutes leurs tâches)."""
    end = monday + timedelta(weeks=weeks)
    projects = load_projects(uid, models, "engineering")
    pids = tuple(sorted(p["id"] for p in projects))
    all_tasks = get_tasks(uid, models, pids, monday, end)

    first_ws = {}   # pid -> 1re date de tâche atelier dans la fenêtre (pour le tri)
    for t in all_tasks:
        if classify_task_type(t["name"]) not in WORKSHOP_TYPES:
            continue
        if t["date_deadline"] < monday or t["date_start"] > end:
            continue
        pid = t["project_id"][0]
        first_ws[pid] = min(first_ws.get(pid, t["date_start"]), t["date_start"])

    sel = sorted([p for p in projects if p["id"] in first_ws],
                 key=lambda p: (first_ws[p["id"]], extract_project_code(p["display_name"])))
    sel_ids = {p["id"] for p in sel}
    tasks = [t for t in all_tasks if t["project_id"][0] in sel_ids]
    return sel, tasks


# ---------- Popup paramètres ----------

@st.dialog("Paramètres de l'affichage écran", width="large")
def mode2_settings_dialog(uid, models):
    s = load_mode2_settings()
    users = load_internal_users(uid, models)
    user_names = dict(users)
    user_ids = [i for i, _ in users]

    st.markdown("**Planning de la semaine**")
    t1, t2 = st.columns(2)
    show_atelier = t1.toggle("Afficher Atelier", value=s["show_atelier"])
    show_montage = t2.toggle("Afficher Montage", value=s["show_montage"])
    atelier = st.multiselect("Employés Atelier", user_ids,
        default=[i for i in s["atelier_user_ids"] if i in user_names],
        format_func=lambda i: user_names.get(i, str(i)))
    montage = st.multiselect("Employés Montage", user_ids,
        default=[i for i in s["montage_user_ids"] if i in user_names],
        format_func=lambda i: user_names.get(i, str(i)))

    st.markdown("**Réceptions** (projets Engineering en cours)")
    suppliers = load_suppliers(uid, models)
    sup_names = dict(suppliers)
    fournisseurs = st.multiselect("Fournisseurs suivis", [i for i, _ in suppliers],
        default=[i for i in resolve_supplier_ids(s, suppliers) if i in sup_names],
        format_func=lambda i: sup_names.get(i, str(i)))

    c1, c2 = st.columns(2)
    if c1.button("Enregistrer", type="primary", use_container_width=True):
        s["show_atelier"] = show_atelier
        s["show_montage"] = show_montage
        s["atelier_user_ids"] = atelier
        s["montage_user_ids"] = montage
        s["supplier_ids"] = fournisseurs
        save_mode2_settings(s)
        st.rerun()
    if c2.button("Annuler", use_container_width=True):
        st.rerun()


# ---------- Rendu mode 2 ----------

def _text_color_for(bg):
    h = bg.lstrip("#")
    r, g, b = (int(h[i:i + 2], 16) for i in (0, 2, 4))
    return "#111" if (0.299 * r + 0.587 * g + 0.114 * b) > 150 else "#fff"


def _esc(s):
    # html.escape + "$" neutralisé (sinon st.markdown peut l'interpréter en LaTeX)
    # + retours à la ligne encodés : une ligne vide dans le HTML fait sortir
    #   st.markdown du mode HTML et la suite s'affiche en texte brut.
    s = html.escape(str(s or ""), quote=True).replace("$", "&#36;")
    return s.replace("\r", "").replace("\n", "&#10;")


MODE2_CSS = """<style>
.m2-title{font-size:15px;font-weight:700;margin:0 0 6px 0;color:#eee;}
.m2-title span{font-weight:400;color:#999;font-size:13px;margin-left:8px;}
.wp{font-size:12px;}
.wp-grid{display:grid;row-gap:3px;}
.wp-head{position:sticky;top:0;background:#0e1117;z-index:3;padding:2px 0 4px;border-bottom:1px solid #444;}
.wp-day{text-align:center;color:#aaa;font-weight:600;}
.wp-day.wp-today{color:#fff;background:rgba(255,255,255,.10);border-radius:4px;}
.wp-group{margin:8px 0 2px;font-weight:700;letter-spacing:.08em;color:#8ab4f8;font-size:11px;
  text-transform:uppercase;border-bottom:1px solid #333;padding-bottom:2px;}
.wp-user{border-bottom:1px solid #222;padding:2px 0;}
.wp-name{align-self:center;color:#ddd;padding-right:6px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
.wp-cell{border-left:1px dashed rgba(255,255,255,.10);min-height:34px;}
.wp-cell.wp-today{background:rgba(255,255,255,.05);}
.wp-task{margin:0 2px;border-radius:4px;padding:2px 6px;overflow:hidden;z-index:1;line-height:1.25;}
.wp-t{font-weight:600;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
.wp-d{font-size:10.5px;opacity:.85;white-space:nowrap;}
.wp-empty{color:#777;font-style:italic;padding:4px 0;}
.rc-row{display:grid;grid-template-columns:62px 1fr;column-gap:10px;padding:6px 4px;border-bottom:1px solid #262626;font-size:12.5px;}
.rc-date{font-weight:700;text-align:center;border-radius:4px;padding:2px 0;line-height:1.2;}
.rc-date small{display:block;font-weight:400;font-size:10.5px;opacity:.8;}
.rc-late{background:#b71c1c;color:#fff;}
.rc-today{background:#FFA000;color:#111;}
.rc-soon{background:rgba(255,255,255,.08);color:#ddd;}
.rc-sup{font-weight:700;color:#eee;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
.rc-info{color:#aaa;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
.rc-sep{font-size:11px;color:#8ab4f8;text-transform:uppercase;letter-spacing:.08em;margin:8px 0 2px;}
</style>"""


def render_header_mode2(uid, models):
    """Mode écran plein écran : pas de header (ni logo, ni titre, ni barre Streamlit).
    Le bouton ⚙️ est placé dans le footer, à gauche du toggle."""
    st.markdown("""<style>
    header[data-testid="stHeader"]{display:none!important;}
    [data-testid="stMainBlockContainer"], .block-container{
        padding-top:0.6rem!important; padding-left:1rem!important; padding-right:1rem!important;
        max-width:100%!important;
    }
    /* Les blocs invisibles (CSS, autorefresh, toggle, ⚙️) prennent chacun ~16 px d'espace :
       on les sort du flux pour qu'ils ne poussent plus le contenu vers le bas. */
    .element-container:has(style), .stElementContainer:has(style),
    .element-container:has(iframe[title*="autorefresh"]), .stElementContainer:has(iframe[title*="autorefresh"]),
    [data-testid="stLayoutWrapper"]:has(> .st-key-display_mode_toggle),
    [data-testid="stLayoutWrapper"]:has(> .st-key-m2_settings_box){
        position:absolute!important; height:0!important; overflow:visible;
    }
    .st-key-m2_settings_box{
        position:fixed; right:290px; bottom:2px; z-index:10001; width:auto!important;
    }
    .st-key-m2_settings_box button{
        background:transparent!important; border:none!important; color:#111!important;
        min-height:0!important; padding:2px 8px!important; font-size:13px!important;
    }
    .st-key-m2_settings_box button:hover{background:rgba(0,0,0,.08)!important;}
    .st-key-m2_settings_box button p{color:#111!important;font-size:13px!important;}
    </style>""", unsafe_allow_html=True)

    with st.container(key="m2_settings_box"):
        if st.button("⚙️ Paramètres", key="m2_settings_btn", help="Paramètres affichage écran"):
            mode2_settings_dialog(uid, models)


def build_week_planning_html(groups, tasks, monday, today):
    days = [monday + timedelta(days=i) for i in range(5)]
    cols = "130px repeat(5, minmax(0,1fr))"
    out = ["<div class='wp'>"]
    head = "".join(
        f"<div class='wp-day{' wp-today' if d == today else ''}'>{JOURS_FR[i]} {d:%d/%m}</div>"
        for i, d in enumerate(days))
    out.append(f"<div class='wp-grid wp-head' style='grid-template-columns:{cols}'><div></div>{head}</div>")

    for gname, users in groups:
        out.append(f"<div class='wp-group'>{_esc(gname)}</div>")
        if not users:
            out.append("<div class='wp-empty'>Aucun employé sélectionné (⚙️ en bas de l'écran)</div>")
            continue
        for user_id, uname in users:
            items = []
            for t in tasks:
                if user_id not in t["user_ids"]:
                    continue
                s_idx = max(0, (t["date_start"] - monday).days)
                e_idx = min(4, (t["date_deadline"] - monday).days)
                if e_idx < 0 or s_idx > 4 or s_idx > e_idx:
                    continue
                items.append((s_idx, e_idx, t))
            items.sort(key=lambda x: (x[0], x[1]))

            # Répartition en "couloirs" pour que les tâches qui se chevauchent s'empilent
            lanes_end, placed = [], []
            for s_idx, e_idx, t in items:
                lane = next((i for i, le in enumerate(lanes_end) if s_idx > le), None)
                if lane is None:
                    lanes_end.append(e_idx)
                    lane = len(lanes_end) - 1
                else:
                    lanes_end[lane] = e_idx
                placed.append((lane, s_idx, e_idx, t))
            n = max(1, len(lanes_end))

            cells = [f"<div class='wp-name' style='grid-row:1/span {n};grid-column:1'>{_esc(uname)}</div>"]
            for li in range(n):
                for di, d in enumerate(days):
                    cells.append(f"<div class='wp-cell{' wp-today' if d == today else ''}' "
                                 f"style='grid-row:{li + 1};grid-column:{di + 2}'></div>")
            for lane, s_idx, e_idx, t in placed:
                ttype = classify_task_type(t["name"])
                bg = COLOR_MAP_DONE[ttype] if t["is_done"] else COLOR_MAP[ttype]
                fg = _text_color_for(bg)
                tip = f"{t['name']} — {t['project']}"
                cells.append(
                    f"<div class='wp-task' title='{_esc(tip)}' style='grid-row:{lane + 1};"
                    f"grid-column:{s_idx + 2}/{e_idx + 3};background:{bg};color:{fg}'>"
                    f"<div class='wp-t'>{_esc(t['name'])}</div>"
                    f"<div class='wp-d'>{t['date_start']:%d/%m} → {t['date_deadline']:%d/%m}</div></div>")
            out.append(f"<div class='wp-grid wp-user' style='grid-template-columns:{cols}'>{''.join(cells)}</div>")
    out.append("</div>")
    return "".join(out)


def render_zone_planning_semaine(uid, models, settings):
    today = date.today()
    monday = today - timedelta(days=today.weekday())
    friday = monday + timedelta(days=4)
    st.markdown(f"<div class='m2-title'>Planning semaine {monday.isocalendar()[1]}"
                f"<span>{monday:%d/%m} → {friday:%d/%m}</span></div>", unsafe_allow_html=True)

    users = dict(load_internal_users(uid, models))
    groups = []
    if settings.get("show_atelier", True):
        groups.append(("Atelier", [(i, users[i]) for i in settings["atelier_user_ids"] if i in users]))
    if settings.get("show_montage", True):
        groups.append(("Montage", [(i, users[i]) for i in settings["montage_user_ids"] if i in users]))
    if not groups:
        st.info("Atelier et Montage sont masqués (⚙️ en bas de l'écran).")
        return
    all_ids = tuple(sorted({i for _, g in groups for i, _ in g}))
    tasks = load_week_tasks_for_users(uid, models, all_ids, monday)

    st.markdown(build_week_planning_html(groups, tasks, monday, today), unsafe_allow_html=True)


def render_zone_gantt_atelier(uid, models, settings, projects, tasks, monday, weeks):
    ct, cs = st.columns([4, 1])
    with ct:
        st.markdown(f"<div class='m2-title'>Gantt atelier — Engineering"
                    f"<span>{len(projects)} projets · {weeks} semaines</span></div>",
                    unsafe_allow_html=True)
    with cs:
        new_weeks = st.slider("Semaines", 2, 8, value=weeks, key="m2_gantt_weeks",
                              label_visibility="collapsed")
        if new_weeks != weeks:
            settings["gantt_weeks"] = new_weeks
            save_mode2_settings(settings)
            st.rerun()

    if not projects:
        st.info("Aucun projet Engineering avec soudure / peinture / câblage / assemblage / test "
                "sur la période.")
        return

    end = monday + timedelta(weeks=weeks)

    # Libellés uniques (évite que 2 projets au libellé identique fusionnent sur 1 ligne)
    labels, seen = {}, {}
    for p in projects:
        lbl = project_label(p)
        seen[lbl] = seen.get(lbl, 0) + 1
        labels[p["id"]] = lbl if seen[lbl] == 1 else f"{lbl} ({seen[lbl]})"
    order = [labels[p["id"]] for p in projects]

    rows = []
    for t in tasks:
        if t["date_deadline"] < monday or t["date_start"] > end:
            continue
        ttype = classify_task_type(t["name"])
        rows.append({
            "Tâche": t["name"],
            "Projet": labels[t["project_id"][0]],
            "Début": pd.Timestamp(t["date_start"]),
            "Fin": pd.Timestamp(t["date_deadline"]) + pd.Timedelta(days=1),   # fin incluse
            "Type": ttype,
            "Légende": ttype + "__done" if t.get("is_done") else ttype,
            "Échéance": t["date_deadline"].strftime("%d/%m/%Y"),
        })
    if not rows:
        st.info("Aucune tâche sur la période.")
        return

    df = pd.DataFrame(rows)
    full_color_map = {**COLOR_MAP, **{k + "__done": v for k, v in COLOR_MAP_DONE.items()}}
    fig = px.timeline(df, x_start="Début", x_end="Fin", y="Projet", color="Légende",
                      color_discrete_map=full_color_map, hover_name="Tâche",
                      hover_data={"Début": False, "Fin": False, "Projet": True,
                                  "Légende": False, "Type": True, "Échéance": True})
    for trace in fig.data:
        if trace.name.endswith("__done"):
            trace.showlegend = False
            trace.name = trace.name.replace("__done", "")

    fig.update_layout(
        barmode="overlay", height=max(260, len(order) * 24 + 70),
        margin=dict(l=10, r=10, t=30, b=10), plot_bgcolor="rgba(0,0,0,0)",
        yaxis=dict(categoryorder="array", categoryarray=list(reversed(order)),
                   title_text="", tickfont=dict(size=11),
                   showgrid=True, gridcolor="rgba(180,180,180,0.15)"),
        xaxis=dict(title_text="", showgrid=False, range=[monday, end],
                   dtick=7 * 24 * 3600 * 1000, tick0=monday.strftime("%Y-%m-%d"),
                   tickformat="S%V<br>%d/%m"),
        legend=dict(orientation="h", yanchor="bottom", y=1.0, xanchor="center", x=0.5,
                    font=dict(size=10), title_text=""),
    )
    today = date.today()
    fig.add_vline(x=today, line_width=2, line_color="white", opacity=0.9)
    d = monday
    while d < end:
        fig.add_vrect(x0=d + timedelta(days=5), x1=d + timedelta(days=7),
                      fillcolor="rgba(255,255,255,0.04)", layer="below", line_width=0)
        d += timedelta(days=7)

    st.plotly_chart(fig, use_container_width=True,
                    config={"displaylogo": False, "displayModeBar": False})


def render_zone_receptions(uid, models, settings, projects):
    st.markdown("<div class='m2-title'>Réceptions à venir"
                "<span>projets Engineering en cours · fournisseurs suivis</span></div>",
                unsafe_allow_html=True)

    suppliers = load_suppliers(uid, models)
    sup_ids = tuple(sorted(resolve_supplier_ids(settings, suppliers)))
    if not sup_ids:
        st.info("Aucun fournisseur sélectionné (⚙️ en bas de l'écran).")
        return
    if not projects:
        st.info("Aucun projet Engineering en cours.")
        return

    # compte analytique -> projet
    aid_to_proj = {p["analytic_account_id"][0]: p for p in projects if p.get("analytic_account_id")}
    lines = load_incoming_po_lines(uid, models, sup_ids)

    groups = {}   # (date, fournisseur, PO, projet) -> [descriptions]
    for l in lines:
        proj = None
        for key in (l.get("analytic_distribution") or {}):
            for part in str(key).split(","):          # clés "12,34" possibles (Odoo 17+)
                if part.strip().isdigit() and int(part) in aid_to_proj:
                    proj = aid_to_proj[int(part)]
                    break
            if proj:
                break
        if not proj:
            continue
        dp = _to_date(l.get("date_planned"))
        sup = (l["partner_id"][1] if l.get("partner_id") else "?").split(", ")[0]
        po = l["order_id"][1] if l.get("order_id") else ""
        groups.setdefault((dp or date.max, sup, po, proj["id"]), []).append(l["name"] or "")

    if not groups:
        st.info("Aucune réception en attente pour ces projets.")
        return

    today = date.today()
    proj_by_id = {p["id"]: p for p in projects}
    out, last_sep = [], None
    for (dp, sup, po, pid), descs in sorted(groups.items(), key=lambda kv: (kv[0][0], kv[0][1])):
        sep = "Sans date" if dp == date.max else ("En retard" if dp < today else "À venir")
        if sep != last_sep:
            out.append(f"<div class='rc-sep'>{sep}</div>")
            last_sep = sep

        if dp == date.max:
            cls, dtxt = "rc-soon", "—"
        else:
            cls = "rc-late" if dp < today else ("rc-today" if dp == today else "rc-soon")
            dtxt = f"{dp:%d/%m}<small>{JOURS_FR[dp.weekday()]}</small>"
        n = len(descs)
        art = short_desc(descs[0].split("\n")[0], 40) if n == 1 else f"{n} articles"
        info = f"{po} · {project_label(proj_by_id[pid])} · {art}"
        out.append(f"<div class='rc-row' title='{_esc(chr(10).join(descs))}'>"
                   f"<div class='rc-date {cls}'>{dtxt}</div>"
                   f"<div><div class='rc-sup'>{_esc(sup)}</div>"
                   f"<div class='rc-info'>{_esc(info)}</div></div></div>")
    st.markdown("".join(out), unsafe_allow_html=True)


def render_mode2_layout(uid, models):
    """70 % : 2 lignes identiques empilées | 30 % : 1 cadre pleine hauteur."""
    h_side = f"calc(100vh - {MODE2_OFFSET_PX}px)"
    h_row  = f"calc(50vh - {MODE2_OFFSET_PX / 2 + 8}px)"   # 8 px = moitié de l'espace entre les 2 lignes
    st.markdown(f"""<style>
    .block-container {{ padding-bottom: 2.5rem !important; }}
    .st-key-m2_top, .st-key-m2_bottom, .st-key-m2_side {{
        border: 1px solid rgba(250,250,250,0.2); border-radius: 8px;
        padding: 12px; box-sizing: border-box;
        overflow-y: auto !important; flex: 0 0 auto !important;
        justify-content: flex-start;
    }}
    .st-key-m2_top, .st-key-m2_bottom {{ height: {h_row} !important; }}
    .st-key-m2_side {{ height: {h_side} !important; }}
    </style>""", unsafe_allow_html=True)
    st.markdown(MODE2_CSS, unsafe_allow_html=True)

    settings = load_mode2_settings()
    today = date.today()
    monday = today - timedelta(days=today.weekday())
    weeks = int(settings.get("gantt_weeks") or 4)

    ws_projects, ws_tasks, ws_error = [], [], None
    try:
        ws_projects, ws_tasks = load_workshop_projects(uid, models, monday, weeks)
    except Exception as e:
        ws_error = e

    left, right = st.columns([7, 3], gap="medium")
    with left:
        with st.container(key="m2_top"):
            try:
                render_zone_planning_semaine(uid, models, settings)
            except Exception as e:
                st.error(f"Planning semaine : {e}")
        with st.container(key="m2_bottom"):
            if ws_error:
                st.error(f"Gantt atelier : {ws_error}")
            else:
                try:
                    render_zone_gantt_atelier(uid, models, settings, ws_projects, ws_tasks, monday, weeks)
                except Exception as e:
                    st.error(f"Gantt atelier : {e}")
    with right:
        with st.container(key="m2_side"):
            try:
                # Tous les projets Engineering en cours (pas seulement ceux du Gantt atelier)
                eng_projects = load_projects(uid, models, "engineering")
                render_zone_receptions(uid, models, settings, eng_projects)
            except Exception as e:
                st.error(f"Réceptions : {e}")


# ============================================================
# MAIN APP
# ============================================================

def main():
    st.set_page_config(page_title="Dashboard", page_icon="🏗️", layout="wide")
    st.markdown("""<style>
    .block-container{padding-top:0.5rem!important;}
    div[data-testid="stToggle"]>label{font-size:13px!important;}
    </style>""", unsafe_allow_html=True)

    try:
        # IMPORTANT : appeler _load_credentials() à CHAQUE rerun, indépendamment
        # du cache. Sinon, si Streamlit redémarre le module Python, les globales
        # DB/USERNAME/PASSWORD reviennent à None mais connect_odoo (cached_resource)
        # renvoie sa valeur cachée sans rappeler _load_credentials, et tous les
        # appels Odoo plantent ensuite avec "cannot marshal None".
        _load_credentials()
        uid, models = connect_odoo()
    except Exception as e:
        st.error(f"Connexion Odoo impossible : {e}")
        return

    st_autorefresh(interval=600000, key="refresh_10min")

    # Toggle mode d'affichage (bas-droite) + footer, rendus pour les 2 modes
    mode2 = render_display_mode_toggle()
    render_footer()

    if mode2:
        render_header_mode2(uid, models)
        render_mode2_layout(uid, models)
        return

    # ===================== MODE 1 (inchangé) =====================

    for k, v in [("months", 3), ("selected_purchase_project_id", None),
                 ("filter_engineering", True), ("filter_standard", False),
                 ("global_project_filter", None),
                 ("global_project_selectbox_nonce", 0)]:
        if k not in st.session_state:
            st.session_state[k] = v

    # Bannière
    c1, c2, c3 = st.columns([1, 4, 1.6])
    with c1:
        st.image("https://upload.wikimedia.org/wikipedia/commons/b/ba/Olsen-Logo.png", width=180)
        st.markdown("<div style='color:green;font-weight:bold;margin-top:20px;'>Connecté Odoo</div>",
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

        # Filtre projet global (sous les toggles, même largeur).
        # Périmètre = projets ACTIFS (non clôturés/non "fait"), même que Gantt/Purchases.
        _all_projects = load_projects(uid, models, fm)
        _global_options = sorted(
            [(p["id"], project_label(p)) for p in _all_projects],
            key=lambda t: t[1].lower()
        )
        _opt_labels = [lbl for _, lbl in _global_options]
        _label_to_id = {lbl: pid for pid, lbl in _global_options}

        _cur_id = st.session_state.get("global_project_filter")
        _cur_idx = None
        if _cur_id is not None:
            for i, (pid, _) in enumerate(_global_options):
                if pid == _cur_id:
                    _cur_idx = i
                    break

        # Selectbox + croix à droite (croix visible seulement si filtre actif)
        _has_filter = st.session_state.get("global_project_filter") is not None
        # Key dynamique : on incrémente le nonce au reset pour forcer le widget
        # à se vider visuellement (Streamlit ne reset pas toujours via pop).
        _sel_key = f"global_project_selectbox_{st.session_state['global_project_selectbox_nonce']}"
        _csel, _cclr = st.columns([7, 1])
        with _csel:
            _sel_label = st.selectbox(
                "Filtre projet",
                options=_opt_labels,
                index=_cur_idx,
                placeholder="Filtre projet",
                key=_sel_key,
                label_visibility="collapsed",
            )
        with _cclr:
            if _has_filter:
                if st.button("✕", key="clear_global_filter",
                             help="Effacer le filtre",
                             use_container_width=True):
                    st.session_state["global_project_filter"] = None
                    st.session_state["global_project_selectbox_nonce"] += 1
                    st.rerun()

        _new_id = _label_to_id.get(_sel_label) if _sel_label else None
        if _new_id != st.session_state.get("global_project_filter"):
            st.session_state["global_project_filter"] = _new_id
            st.rerun()

    GLOBAL_PROJECT_ID = st.session_state.get("global_project_filter")

    tab1, tab2, tab3 = st.tabs(["Planning", "Purchases", "Analytique"])

    # ── ONGLET 1 : PLANNING ──────────────────────────────────────
    with tab1:
        # On charge TOUS les projets actifs et leurs tâches sans tenir compte du filtre
        # projet global ici : les @st.cache_data restent valides quel que soit ce filtre,
        # donc bascule filtre/défiltre = instantanée après le premier chargement.
        projects_all_active = load_projects(uid, models, fm)
        months = st.session_state["months"]
        weeks  = build_weeks_horizon(months)
        # Tuple trié → hash stable et identique quel que soit le filtre projet global
        _all_pids = tuple(sorted(p["id"] for p in projects_all_active))
        all_tasks = get_tasks(uid, models, _all_pids, weeks[0][1], weeks[-1][2])

        # Maintenant on applique le filtre projet global au niveau de l'affichage
        if GLOBAL_PROJECT_ID is not None:
            projects = [p for p in projects_all_active if p["id"] == GLOBAL_PROJECT_ID]
        else:
            projects = projects_all_active
        _pids_visibles = {p["id"] for p in projects}
        tasks = [t for t in all_tasks if t["project_id"][0] in _pids_visibles]

        # Titre Gantt + slider mois + toggle "Par étape" sur la même ligne
        _gt1, _gt2, _gt3 = st.columns([4, 1, 1])
        with _gt1:
            st.subheader("Gantt")
        with _gt2:
            new_months = st.slider("Mois", 1, 6, st.session_state["months"],
                                   key="planning_months_slider", label_visibility="collapsed")
            if new_months != st.session_state["months"]:
                st.session_state["months"] = new_months
                st.rerun()
        with _gt3:
            _sort_by_stage = st.toggle("Par étape", value=False, key="gantt_sort_stage")

        today      = date.today()
        start_view = today
        end_view   = today + timedelta(days=30 * months)

        # Libellé d'affichage : orange si pas de date de fin projet dans Odoo.
        # Plotly accepte du HTML dans les ticktext (<span style="color:..">).
        def _proj_display_label(proj):
            base = project_label(proj)
            if proj.get("date_end") is None:
                return f"<span style='color:#FFA000'>{base}</span>"
            return base
        _id_to_display = {p["id"]: _proj_display_label(p) for p in projects}

        gantt_data = []
        for t in tasks:
            proj = next((p for p in projects if p['id'] == t['project_id'][0]), None)
            if not proj:
                continue
            label    = _id_to_display[proj["id"]]
            task_type = classify_task_type(t["name"])
            color = COLOR_MAP_DONE[task_type] if t.get("is_done") else COLOR_MAP[task_type]

            start_dt = t["date_start"]
            end_dt   = t["date_deadline"]
            if start_dt >= end_dt:
                end_dt = start_dt + timedelta(days=1)

            gantt_data.append({
                "Tâche":        t["name"],
                "Projet":       label,
                "Début":        start_dt,
                "Fin":          end_dt,
                "Type":         task_type,
                "is_done":      t.get("is_done", False),
                "deadline_str": str(t["date_deadline"]),
                "color":        color,
            })

        # Ligne fantôme pour les projets sans tâche planifiée
        _labels_avec_tache = {row["Projet"] for row in gantt_data}
        for proj in projects:
            lbl = _id_to_display[proj["id"]]
            if lbl in _labels_avec_tache:
                continue
            gantt_data.append({
                "Tâche":        "(aucune tâche planifiée)",
                "Projet":       lbl,
                "Début":        today,
                "Fin":          today,
                "Type":         "Autres",
                "is_done":      False,
                "deadline_str": "",
                "color":        "rgba(0,0,0,0)",
                "_empty":       True,
            })

        if gantt_data:
            df_gantt = pd.DataFrame(gantt_data)
            if "_empty" not in df_gantt.columns:
                df_gantt["_empty"] = False
            df_gantt["_empty"] = df_gantt["_empty"].fillna(False).astype(bool)

            df_gantt["code"]  = df_gantt["Projet"].apply(extract_project_code)

            # Étape Odoo de chaque projet, pour le tri "Par étape" et les bandes alternées
            _label_to_stage = {_id_to_display[p["id"]]: p.get("stage", "—") for p in projects}
            df_gantt["stage"] = df_gantt["Projet"].map(_label_to_stage).fillna("—")

            if _sort_by_stage:
                # Tri par étape : facture finale en HAUT (étape la plus avancée),
                # puis ..., kick-off, puis projets SANS rang tout en BAS.
                # L'axe Y est inversé : 1re ligne triée = en bas → tri ascendant
                # avec inconnus (rang -1) avant tout connu (rang 0..N).
                _UNKNOWN = len(STAGE_ORDER)
                def _rank_for_sort(stage_name):
                    r = _stage_rank(stage_name)
                    return -1 if r >= _UNKNOWN else r
                df_gantt["stage_rank"] = df_gantt["stage"].apply(_rank_for_sort)
                df_gantt = df_gantt.sort_values(["stage_rank", "code"])
            else:
                # Tri par date de fin croissante (dates lointaines en haut, axe inversé).
                # Projets sans date_end (oranges) restent tout en bas.
                df_gantt["date_end_proj"] = pd.to_datetime(
                    df_gantt["Projet"].map(
                        {_id_to_display[p["id"]]: p.get("date_end") for p in projects}
                    ), errors="coerce")
                df_gantt = df_gantt.sort_values(
                    ["date_end_proj", "code"],
                    ascending=[True, True],
                    na_position="first",
                )
            df_gantt["Projet_display"] = df_gantt["Projet"]

            df_gantt["Légende"] = df_gantt.apply(
                lambda r: r["Type"] + "__done" if r["is_done"] else r["Type"], axis=1)
            full_color_map = {**COLOR_MAP, **{k + "__done": v for k, v in COLOR_MAP_DONE.items()}}

            df_gantt["Début"] = pd.to_datetime(df_gantt["Début"])
            df_gantt["Fin"]   = pd.to_datetime(df_gantt["Fin"])
            mask = (df_gantt["Fin"] <= df_gantt["Début"]) & (~df_gantt["_empty"])
            df_gantt.loc[mask, "Fin"] = df_gantt.loc[mask, "Début"] + pd.Timedelta(days=1)

            fig = px.timeline(
                df_gantt,
                x_start="Début", x_end="Fin", y="Projet_display",
                color="Légende",
                color_discrete_map=full_color_map,
                hover_name="Tâche",
                hover_data={"Début": True, "Fin": True, "Type": True, "Projet_display": False,
                            "Légende": False, "is_done": False},
            )
            for trace in fig.data:
                if trace.name.endswith("__done"):
                    trace.showlegend = False
                    trace.name = trace.name.replace("__done", "")

            n_proj = len(df_gantt["Projet_display"].unique())
            fig.update_layout(
                barmode="overlay",
                dragmode="pan",
                height=max(500, n_proj * 18 + 140),
                bargap=0.3, bargroupgap=0.1,
                margin=dict(l=20, r=20, t=40, b=20),
                yaxis=dict(categoryorder="array",
                           categoryarray=list(reversed(df_gantt["Projet_display"].unique().tolist())),
                           tickfont=dict(size=12),
                           title_text="",
                           showgrid=True, gridcolor="rgba(180,180,180,0.18)"),
                xaxis=dict(title_text="", showgrid=False),
                plot_bgcolor="rgba(0,0,0,0)",
                legend=dict(orientation="h", yanchor="bottom", y=1.02,
                            xanchor="center", x=0.5, font=dict(size=10))
            )
            fig.update_xaxes(range=[start_view, end_view])
            fig.add_vline(x=today, line_width=2, line_color="white", opacity=0.9)

            cur = date(today.year, today.month, 1)
            while True:
                cur = date(cur.year + 1, 1, 1) if cur.month == 12 else date(cur.year, cur.month + 1, 1)
                if cur > end_view:
                    break
                fig.add_vline(x=cur, line_width=1, line_dash="dot", line_color="rgba(200,200,200,0.35)")

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

            # Bandes de fond alternées par étape (uniquement en mode "Par étape")
            if _sort_by_stage:
                cat_order = list(reversed(df_gantt["Projet_display"].unique().tolist()))
                proj_to_stage = dict(zip(df_gantt["Projet_display"], df_gantt["stage"]))
                blocks = []  # (stage, i_start, i_end)
                for idx, cat in enumerate(cat_order):
                    stg = proj_to_stage.get(cat, "—")
                    if blocks and blocks[-1][0] == stg:
                        blocks[-1] = (stg, blocks[-1][1], idx)
                    else:
                        blocks.append((stg, idx, idx))
                for band_i, (stg, i0, i1) in enumerate(blocks):
                    if band_i % 2 == 1:
                        fig.add_hrect(
                            y0=i0 - 0.5, y1=i1 + 0.5,
                            fillcolor="rgba(255,255,255,0.06)",
                            layer="below", line_width=0,
                        )

            st.plotly_chart(fig, use_container_width=True, config={"displaylogo": False})
        else:
            st.info("Aucun projet à afficher avec ce filtre.")

        st.markdown(f"<div style='font-size:14px;'>Projets affichés : <b>{len(projects)}</b></div>",
                    unsafe_allow_html=True)

        st.subheader("Tâches du projet")
        if GLOBAL_PROJECT_ID is not None and projects:
            # Filtre projet global actif : on affiche directement ses tâches
            tid = projects[0]["id"]
            tlist = sorted([t for t in tasks if t["project_id"][0] == tid],
                           key=lambda x: x["date_deadline"])
            if tlist:
                for t in tlist:
                    wd      = t["date_deadline"].weekday()
                    we_flag = " **[WE]**" if wd >= 5 else ""
                    done    = " (Terminé)" if t.get("is_done") else ""
                    st.write(f"- **{t['name']}**{done}{we_flag} — {t['date_deadline'].strftime('%d-%m-%Y')}")
            else:
                st.info("Aucune tâche pour ce projet.")
        else:
            st.info("Sélectionne un projet dans le filtre en haut à droite pour voir ses tâches.")

    # ── ONGLET 2 : PURCHASES ─────────────────────────────────────
    with tab2:
        st.markdown("### Purchases par projet")

        # Pré-calcul caché pour TOUS les projets actifs : le filtre projet global
        # ne déclenche plus de recalcul, juste un re-filtrage côté Python.
        purchase_data, projects_all = compute_all_purchase_data(uid, models, fm)

        # Filtre projet global : isoler ce projet uniquement (s'il fait partie
        # des projets actifs ; sinon on respecte le filtre habituel).
        if GLOBAL_PROJECT_ID is not None:
            projects_all = [p for p in projects_all if p["id"] == GLOBAL_PROJECT_ID]

        # Tri : rouges (grey>0) d'abord, oranges (orange>0 sans grey) ensuite,
        # puis les autres.
        def _sort_key(p):
            sm, _ = purchase_data[p["id"]]
            if sm["grey"] > 0:
                return (0, -sm["grey"] - sm["orange"])
            if sm["orange"] > 0:
                return (1, -sm["orange"])
            return (2, 0)
        projects_all = sorted(projects_all, key=_sort_key)

        # CSS pour colorer en orange les boutons des vignettes wrappées
        # avec st.container(key="orange_btn_*"). Streamlit injecte la classe
        # `.st-key-orange_btn_<id>` autour du container.
        st.markdown("""
        <style>
        [class*="st-key-orange_btn_"] button {
            background-color: #FFA000 !important;
            color: white !important;
            border: 1px solid #FFA000 !important;
        }
        [class*="st-key-orange_btn_"] button:hover {
            background-color: #FF8C00 !important;
            border-color: #FF8C00 !important;
        }
        </style>
        """, unsafe_allow_html=True)

        for i in range(0, len(projects_all), 6):
            cols = st.columns(6)
            for col, p in zip(cols, projects_all[i:i+6]):
                with col:
                    sm, _ = purchase_data[p['id']]
                    tot   = max(sm["total"], 1)
                    is_red    = sm["grey"] > 0
                    is_orange = (not is_red) and sm["orange"] > 0
                    tc        = "red" if is_red else "#FFA000" if is_orange else "white"
                    btn_label = (f"{p['company']}\n "
                                 f"{short_desc(clean_description_from_display_name(p['display_name']), 25)}")

                    if is_orange:
                        # Wrapper pour appliquer le CSS orange via la classe st-key-*
                        with st.container(key=f"orange_btn_{p['id']}"):
                            clicked = st.button(btn_label, key=f"proj_btn_{p['id']}")
                    else:
                        clicked = st.button(
                            btn_label,
                            key=f"proj_btn_{p['id']}",
                            type="primary" if is_red else "secondary",
                        )
                    if clicked:
                        st.session_state["selected_purchase_project_id"] = p['id']

                    st.markdown(f"""
                        <div style="width:100%;height:12px;border-radius:6px;overflow:hidden;
                            display:flex;border:1px solid #FFFFFF;margin-top:4px;">
                            <div style="width:{100*sm['grey']//tot}%;background:#757575;"></div>
                            <div style="width:{100*sm['orange']//tot}%;background:#FFA000;"></div>
                            <div style="width:{100*sm['white']//tot}%;background:#FFFFFF;"></div>
                            <div style="width:{100*sm['blue']//tot}%;background:#1565C0;"></div>
                            <div style="width:{100*sm['green']//tot}%;background:#2E7D32;"></div>
                        </div>
                        <div style="text-align:right;font-size:12px;color:{tc};margin-top:2px;">
                            {sm['green']} / {sm['total']} lignes
                        </div>""",
                        unsafe_allow_html=True)

        st.markdown("---")
        st.subheader("Détail lignes d'achat")

        # Si filtre projet global actif → afficher d'office le détail de ce projet
        if GLOBAL_PROJECT_ID is not None and projects_all:
            sel_id = projects_all[0]["id"]
        else:
            sel_id = st.session_state.get("selected_purchase_project_id")

        if sel_id is None:
            st.info("Clique sur une vignette pour voir le détail.")
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
                            <div><b>Reçu:</b> {row['Received']}</div>
                            <div><b>Date:</b> {dd}</div>
                        </div>""", unsafe_allow_html=True)

    # ── ONGLET 3 : ANALYTIQUE ────────────────────────────────────
    with tab3:
        st.markdown("### Bilan analytique")

        with st.spinner("Chargement analytiques..."):
            analytics, df_monthly, marge_pond, projects_ana = load_all_analytics(uid, models, fm)

        if not analytics:
            st.info("Aucune donnée disponible.")
            return

        # Détection du mismatch : si le code Sxx-xxxxx extrait du NOM du compte
        # analytique diffère de celui du projet, c'est qu'on récupère par erreur
        # les chiffres d'un autre projet. On marque la ligne pour ne pas
        # double-compter et afficher des tirets dans le tableau.
        def _is_mismatch(p):
            aa = p.get("analytic_account_id")
            if not aa:
                return False  # pas de compte → traité ailleurs
            code_proj = extract_project_code(p.get("display_name", ""))
            code_acc  = extract_project_code(aa[1] or "")
            # Si un seul des deux codes manque on ne juge pas (pas de doublon créé).
            if not code_proj or not code_acc:
                return False
            return code_proj != code_acc

        mismatch_ids = {p["id"] for p in projects_ana if _is_mismatch(p)}

        # ── Statistiques générales : projets NON clôturés (en cours) ──
        # On exclut les mismatchs (double comptage compte analytique partagé)
        # ET les projets entièrement facturés (a_facturer == 0) : ils restent
        # ouverts pour suivi mais n'ont plus rien à venir.
        actifs = [p for p in projects_ana
                  if not p.get("is_closed")
                  and p["id"] not in mismatch_ids
                  and analytics.get(p["id"])
                  and analytics[p["id"]]["ca_total"] > 0
                  and abs(analytics[p["id"]]["a_facturer"]) > 0.01]

        s_ca_total  = sum(analytics[p["id"]]["ca_total"]   for p in actifs)
        s_a_fac     = sum(analytics[p["id"]]["a_facturer"] for p in actifs)
        ratio_fac   = (s_a_fac / s_ca_total * 100) if s_ca_total > 0 else 0.0

        st.markdown("<div style='font-size:13px;color:#aaa;margin-bottom:6px;'>"
                    "Statistiques générales (projets en cours, tous millésimes)</div>",
                    unsafe_allow_html=True)

        m1, m2, m3 = st.columns(3)
        m1.metric("Ventes en cours", fmt_eur(s_ca_total),
                  help="CA total des projets non clôturés")
        m2.metric("À facturer en cours", fmt_eur(s_a_fac),
                  help="Somme à facturer des projets non clôturés")
        m3.metric("Ratio À facturer / CA", f"{ratio_fac:.1f} %",
                  help="À facturer / CA total, projets non clôturés")

        st.markdown("---")

        # ── Titre + toggle clôturés/en cours sur la même ligne ──
        _dt1, _dt2 = st.columns([4, 1])
        with _dt1:
            st.markdown("#### Détail par projet")
        with _dt2:
            show_closed = st.toggle("Projets clôturés", value=False, key="ana_show_closed")

        # Filtre projet global : appliqué uniquement en mode "en cours".
        # En mode "clôturés", on bypass le filtre global (on montre tous les clôturés).
        if show_closed:
            projects_filtered = [p for p in projects_ana if p.get("is_closed")]
        else:
            projects_filtered = [p for p in projects_ana if not p.get("is_closed")]
            if GLOBAL_PROJECT_ID is not None:
                projects_filtered = [p for p in projects_filtered if p["id"] == GLOBAL_PROJECT_ID]

        rows = []
        for p in projects_filtered:
            a = analytics.get(p["id"])
            if a is None:
                continue
            mm = p["id"] in mismatch_ids
            rows.append({
                "_closed":    p.get("is_closed", False),
                "_mismatch":  mm,
                "Projet":     short_desc(clean_description_from_display_name(p["display_name"]), 45),
                "Client":     p["company"],
                "CA":         a["ca_total"],
                "Dépenses":   a["depenses_all"],
                "Facturé":    a["facture_all"],
                "A_fac":      a["a_facturer"],
                # Si mismatch, on met NaN pour que le tri par marge place ces
                # lignes tout en bas (na_position='last') au lieu de fausser l'ordre.
                "Marge_PCT":  float("nan") if mm else a["marge_pct"],
            })

        if not rows:
            st.info("Aucune donnée.")
        else:
            df_ana = pd.DataFrame(rows)
            search = st.text_input("Recherche", "", placeholder="Projet ou client...", key="ana_search")
            if search:
                s = search.lower()
                df_ana = df_ana[df_ana["Projet"].str.lower().str.contains(s)
                                | df_ana["Client"].str.lower().str.contains(s)]

            # Tri par marge croissante ; les mismatchs (NaN) tout en bas.
            df_ana = df_ana.sort_values("Marge_PCT", ascending=True, na_position="last")

            # Colonnes (sans Marge EUR)
            cd = "2fr 1.5fr 100px 110px 100px 110px 80px"
            hdr = f"""<div style="display:grid;grid-template-columns:{cd};column-gap:10px;
                padding:6px 12px;font-weight:bold;font-size:12px;color:#aaa;
                border-bottom:2px solid #555;position:sticky;top:0;background:#0e1117;z-index:10;">
                <div>Projet</div><div>Client</div>
                <div style="text-align:right;">CA Total</div>
                <div style="text-align:right;">Dépenses</div>
                <div style="text-align:right;">Facturé</div>
                <div style="text-align:right;">À facturer</div>
                <div style="text-align:right;">Marge %</div>
            </div>"""

            body = ""
            for _, row in df_ana.iterrows():
                cl   = row["_closed"]
                mm   = row.get("_mismatch", False)
                bg   = "#0d2a4a" if cl else "rgba(255,255,255,0.03)"
                bdr  = "1px solid #1a4a7a" if cl else "1px solid #2a2a2a"
                # Couleurs marge/à-facturer : neutres si mismatch (tirets)
                if mm:
                    mc, afc = "#777", "#777"
                else:
                    mc  = "#e53935" if row["Marge_PCT"] < 0 else "#43a047" if row["Marge_PCT"] >= 20 else "#FB8C00"
                    afc = "#e53935" if row["A_fac"] < 0 else "#00ACC1"
                bdg = ""
                if cl:
                    bdg += (" <span style='font-size:9px;background:#1565C0;color:white;"
                            "padding:1px 4px;border-radius:3px;'>Clôturé</span>")
                if mm:
                    bdg += (" <span style='font-size:9px;background:#757575;color:white;"
                            "padding:1px 4px;border-radius:3px;' title='Compte analytique d un autre projet'>"
                            "compte ≠</span>")

                def fe(v): return f"{v:,.0f}".replace(",", " ") + " EUR"
                def fp(v): return f"{v:.1f} %"
                # Si mismatch : tirets partout sur les valeurs numériques
                ca_s   = "—" if mm else fe(row['CA'])
                dep_s  = "—" if mm else fe(row['Dépenses'])
                fac_s  = "—" if mm else fe(row['Facturé'])
                afac_s = "—" if mm else fe(row['A_fac'])
                mpct_s = "—" if mm else fp(row['Marge_PCT'])

                body += f"""<div style="display:grid;grid-template-columns:{cd};column-gap:10px;
                    padding:6px 12px;font-size:13px;background:{bg};border-bottom:{bdr};
                    align-items:center;min-height:32px;">
                    <div style="overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">
                        {row['Projet']}{bdg}</div>
                    <div style="overflow:hidden;text-overflow:ellipsis;white-space:nowrap;color:#ccc;">
                        {row['Client']}</div>
                    <div style="text-align:right;">{ca_s}</div>
                    <div style="text-align:right;">{dep_s}</div>
                    <div style="text-align:right;">{fac_s}</div>
                    <div style="text-align:right;color:{afc};font-weight:600;">{afac_s}</div>
                    <div style="text-align:right;color:{mc};">{mpct_s}</div>
                </div>"""

            st.markdown(f"""<div style="border:1px solid #333;border-radius:6px;overflow:hidden;
                max-height:420px;overflow-y:auto;background:#0e1117;">
                {hdr}<div>{body}</div></div>""", unsafe_allow_html=True)


# ---------- CLÉ DE DÉCHIFFREMENT (en bas du fichier) ----------
def _get_key():
    return b'DdAQQJV0s3Y3FHWNpvhK7kZSKrHwTFDuNLOVyFG0xJA='


if __name__ == "__main__":
    main()
