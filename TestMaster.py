import xmlrpc.client
from datetime import datetime, timedelta, date
import streamlit as st
import pandas as pd
import plotly.express as px
from streamlit_autorefresh import st_autorefresh
from streamlit_plotly_events import plotly_events

# ---------- CONFIG ODOO ----------
ODOO_URL = "https://olsen-engineering.odoo.com"
DB = "mynalios-olsen-main-7388485"
USERNAME = "f.mordant@olsen-engineering.com"
PASSWORD = "a9a52b95f9ba02f3d813aa02e113d51ffac6de1d"


# ============================================================
# 🔧 ODOO HELPERS + CACHE
# ============================================================

def connect_odoo():
    common = xmlrpc.client.ServerProxy(f"{ODOO_URL}/xmlrpc/2/common")
    uid = common.authenticate(DB, USERNAME, PASSWORD, {})
    if not uid:
        raise Exception("Échec d'authentification Odoo")
    models = xmlrpc.client.ServerProxy(f"{ODOO_URL}/xmlrpc/2/object")
    return uid, models


@st.cache_data(ttl=300)
def cached_get_projects():
     uid, models = connect_odoo()

    tag_engineering = models.execute_kw(
        DB, uid, PASSWORD,
        'project.tags', 'search', [[('name', '=', 'Engineering')]]
    )

    tag_prolig = models.execute_kw(
        DB, uid, PASSWORD,
        'project.tags', 'search', [[('name', 'ilike', 'PRO (LIG)')]]
    )

    domain = [
        ('stage_id.name', 'not in', ['Cloturé', 'Template', 'Annulé']),
        ('tag_ids', 'in', tag_engineering),
        ('tag_ids', 'in', tag_prolig),
    ]

    fields = ['id', 'display_name', 'partner_id', 'name']

    projects = models.execute_kw(
        DB, uid, PASSWORD,
        'project.project', 'search_read', [domain], {'fields': fields}
    )

    # Top company
    for p in projects:
        partner_id = p["partner_id"][0] if p["partner_id"] else False
        if not partner_id:
            p["company"] = "N/A"
            continue

        data = models.execute_kw(
            DB, uid, PASSWORD,
            "res.partner", "read",
            [partner_id], {"fields": ["name", "parent_id"]}
        )[0]

        if not data["parent_id"]:
            p["company"] = data["name"]
        else:
            p["company"] = data["parent_id"][1]

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
    filtered.sort(key=lambda p: p['display_name'])
    return filtered


def get_tasks(uid, models, project_ids, start_date, end_date):
    domain = [
        ('project_id', 'in', project_ids),
        ('date_deadline', '!=', False),
        ('tag_ids.name', 'in', ['Engineering', 'PRO (LIG)', 'PRO(LIG)']),
    ]

    fields = ['id', 'name', 'project_id', 'date_deadline']

    tasks = models.execute_kw(
        DB, uid, PASSWORD,
        'project.task', 'search_read', [domain], {'fields': fields}
    )

    for t in tasks:
        raw = t['date_deadline']
        if raw:
            raw = raw.split(" ")[0]
            t['date_deadline'] = datetime.strptime(raw, '%Y-%m-%d').date()

    return tasks


# ============================================================
# 🔧 PARSING DES NOMS / DESCRIPTIONS
# ============================================================

def clean_description_from_display_name(display_name: str) -> str:
    if " - " not in display_name:
        return display_name

    parts = display_name.split(" - ")
    if len(parts) >= 2 and parts[-1].strip() == parts[-2].strip():
        parts = parts[:-1]

    desc_clean = " - ".join(parts)
    return desc_clean


def short_desc(desc: str, max_len: int) -> str:
    if len(desc) <= max_len:
        return desc
    return desc[:max_len].rstrip() + "…"


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
    "Soudure",
    "Peinture",
    "Assemblage",
    "Câblage",
    "Test",
    "Montage",
    "Mise en service",
    "Réception",
    "Autres"
]

COLOR_MAP = {
    "Soudure": "#1E88E5",
    "Peinture": "#FDD835",
    "Assemblage": "#43A047",
    "Câblage": "#8E24AA",
    "Test": "#FB8C00",
    "Montage": "#E53935",
    "Mise en service": "#EC407A",
    "Réception": "#6D4C41",
    "Autres": "#9E9E9E"
}


def classify_task_type(name):
    n = name.lower()

    if "soud" in n: return "Soudure"
    if "peint" in n: return "Peinture"
    if "assembl" in n: return "Assemblage"
    if "cabl" in n or "câbl" in n: return "Câblage"
    if "test" in n: return "Test"
    if "montage" in n: return "Montage"
    if "mise en service" in n or "mes" in n: return "Mise en service"
    if "récept" in n or "recept" in n: return "Réception"

    return "Autres"
# ============================================================
# 🔧 PURCHASE LOADER (SUMMARY + LIGNES, AVEC FILTRE INVOICE POLICY)
# ============================================================

@st.cache_data(ttl=300)
def load_purchase_data(uid, models, project_name):
    analytic_ids = models.execute_kw(
        DB, uid, PASSWORD,
        "account.analytic.account", "search",
        [[("name", "ilike", project_name)]]
    )

    if not analytic_ids:
        return {
            "summary": {"orange": 0, "grey": 0, "white": 0, "green": 0, "total": 0},
            "lines": []
        }

    po_ids = models.execute_kw(
        DB, uid, PASSWORD,
        "purchase.order", "search",
        [[
            ("analytic_account_id", "in", analytic_ids),
            ("state", "=", "purchase"),
        ]]
    )

    if not po_ids:
        return {
            "summary": {"orange": 0, "grey": 0, "white": 0, "green": 0, "total": 0},
            "lines": []
        }

    po_data = models.execute_kw(
        DB, uid, PASSWORD,
        "purchase.order", "read",
        [po_ids],
        {"fields": ["id", "user_id", "name"]}
    )
    buyer_map = {po["id"]: (po["user_id"][1] if po["user_id"] else "Unknown") for po in po_data}
    po_name_map = {po["id"]: po["name"] for po in po_data}

    lines = models.execute_kw(
        DB, uid, PASSWORD,
        "purchase.order.line", "search_read",
        [[("order_id", "in", po_ids)]],
        {
            "fields": [
                "name",
                "product_qty",
                "qty_received",
                "date_planned",
                "order_id",
                "product_id",
            ]
        }
    )

    product_ids = list({l["product_id"][0] for l in lines if l.get("product_id")})
    policy_map = {}
    if product_ids:
        products = models.execute_kw(
            DB, uid, PASSWORD,
            "product.product", "read",
            [product_ids],
            {"fields": ["invoice_policy"]}
        )
        policy_map = {p["id"]: p["invoice_policy"] for p in products}

    today = date.today()
    orange = grey = white = green = 0
    formatted = []

    for l in lines:
        pid = l.get("product_id")
        if pid:
            product_id = pid[0]
            policy = policy_map.get(product_id)
            if policy == "order":
                continue  # exclu partout

        qty_ordered = l["product_qty"]
        qty_received = l["qty_received"]

        if l["date_planned"]:
            d = l["date_planned"].split(" ")[0]
            date_planned = datetime.strptime(d, "%Y-%m-%d").date()
        else:
            date_planned = None

        if qty_received >= qty_ordered:
            color = "#2E7D32"; rank = 3
            green += 1
        elif qty_received > 0:
            color = "#FFA000"; rank = 0
            orange += 1
        elif date_planned and date_planned < today:
            color = "#757575"; rank = 1
            grey += 1
        else:
            color = "#FFFFFF"; rank = 2
            white += 1

        po_id = l["order_id"][0]
        po_name = po_name_map.get(po_id, str(po_id))

        formatted.append({
            "PO": po_name,
            "Buyer": buyer_map.get(po_id, "Unknown"),
            "Description": short_desc(l["name"], 50),
            "Ordered": qty_ordered,
            "Received": qty_received,
            "Planned Date": date_planned,
            "Color": color,
            "Rank": rank
        })

    total = orange + grey + white + green
    formatted.sort(key=lambda x: x["Rank"])

    summary = {
        "orange": orange,
        "grey": grey,
        "white": white,
        "green": green,
        "total": total
    }

    return {
        "summary": summary,
        "lines": formatted
    }


# ============================================================
# 🔵 STREAMLIT APP
# ============================================================

def main():
    st.set_page_config(page_title="Master Planning Odoo", layout="wide")

    st.markdown("""
    <style>
    .block-container { padding-top: 0.5rem !important; }
    </style>
    """, unsafe_allow_html=True)

    try:
        uid, models = connect_odoo()
        projects = cached_get_projects()
    except Exception as e:
        st.error(f"Connexion Odoo impossible : {e}")
        return

    st_autorefresh(interval=600000, key="refresh_10min")

    if "months" not in st.session_state:
        st.session_state["months"] = 3
    if "selected_purchase_project_id" not in st.session_state:
        st.session_state["selected_purchase_project_id"] = None
    if "gantt_filter" not in st.session_state:
        st.session_state["gantt_filter"] = None

    col1, col2, col3 = st.columns([1, 3, 1])
    with col1:
        st.image("https://upload.wikimedia.org/wikipedia/commons/b/ba/Olsen-Logo.png", width=180)
    with col2:
        st.markdown(
            "<h2 style='text-align:center;margin-top:10px;'>Olsen Dashboard</h2>",
            unsafe_allow_html=True
        )
    with col3:
        st.markdown(
            "<div style='text-align:right;color:green;font-weight:bold;margin-top:20px;'>🟢 Connecté à Odoo</div>",
            unsafe_allow_html=True
        )

    tab1, tab2 = st.tabs(["📅 Planning", "📦 Purchases"])

    # ============================================================
    # 🟦 ONGLET 1 — MASTER PLANNING
    # ============================================================
    with tab1:
        projects = cached_get_projects()
        project_ids = [p['id'] for p in projects]

        months = st.session_state["months"]
        weeks = build_weeks_horizon(months)
        tasks = get_tasks(uid, models, project_ids, weeks[0][1], weeks[-1][2])

        def project_label(p):
            client = p["company"]
            desc_clean = clean_description_from_display_name(p['display_name'])
            desc_short = short_desc(desc_clean, 20)
            return f"{client} - {desc_short}"

        st.subheader("📊 Gantt")

        gantt_data = []
        for t in tasks:
            proj = next(p for p in projects if p['id'] == t['project_id'][0])
            gantt_data.append({
                "Tâche": t["name"],
                "Projet": project_label(proj),
                "Début": t["date_deadline"] - timedelta(days=3),
                "Fin": t["date_deadline"] + timedelta(days=3),
                "Type": classify_task_type(t["name"])
            })

        if gantt_data:
            df_gantt = pd.DataFrame(gantt_data)
            df_gantt["Type détaillé"] = pd.Categorical(
                df_gantt["Type"],
                categories=COLOR_ORDER,
                ordered=True
            )

            current_filter = st.session_state.get("gantt_filter", None)
            if current_filter:
                df_gantt = df_gantt[df_gantt["Projet"] == current_filter]

            fig = px.timeline(
                df_gantt,
                x_start="Début",
                x_end="Fin",
                y="Projet",
                color="Type détaillé",
                color_discrete_map=COLOR_MAP
            )

            fig.update_yaxes(autorange="reversed")
            fig.update_layout(
                dragmode="pan",
                height=650,
                margin=dict(l=20, r=20, t=40, b=20),
                yaxis=dict(tickfont=dict(size=12)),
                legend=dict(
                    orientation="h",
                    yanchor="bottom",
                    y=1.02,
                    xanchor="center",
                    x=0.5,
                    font=dict(size=10)
                )
            )

            today = date.today()
            fig.add_vline(
                x=today,
                line_width=2,
                line_color="white",
                opacity=0.9
            )

            start_view = today
            end_view = today + timedelta(days=30 * months)
            fig.update_xaxes(range=[start_view, end_view])

            clicked = plotly_events(
                fig,
                click_event=True,
                hover_event=False,
                select_event=False,
                override_height=650,
                override_width="100%"
            )

            if clicked:
                proj_clicked = clicked[0]["y"]
                if st.session_state.get("gantt_filter", None) == proj_clicked:
                    st.session_state["gantt_filter"] = None
                else:
                    st.session_state["gantt_filter"] = proj_clicked

            if st.session_state.get("gantt_filter", None):
                st.markdown(f"**Projet filtré :** {st.session_state['gantt_filter']}")
            else:
                st.markdown("**Vue : Tous les projets**")
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
            st.experimental_rerun()

        st.subheader("🔍 Tâches du projet sélectionné")

        if st.session_state.get("gantt_filter", None):
            proj_label = st.session_state["gantt_filter"]
            proj = next(p for p in projects if project_label(p) == proj_label)
            proj_id = proj["id"]

            tasks_for_project = [t for t in tasks if t["project_id"][0] == proj_id]

            if tasks_for_project:
                for t in tasks_for_project:
                    st.write(f"- **{t['name']}** — deadline : {t['date_deadline']}")
            else:
                st.info("Aucune tâche pour ce projet.")
        else:
            st.info("Clique sur une barre du Gantt pour filtrer un projet.")
    # ============================================================
    # 🟩 ONGLET 2 — PURCHASE TRACKING (HYBRIDE)
    # ============================================================
    with tab2:
        st.markdown("### 📦 Purchases par projet")

        projects = cached_get_projects()

        purchase_data = {
            p['id']: load_purchase_data(uid, models, p['display_name'])
            for p in projects
        }

        cols_per_row = 6
        for i in range(0, len(projects), cols_per_row):
            cols = st.columns(cols_per_row)
            for col, p in zip(cols, projects[i:i+cols_per_row]):
                with col:
                    client = p["company"]
                    desc_clean = clean_description_from_display_name(p['display_name'])
                    desc_short_25 = short_desc(desc_clean, 25)

                    data = purchase_data[p['id']]
                    summary = data["summary"]
                    total = summary["total"]
                    total_safe = max(total, 1)
                    non_green = total - summary["green"]
                    pct_orange = 100 * summary["orange"] / total_safe
                    pct_grey = 100 * summary["grey"] / total_safe
                    pct_white = 100 * summary["white"] / total_safe
                    pct_green = 100 * summary["green"] / total_safe

                    if non_green == 0:
                        text_color = "white"
                    elif summary["grey"] > 0:
                        text_color = "red"
                    else:
                        text_color = "#FFA000"

                    if st.button(
                        f"{client}\n {desc_short_25}",
                        key=f"proj_btn_{p['id']}"
                    ):
                        st.session_state["selected_purchase_project_id"] = p['id']

                    st.markdown(
                        f"""
                        <div style="
                            width:100%;
                            height:12px;
                            border-radius:6px;
                            overflow:hidden;
                            display:flex;
                            margin-top:4px;
                            border:1px solid #444;
                        ">
                            <div style="width:{pct_orange}%;background:#FFA000;"></div>
                            <div style="width:{pct_grey}%;background:#757575;"></div>
                            <div style="width:{pct_white}%;background:#FFFFFF;"></div>
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
            p = next(p for p in projects if p['id'] == selected_purchase_project_id)
            st.markdown(
                f"<div style='font-size:15px;'><b>Projet sélectionné :</b> "
                f"{p['company']} - {p.get('name') or p['display_name']}</div>",
                unsafe_allow_html=True
            )

            data = purchase_data[p['id']]
            lines = data["lines"]

            if not lines:
                st.info("Aucune ligne d'achat trouvée pour ce projet.")
            else:
                st.markdown(f"**Total lignes : {len(lines)}**")
                for row in lines:
                    st.markdown(
                        f"""
                        <div style="
                            background:{row['Color']};
                            padding:8px 12px;
                            border-radius:4px;
                            margin-bottom:5px;
                            border:1px solid #555;
                            font-size:14px;
                            color:black;
                            display:grid;
                            grid-template-columns: 90px 190px 1fr 80px 90px 110px;
                            column-gap:12px;
                            text-align:left;
                            align-items:center;
                            line-height:1.3;
                        ">
                            <div style="white-space:nowrap;"><b>PO:</b> {row['PO']}</div>
                            <div style="white-space:nowrap;"><b>Buyer:</b> {row['Buyer']}</div>
                            <div><b>Description:</b> {row['Description']}</div>
                            <div style="white-space:nowrap;"><b>Ord.:</b> {row['Ordered']}</div>
                            <div style="white-space:nowrap;"><b>Reçu:</b> {row['Received']}</div>
                            <div style="white-space:nowrap;"><b>Date:</b> {row['Planned Date']}</div>
                        </div>
                        """,
                        unsafe_allow_html=True
                    )

    st.markdown("""
    <style>
    .footer {
        position: fixed;
        left: 0;
        bottom: 0;
        width: 100%;
        background-color: rgba(240,240,240,0.85);
        color: #333;
        text-align: center;
        padding: 6px 0;
        font-size: 14px;
        border-top: 1px solid #ccc;
        z-index: 9999;
    }
    </style>

    <div class="footer">
        C Flow - Powered by Olsen-Engineering
    </div>
    """, unsafe_allow_html=True)


if __name__ == "__main__":
    main()
