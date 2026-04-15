import xmlrpc.client
from datetime import datetime, timedelta, date
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


def get_top_company(uid, models, partner_id):
    if not partner_id:
        return "N/A"

    pid = partner_id[0]
    data = models.execute_kw(
        DB, uid, PASSWORD,
        "res.partner", "read",
        [pid], {"fields": ["name", "parent_id"]}
    )[0]

    if not data["parent_id"]:
        return data["name"]

    parent = data["parent_id"]
    return parent[1]


def get_projects(uid, models):
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

    for p in projects:
        p["company"] = get_top_company(uid, models, p["partner_id"])

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
# 🔧 PURCHASE TRACKING HELPERS (ULTRA RAPIDE : PAR PROJET UNIQUEMENT)
# ============================================================

def get_purchase_lines(uid, models, project_name):
    analytic_ids = models.execute_kw(
        DB, uid, PASSWORD,
        "account.analytic.account", "search",
        [[("name", "ilike", project_name)]]
    )

    if not analytic_ids:
        return []

    po_ids = models.execute_kw(
        DB, uid, PASSWORD,
        "purchase.order", "search",
        [[("analytic_account_id", "in", analytic_ids)]]
    )

    if not po_ids:
        return []

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
            ]
        }
    )

    today = date.today()
    formatted = []

    for l in lines:
        qty_ordered = l["product_qty"]
        qty_received = l["qty_received"]

        if l["date_planned"]:
            d = l["date_planned"].split(" ")[0]
            date_planned = datetime.strptime(d, "%Y-%m-%d").date()
        else:
            date_planned = None

        if qty_received >= qty_ordered:
            color = "#2E7D32"; rank = 3          # vert vif
        elif qty_received > 0:
            color = "#FFA000"; rank = 0          # orange vif
        elif date_planned and date_planned < today:
            color = "#757575"; rank = 1          # gris foncé
        else:
            color = "#FFFFFF"; rank = 2          # blanc

        po_id = l["order_id"][0]
        po_name = po_name_map.get(po_id, str(po_id))

        formatted.append({
            "PO": po_name,
            "Buyer": buyer_map.get(po_id, "Unknown"),
            "Description": l["name"],
            "Ordered": qty_ordered,
            "Received": qty_received,
            "Planned Date": date_planned,
            "Color": color,
            "Rank": rank
        })

    formatted.sort(key=lambda x: x["Rank"])
    return formatted


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
        connected = True
    except Exception as e:
        st.error(f"Connexion Odoo impossible : {e}")
        return

    # 🔁 Refresh toutes les 10 minutes
    st_autorefresh(interval=600000, key="refresh_10min")

    if "months" not in st.session_state:
        st.session_state["months"] = 3

    if "selected_purchase_project_id" not in st.session_state:
        st.session_state["selected_purchase_project_id"] = None

    # ---------- BANNIÈRE ----------
    col1, col2, col3 = st.columns([1, 3, 1])
    with col1:
        st.image("https://upload.wikimedia.org/wikipedia/commons/b/ba/Olsen-Logo.png", width=180)
    with col2:
        st.markdown("<h2 style='text-align:center;margin-top:10px;'>Master Planning</h2>", unsafe_allow_html=True)
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
        projects = get_projects(uid, models)
        project_ids = [p['id'] for p in projects]

        months = st.session_state["months"]
        weeks = build_weeks_horizon(months)
        tasks = get_tasks(uid, models, project_ids, weeks[0][1], weeks[-1][2])
        grid, detailed = map_tasks_to_grid(projects, tasks, weeks)

        def project_label(p):
            client = p["company"]
            code = p.get('name') or p['display_name']
            desc = p['display_name']
            if " - " in desc:
                desc = desc.split(" - ", 1)[1]
            base_desc = desc
            desc_short = base_desc if len(base_desc) <= 20 else base_desc[:20] + "…"
            return f"{client} - {code} - {desc_short}"

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
                yaxis=dict(tickfont=dict(size=11))
            )

            fig.update_layout(
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

            st.plotly_chart(
                fig,
                use_container_width=True,
                config={"displaylogo": False}
            )
        else:
            st.info("Aucune tâche à afficher dans le Gantt.")

        # Slider SOUS le Gantt
        col_s1, col_s2 = st.columns([3, 1])
        with col_s1:
            new_months = st.slider("", 1, 6, months)
        with col_s2:
            st.markdown(f"<div style='margin-top:10px;'>Projets : <b>{len(projects)}</b></div>", unsafe_allow_html=True)

        if new_months != months:
            st.session_state["months"] = new_months
            st.experimental_rerun()

        # Recherche des tâches par projet
        st.subheader("🔍 Tâches du projet sélectionné")

        project_labels = [project_label(p) for p in projects]
        selected_label = st.selectbox("Choisis un projet", project_labels)

        selected_project = next(p for p in projects if project_label(p) == selected_label)
        row_index = next(i for i, p in enumerate(projects) if p['id'] == selected_project['id'])

        tasks_for_project = []
        for (r, c), task_list in detailed.items():
            if r == row_index:
                tasks_for_project.extend(task_list)

        if tasks_for_project:
            for t in tasks_for_project:
                st.write(f"- **{t['name']}** — deadline : {t['date_deadline']}")
        else:
            st.info("Aucune tâche pour ce projet.")

    # ============================================================
    # 🟩 ONGLET 2 — PURCHASE TRACKING (ULTRA RAPIDE)
    # ============================================================
    with tab2:
        st.markdown("### 📦 Purchases par projet")

        projects = get_projects(uid, models)

        cols_per_row = 5
        for i in range(0, len(projects), cols_per_row):
            cols = st.columns(cols_per_row)
            for col, p in zip(cols, projects[i:i+cols_per_row]):
                with col:
                    client = p["company"]
                    code = p.get('name') or p['display_name']
                    desc = p['display_name']
                    if " - " in desc:
                        desc = desc.split(" - ", 1)[1]
                    desc_short = (desc[:40] + "…") if len(desc) > 40 else desc

                    if st.button(
                        f"{client}\n{code}\n{desc_short}",
                        key=f"proj_btn_{p['id']}"
                    ):
                        st.session_state["selected_purchase_project_id"] = p['id']

        st.markdown("---")
        st.subheader("📋 Détail des lignes d'achat du projet sélectionné")

        selected_purchase_project_id = st.session_state.get("selected_purchase_project_id", None)
        if selected_purchase_project_id is None:
            st.info("Clique sur une vignette projet pour charger les lignes d'achat.")
        else:
            p = next(p for p in projects if p['id'] == selected_purchase_project_id)
            st.markdown(f"**Projet sélectionné :** {p['company']} - {p.get('name') or p['display_name']}")

            # 🔥 Chargement des lignes UNIQUEMENT pour ce projet
            lines = get_purchase_lines(uid, models, p['display_name'])

            if not lines:
                st.info("Aucune ligne d'achat trouvée pour ce projet.")
            else:
                st.markdown(f"**Total lignes : {len(lines)}**")
                for row in lines:
                    st.markdown(
                        f"""
                        <div style="
                            background:{row['Color']};
                            padding:6px 10px;
                            border-radius:4px;
                            margin-bottom:4px;
                            border:1px solid #555;
                            font-size:13px;
                            color:black;
                            display:grid;
                            grid-template-columns: 90px 190px 1fr 80px 90px 110px;
                            column-gap:12px;
                            text-align:left;
                            align-items:center;
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
