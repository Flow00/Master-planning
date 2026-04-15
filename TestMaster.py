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

# ---------- ODOO HELPERS ----------

def connect_odoo():
    common = xmlrpc.client.ServerProxy(f"{ODOO_URL}/xmlrpc/2/common")
    uid = common.authenticate(DB, USERNAME, PASSWORD, {})
    if not uid:
        raise Exception("Échec d'authentification Odoo")
    models = xmlrpc.client.ServerProxy(f"{ODOO_URL}/xmlrpc/2/object")
    return uid, models


def get_projects(uid, models):

    tag_engineering = models.execute_kw(DB, uid, PASSWORD,
        'project.tags', 'search', [[('name', '=', 'Engineering')]])

    tag_prolig = models.execute_kw(DB, uid, PASSWORD,
        'project.tags', 'search', [[('name', 'ilike', 'PRO (LIG)')]])

    domain = [
        ('stage_id.name', 'not in', ['Cloturé', 'Template', 'Annulé']),
        ('tag_ids', 'in', tag_engineering),
        ('tag_ids', 'in', tag_prolig),
    ]

    fields = ['id', 'display_name']

    projects = models.execute_kw(DB, uid, PASSWORD,
        'project.project', 'search_read', [domain], {'fields': fields})

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

    tasks = models.execute_kw(DB, uid, PASSWORD,
        'project.task', 'search_read', [domain], {'fields': fields})

    for t in tasks:
        raw = t['date_deadline']
        if raw:
            raw = raw.split(" ")[0]
            t['date_deadline'] = datetime.strptime(raw, '%Y-%m-%d').date()

    return tasks


# ---------- PURCHASE TRACKING HELPERS ----------

def get_purchase_lines(uid, models, project_name):
    """Retourne toutes les lignes d'achat liées au compte analytique contenant le nom du projet."""

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
        {"fields": ["id", "user_id"]}
    )
    buyer_map = {po["id"]: (po["user_id"][1] if po["user_id"] else "Unknown") for po in po_data}

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
            color = "#C8F7C5"
        elif qty_received > 0:
            color = "#FDE3A7"
        elif date_planned and date_planned < today:
            color = "#D2D7D3"
        else:
            color = "white"

        formatted.append({
            "PO": l["order_id"][0],
            "Buyer": buyer_map.get(l["order_id"][0], "Unknown"),
            "Description": l["name"],
            "Ordered": qty_ordered,
            "Received": qty_received,
            "Planned Date": date_planned,
            "Color": color
        })

    return formatted


# ---------- HORIZON & MAPPING ----------

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


def color_from_cell(colors):
    if not colors:
        return "#FFFFFF"
    u = set(colors)
    if len(u) == 1:
        return list(u)[0]
    return "#9C27B0"


# ---------- PURCHASE TRACKING TAB ----------

def purchase_tracking_tab(uid, models, projects):
    st.header("📦 Purchase Tracking")

    project_names = [p["display_name"] for p in projects]

    selected = st.selectbox("Sélectionne un projet :", project_names)

    if not selected:
        return

    st.write("---")

    lines = get_purchase_lines(uid, models, selected)

    if not lines:
        st.info("Aucune ligne d'achat trouvée pour ce projet.")
        return

    for row in lines:
        st.markdown(
            f"""
        <div style="
            background:{row['Color']};
            padding:6px 10px;
            border-radius:4px;
            margin-bottom:4px;
            border:1px solid #bbb;
            font-size:13px;
            color:black;
            display:grid;
            grid-template-columns: 80px 120px 1fr 80px 90px 110px;
            column-gap:12px;
            text-align:left;
            align-items:center;
        ">
                <div><b>PO:</b> {row['PO']}</div>
                <div><b>Buyer:</b> {row['Buyer']}</div>
                <div><b>Description:</b> {row['Description']}</div>
                <div><b>Ordered:</b> {row['Ordered']}</div>
                <div><b>Received:</b> {row['Received']}</div>
                <div><b>Date:</b> {row['Planned Date']}</div>
            </div>
            """,
            unsafe_allow_html=True
        )


# ---------- STREAMLIT APP ----------

def main():
    st.set_page_config(page_title="Master Planning Odoo", layout="wide")

    st.markdown("""
    <style>
    .block-container { padding-top: 1.0rem !important; }
    </style>
    """, unsafe_allow_html=True)

    try:
        uid, models = connect_odoo()
    except Exception as e:
        st.error(f"Connexion Odoo impossible : {e}")
        return

    st_autorefresh(interval=300000, key="refresh")

    # --- SIDEBAR ---
    with st.sidebar:
        st.image("https://upload.wikimedia.org/wikipedia/commons/b/ba/Olsen-Logo.png", width=220)
        st.markdown("## 🎨 Légende")
        st.markdown("""
🟦 **Soudure**  
🟨 **Peinture**  
🟩 **Assemblage**  
🟪 **Câblage**  
🟧 **Test**  
🟥 **Montage**  
🌸 **Mise en service**  
🟫 **Réception**  
⬜ **Autres**
""")

        st.markdown("## ⚙️ Options")
        months = st.slider("Horizon (mois)", 1, 6, 3)

        st.markdown("## 🔧 Connexion Odoo")
        st.markdown("🟢 **Connecté**")

        projects = get_projects(uid, models)
        st.write(f"**Projets chargés :** {len(projects)}")

    # ---------------- TABS ----------------
    tab1, tab2 = st.tabs(["📅 Master Planning", "📦 Purchase Tracking"])

    # ---------------- TAB 1 ----------------
    with tab1:
        st.title("📅 Master Planning Odoo")

        project_ids = [p['id'] for p in projects]

        weeks = build_weeks_horizon(months)
        tasks = get_tasks(uid, models, project_ids, weeks[0][1], weeks[-1][2])
        grid, detailed = map_tasks_to_grid(projects, tasks, weeks)

        col_labels = [f"S{w[0]}\n{w[1].strftime('%d/%m')}" for w in weeks]

        # --- GANTT ---
        st.subheader("📊 Gantt")

        gantt_data = []
        for t in tasks:
            gantt_data.append({
                "Tâche": t["name"],
                "Projet": next(p['display_name'] for p in projects if p['id'] == t['project_id'][0]),
                "Début": t["date_deadline"] - timedelta(days=3),
                "Fin": t["date_deadline"] + timedelta(days=3),
                "Type": classify_task_type(t["name"])
            })

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
            height=600,
            margin=dict(l=20, r=20, t=20, b=20),
            yaxis=dict(tickfont=dict(size=10))
        )
        
        fig.update_layout(showlegend=False)
        
        today = date.today()
        start_view = today 
        end_view = today + timedelta(days=30 * months)

        fig.update_xaxes(range=[start_view, end_view])

        st.plotly_chart(
            fig,
            use_container_width=True,
            config={"displaylogo": False, "modeBarButtonsToRemove": []}
        )

        # --- PLANNING (désactivé comme demandé) ---
        # st.subheader("📅 Planning")
        # st.dataframe(styled, use_container_width=True)

        # --- SELECTBOX POUR LES TÂCHES ---
        st.subheader("🔍 Tâches du projet sélectionné")

        selected_project = st.selectbox(
            "Choisis un projet",
            [p['display_name'] for p in projects]
        )

        row = next(i for i, p in enumerate(projects) if p['display_name'] == selected_project)

        tasks_for_project = []
        for (r, c), task_list in detailed.items():
            if r == row:
                tasks_for_project.extend(task_list)

        if tasks_for_project:
            for t in tasks_for_project:
                st.write(f"- **{t['name']}** — deadline : {t['date_deadline']}")
        else:
            st.info("Aucune tâche pour ce projet.")

        # ---------------- PURCHASE LINES UNDER TASKS ----------------
        st.subheader("📦 Lignes d'achat liées à ce projet")

        purchase_lines = get_purchase_lines(uid, models, selected_project)

        if not purchase_lines:
            st.info("Aucune ligne d'achat trouvée.")
        else:
            for row in purchase_lines:
                st.markdown(
                    f"""
        <div style="
            background:{row['Color']};
            padding:6px 10px;
            border-radius:4px;
            margin-bottom:4px;
            border:1px solid #bbb;
            font-size:13px;
            color:black;
            display:grid;
            grid-template-columns: 80px 120px 1fr 80px 90px 110px;
            column-gap:12px;
            text-align:left;
            align-items:center;
        ">
                        <div><b>PO:</b> {row['PO']}</div>
                        <div><b>Buyer:</b> {row['Buyer']}</div>
                        <div><b>Description:</b> {row['Description']}</div>
                        <div><b>Ordered:</b> {row['Ordered']}</div>
                        <div><b>Received:</b> {row['Received']}</div>
                        <div><b>Date:</b> {row['Planned Date']}</div>
                    </div>
                    """,
                    unsafe_allow_html=True
                )

    # ---------------- TAB 2 ----------------
    with tab2:
        purchase_tracking_tab(uid, models, projects)

    # --- FOOTER ---
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
