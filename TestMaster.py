import xmlrpc.client
from datetime import datetime, timedelta, date
import streamlit as st
import pandas as pd
import plotly.express as px

# ---------- CONFIG ODOO ----------
ODOO_URL = "https://olsen-engineering.odoo.com"
DB = "mynalios-olsen-main-7388485"
USERNAME = "f.mordant@olsen-engineering.com"
PASSWORD = "fmo@2022+"

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

    # ---- OPTIMISATION ----
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
        #('stage_id.name', '!=', 'terminée'),
        ('date_deadline', '!=', False),
        #('date_deadline', '>=', start_date.strftime('%Y-%m-%d')),
        #('date_deadline', '<=', end_date.strftime('%Y-%m-%d')),
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


# --- LÉGENDE UNIFORMISÉE ---

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

    st.title("📅 Master Planning Odoo")

    # --- DATA ---
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

    # Mode pan par défaut
    fig.update_layout(
        dragmode="pan",
        height=600,
        margin=dict(l=20, r=20, t=20, b=20),
        yaxis=dict(tickfont=dict(size=10))
    )
    
    fig.update_layout(showlegend=False)
    
    # 👉 Vue centrée sur aujourd’hui, dépendante du slider "months"
    today = date.today()
    start_view = today 
    end_view = today + timedelta(days=30 * months)

    fig.update_xaxes(range=[start_view, end_view])

    # Bouton plein écran
    st.plotly_chart(
        fig,
        use_container_width=True,
        config={"displaylogo": False, "modeBarButtonsToRemove": []}
    )

    # --- PLANNING ---
    st.subheader("📅 Planning")

    data = []
    for row, p in enumerate(projects):
        row_vals = []
        for col in range(len(weeks)):
            colors = grid.get((row, col), [])
            row_vals.append(len(colors))
        data.append(row_vals)

    df = pd.DataFrame(data, index=[p['display_name'] for p in projects], columns=col_labels)

    styled = df.style.apply(
        lambda _: pd.DataFrame(
            [[f"background-color: {color_from_cell(grid.get((r, c), []))}"
              for c in range(len(weeks))]
             for r in range(len(projects))],
            index=df.index, columns=df.columns
        ),
        axis=None
    )

    st.dataframe(styled, use_container_width=True)

    # --- CSS COMPACT ---
    st.markdown("""
    <style>

    div[data-testid="stDataFrame"] table {
        margin-right: 140px !important;
    }

    div[data-testid="stDataFrame"] table tbody th {
        max-width: 150px !important;
        width: 150px !important;
        overflow: hidden !important;
        text-overflow: ellipsis !important;
        white-space: nowrap !important;
    }

    div[data-testid="stDataFrame"] table tbody td:nth-child(2),
    div[data-testid="stDataFrame"] table thead th:nth-child(2) {
        max-width: 55px !important;
        width: 55px !important;
    }

    div[data-testid="stDataFrame"] table tbody td:nth-child(n+3),
    div[data-testid="stDataFrame"] table thead th:nth-child(n+3) {
        max-width: 32px !important;
        width: 32px !important;
        padding-left: 1px !important;
        padding-right: 1px !important;
    }

    div[data-testid="stDataFrame"] table tbody tr td,
    div[data-testid="stDataFrame"] table tbody tr th {
        padding-top: 2px !important;
        padding-bottom: 2px !important;
        height: 10px !important;
    }

    </style>
    """, unsafe_allow_html=True)

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
