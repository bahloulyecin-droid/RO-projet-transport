"""
Problème de Transport — Application Streamlit
Mini-projet de Recherche Opérationnelle
"""
import io, random
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import pulp
import streamlit as st

st.set_page_config(page_title="Problème de Transport — RO", page_icon="🚛",
                   layout="wide", initial_sidebar_state="expanded")

st.markdown("""<style>
.main-header{background:linear-gradient(135deg,#1a237e,#3949ab);padding:2rem 2.5rem;
border-radius:14px;margin-bottom:1.8rem;box-shadow:0 4px 20px rgba(26,35,126,.35);}
.main-header h1{color:#fff;margin:0;font-size:2rem;}
.main-header p{color:#c5cae9;margin:.4rem 0 0;font-size:1rem;}
.kpi-card{background:linear-gradient(135deg,#e8eaf6,#fff);border-left:5px solid #3949ab;
border-radius:10px;padding:1.1rem 1.4rem;box-shadow:0 2px 10px rgba(0,0,0,.08);}
.kpi-card h3{color:#3949ab;margin:0 0 .3rem;font-size:.9rem;text-transform:uppercase;}
.kpi-card p{color:#1a237e;margin:0;font-size:1.8rem;font-weight:700;}
.section-title{font-size:1.15rem;font-weight:700;color:#283593;
border-bottom:2px solid #3949ab;padding-bottom:.4rem;margin:1.6rem 0 1rem;}
.badge-ok{background:#e8f5e9;color:#2e7d32;border:1px solid #a5d6a7;
padding:.3rem .8rem;border-radius:20px;font-weight:600;}
.badge-warn{background:#fff8e1;color:#f57f17;border:1px solid #ffe082;
padding:.3rem .8rem;border-radius:20px;font-weight:600;}
.badge-err{background:#ffebee;color:#c62828;border:1px solid #ef9a9a;
padding:.3rem .8rem;border-radius:20px;font-weight:600;}
.upload-box{background:#f8f9ff;border:2px dashed #3949ab;border-radius:12px;
padding:1.5rem;text-align:center;margin:1rem 0;}
</style>""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════
#  FONCTIONS UTILITAIRES
# ══════════════════════════════════════════════════════════════

def make_excel_bytes(costs_matrix, source_names, client_names, capacities, demands):
    """Génère un fichier Excel (3 feuilles) à partir des données fournies."""
    df_costs = pd.DataFrame(costs_matrix, index=source_names, columns=client_names)
    df_caps  = pd.DataFrame({"Capacite": capacities}, index=source_names)
    df_dem   = pd.DataFrame({"Demande":  demands},    index=client_names)
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df_costs.to_excel(w, sheet_name="Couts")
        df_caps.to_excel(w,  sheet_name="Capacites")
        df_dem.to_excel(w,   sheet_name="Demandes")
        for sn in w.sheets:
            ws = w.sheets[sn]
            for col in ws.columns:
                ml = max((len(str(c.value)) for c in col if c.value), default=8)
                ws.column_dimensions[col[0].column_letter].width = ml + 4
    return buf.getvalue()


def generate_random_excel(n_src=3, n_cli=4):
    """Génère un Excel exemple avec données aléatoires cohérentes."""
    rng  = random.Random(42)
    srcs = [f"Source {i+1}" for i in range(n_src)]
    clis = [f"Client {j+1}" for j in range(n_cli)]
    costs = [[rng.randint(1, 50) for _ in range(n_cli)] for _ in range(n_src)]
    dems  = [rng.randint(50, 200) for _ in range(n_cli)]
    td    = sum(dems)
    caps_raw = [rng.randint(80, 250) for _ in range(n_src)]
    ratio    = td / sum(caps_raw) if sum(caps_raw) < td else 1
    caps     = [max(50, int(c * ratio * 1.15)) for c in caps_raw]
    return make_excel_bytes(costs, srcs, clis, caps, dems)


def solve_transport(costs_arr, capacities, demands, source_names, client_names):
    """Résout le problème de transport avec PuLP/CBC."""
    m, n = len(capacities), len(demands)
    prob = pulp.LpProblem("Transport", pulp.LpMinimize)
    x = [[pulp.LpVariable(f"x_{i}_{j}", lowBound=0) for j in range(n)] for i in range(m)]
    prob += pulp.lpSum(costs_arr[i][j] * x[i][j] for i in range(m) for j in range(n))
    for i in range(m):
        prob += pulp.lpSum(x[i][j] for j in range(n)) <= capacities[i]
    for j in range(n):
        prob += pulp.lpSum(x[i][j] for i in range(m)) == demands[j]
    prob.solve(pulp.PULP_CBC_CMD(msg=0))
    status = pulp.LpStatus[prob.status]
    if prob.status != 1:
        return status, None, None
    Z = pulp.value(prob.objective)
    plan = np.array([[pulp.value(x[i][j]) or 0.0 for j in range(n)] for i in range(m)])
    return status, Z, pd.DataFrame(plan, index=source_names, columns=client_names)


def display_results(Z_opt, plan_df, source_names, client_names, costs_arr, capacities, demands):
    """Affiche tous les résultats après résolution."""
    st.success("✅ Solution optimale trouvée — Statut : **Optimal**")

    # ── KPI ──
    st.markdown('<div class="section-title">📊 Résultats clés</div>', unsafe_allow_html=True)
    total_shipped = float(plan_df.values.sum())
    routes_active = int((plan_df.values > 0.001).sum())
    avg_cost      = Z_opt / total_shipped if total_shipped > 0 else 0
    k1, k2, k3, k4 = st.columns(4)
    k1.markdown(f'<div class="kpi-card"><h3>Coût Total Minimal</h3><p>{Z_opt:,.2f}</p></div>', unsafe_allow_html=True)
    k2.markdown(f'<div class="kpi-card"><h3>Routes Actives</h3><p>{routes_active}</p></div>', unsafe_allow_html=True)
    k3.markdown(f'<div class="kpi-card"><h3>Qté Totale Livrée</h3><p>{total_shipped:,.0f}</p></div>', unsafe_allow_html=True)
    k4.markdown(f'<div class="kpi-card"><h3>Coût Moyen / Unité</h3><p>{avg_cost:,.2f}</p></div>', unsafe_allow_html=True)
    st.markdown("---")

    # ── Plan de transport ──
    st.markdown('<div class="section-title">📋 Plan de Transport Optimal (xᵢⱼ)</div>', unsafe_allow_html=True)
    plan_disp = plan_df.copy()
    plan_disp["✈ Total expédié"] = plan_disp.sum(axis=1)
    totals = plan_disp.sum(axis=0); totals.name = "📦 Total reçu"
    plan_disp = pd.concat([plan_disp, totals.to_frame().T])

    def style_plan(val):
        if not isinstance(val, (int, float, np.floating)): return ""
        if float(val) < 0.001: return "color:#bbb;"
        return "background-color:#e3f2fd;font-weight:600;color:#1565c0;"

    st.dataframe(plan_disp.style.map(style_plan).format("{:.1f}"), use_container_width=True)
    st.markdown("---")

    # ── Graphiques ──
    melt = []
    for i, src in enumerate(source_names):
        for j, cli in enumerate(client_names):
            qty = float(plan_df.loc[src, cli])
            if qty > 0.001:
                melt.append({"Source": src, "Client": cli, "Quantité": qty,
                             "Coût unitaire": costs_arr[i][j],
                             "Coût route": round(qty * costs_arr[i][j], 2)})
    mdf = pd.DataFrame(melt)

    tab1, tab2, tab3 = st.tabs(["🔥 Heatmap des flux", "📊 Bar Chart par source", "📈 Bar Chart par client"])

    with tab1:
        fig = go.Figure(go.Heatmap(
            z=plan_df.values, x=client_names, y=source_names,
            colorscale="Blues", text=np.round(plan_df.values, 1),
            texttemplate="%{text}", textfont={"size": 13},
            colorbar=dict(title="Quantité")))
        fig.update_layout(title="Heatmap des flux xᵢⱼ", xaxis_title="Clients",
                          yaxis_title="Sources", height=420, template="plotly_white")
        st.plotly_chart(fig, use_container_width=True)

    with tab2:
        if not mdf.empty:
            f2 = px.bar(mdf, x="Source", y="Quantité", color="Client", barmode="stack",
                        title="Quantités expédiées par source", template="plotly_white",
                        height=420, color_discrete_sequence=px.colors.qualitative.Bold)
            for i, src in enumerate(source_names):
                f2.add_annotation(x=src, y=capacities[i],
                    text=f"cap={capacities[i]:,.0f}", showarrow=True,
                    arrowhead=2, font=dict(size=10, color="red"), bgcolor="white")
            st.plotly_chart(f2, use_container_width=True)

    with tab3:
        if not mdf.empty:
            f3 = px.bar(mdf, x="Client", y="Quantité", color="Source", barmode="stack",
                        title="Quantités reçues par client", template="plotly_white",
                        height=420, color_discrete_sequence=px.colors.qualitative.Pastel)
            for j, cli in enumerate(client_names):
                f3.add_annotation(x=cli, y=demands[j],
                    text=f"dem={demands[j]:,.0f}", showarrow=True,
                    arrowhead=2, font=dict(size=10, color="green"), bgcolor="white")
            st.plotly_chart(f3, use_container_width=True)

    st.markdown("---")

    # ── Routes actives ──
    st.markdown('<div class="section-title">🗺️ Détail des routes actives</div>', unsafe_allow_html=True)
    if not mdf.empty:
        rd = mdf.copy()
        rd["% coût total"] = (rd["Coût route"] / Z_opt * 100).round(1).astype(str) + "%"
        rd["Route"] = rd["Source"] + " → " + rd["Client"]
        rd = rd[["Route","Source","Client","Quantité","Coût unitaire","Coût route","% coût total"]]
        st.dataframe(rd.sort_values("Coût route", ascending=False).set_index("Route"),
                     use_container_width=True)

    st.markdown("---")

    # ── Export résultats ──
    st.markdown('<div class="section-title">💾 Export des résultats</div>', unsafe_allow_html=True)
    col_e1, col_e2 = st.columns(2)

    buf_res = io.BytesIO()
    with pd.ExcelWriter(buf_res, engine="openpyxl") as w:
        plan_disp.to_excel(w, sheet_name="Plan_Optimal")
        if not mdf.empty:
            rd.to_excel(w, sheet_name="Routes_Actives", index=False)
        pd.DataFrame({
            "Indicateur": ["Coût minimal","Routes actives","Qté livrée","Offre","Demande"],
            "Valeur": [Z_opt, routes_active, total_shipped, sum(capacities), sum(demands)]
        }).to_excel(w, sheet_name="KPI", index=False)

    col_e1.download_button("📊 Résultats (Excel)", data=buf_res.getvalue(),
        file_name="resultats_transport.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True)

    data_excel = make_excel_bytes(costs_arr.tolist(), source_names, client_names,
                                  capacities, demands)
    col_e2.download_button("📥 Données utilisées (Excel)", data=data_excel,
        file_name="donnees_utilisees.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        help="Rechargeable directement en Mode Fichier Excel")
    st.caption("💡 'Données utilisées' = vos coûts/capacités/demandes, "
               "rechargeable en Mode Fichier Excel pour tester d'autres scénarios.")


def display_balance_and_solve(source_names, client_names, costs_arr, capacities, demands):
    """Affiche l'équilibre et le bouton de résolution."""
    total_supply = sum(capacities)
    total_demand = sum(demands)

    st.markdown('<div class="section-title">⚖️ Vérification de l\'équilibre</div>',
                unsafe_allow_html=True)
    e1, e2, e3 = st.columns(3)
    e1.markdown(f'<div class="kpi-card"><h3>Offre Totale</h3><p>{total_supply:,.0f}</p></div>',
                unsafe_allow_html=True)
    e2.markdown(f'<div class="kpi-card"><h3>Demande Totale</h3><p>{total_demand:,.0f}</p></div>',
                unsafe_allow_html=True)
    diff = total_supply - total_demand
    if   diff == 0: badge = '<span class="badge-ok">✅ Équilibré (Δ = 0)</span>'
    elif diff  > 0: badge = f'<span class="badge-warn">⚠️ Excédent offre : +{diff:,.0f}</span>'
    else:           badge = f'<span class="badge-err">❌ Déficit offre : {diff:,.0f} — Infaisable</span>'
    e3.markdown(f"**Statut :** {badge}", unsafe_allow_html=True)
    e3.caption("Si Offre < Demande → problème infaisable")

    st.markdown("---")

    if st.button("🚀 Résoudre le problème", type="primary", use_container_width=True):
        if total_supply < total_demand:
            st.error("❌ Infaisable : Offre totale < Demande totale. Augmentez les capacités.")
        else:
            with st.spinner("Résolution en cours avec PuLP / CBC…"):
                status, Z_opt, plan_df = solve_transport(
                    costs_arr, capacities, demands, source_names, client_names)
            if status != "Optimal":
                st.error(f"❌ Statut solveur : **{status}**. Vérifiez vos données.")
            else:
                display_results(Z_opt, plan_df, source_names, client_names,
                                costs_arr, capacities, demands)


# ══════════════════════════════════════════════════════════════
#  SIDEBAR
# ══════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("## ⚙️ Configuration")
    mode = st.radio("Mode de saisie", ["✏️ Manuel", "📁 Fichier Excel"], index=0)

    st.markdown("---")
    st.markdown("### 📥 Générateur de fichier exemple")
    st.caption("Données aléatoires cohérentes — taille personnalisable")
    c1, c2 = st.columns(2)
    ex_ns = int(c1.number_input("Sources", 2, 8, 3, key="ex_ns"))
    ex_nc = int(c2.number_input("Clients", 2, 8, 4, key="ex_nc"))
    st.download_button(
        label="⬇️ Télécharger Excel Exemple",
        data=generate_random_excel(ex_ns, ex_nc),
        file_name=f"exemple_{ex_ns}x{ex_nc}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        help=f"Génère un fichier avec {ex_ns} sources et {ex_nc} clients"
    )

    st.markdown("---")
    st.markdown("""**Modèle mathématique**
```
min  Σᵢ Σⱼ cᵢⱼ · xᵢⱼ

s.t. Σⱼ xᵢⱼ ≤ sᵢ  ∀i
     Σᵢ xᵢⱼ = dⱼ  ∀j
         xᵢⱼ ≥ 0
```
*Solveur : PuLP / CBC*""")


# ══════════════════════════════════════════════════════════════
#  HEADER
# ══════════════════════════════════════════════════════════════
st.markdown("""<div class="main-header">
  <h1>🚛 Problème de Transport</h1>
  <p>Optimisation du coût de distribution — Programmation Linéaire (PuLP)</p>
</div>""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════
#  MODE MANUEL
# ══════════════════════════════════════════════════════════════
if mode == "✏️ Manuel":
    st.markdown('<div class="section-title">📝 Saisie manuelle des données</div>',
                unsafe_allow_html=True)

    co1, co2 = st.columns(2)
    n_src = int(co1.number_input("Nombre de sources (usines)", 2, 10, 3, step=1))
    n_cli = int(co2.number_input("Nombre de clients", 2, 10, 4, step=1))

    src_names = [f"Source {i+1}" for i in range(n_src)]
    cli_names = [f"Client {j+1}" for j in range(n_cli)]

    # Reset valeurs par défaut si dimensions changent
    dim_key = (n_src, n_cli)
    if st.session_state.get("prev_dim") != dim_key:
        rng2 = random.Random(99)
        st.session_state["def_costs"] = {(i,j): rng2.randint(1,30)
                                          for i in range(n_src) for j in range(n_cli)}
        st.session_state["def_caps"]  = {i: random.randint(150,300) for i in range(n_src)}
        st.session_state["def_dem"]   = {j: random.randint(80,150)  for j in range(n_cli)}
        st.session_state["prev_dim"]  = dim_key

    # ── Matrice des coûts ──
    st.markdown("#### 💰 Matrice des coûts `c_ij`")
    st.caption("Coût unitaire d'envoi de la Source i vers le Client j")

    hdr = st.columns([1.5] + [1]*n_cli)
    hdr[0].markdown("**Source \\ Client**")
    for j, cli in enumerate(cli_names):
        hdr[j+1].markdown(f"**{cli}**")

    cost_vals = []
    for i, src in enumerate(src_names):
        rc = st.columns([1.5] + [1]*n_cli)
        rc[0].markdown(f"**{src}**")
        row = []
        for j in range(n_cli):
            v = rc[j+1].number_input("", min_value=0, max_value=9999,
                value=st.session_state["def_costs"][(i,j)],
                key=f"cost_{i}_{j}", label_visibility="collapsed")
            row.append(int(v))
        cost_vals.append(row)

    # ── Capacités & Demandes ──
    st.markdown("#### 🏭 Capacités  |  📦 Demandes")
    cs, cd = st.columns(2)

    cap_vals = []
    with cs:
        st.caption("Capacité de chaque source (sᵢ)")
        for i, src in enumerate(src_names):
            cap_vals.append(int(st.number_input(src, 1, 999999,
                value=st.session_state["def_caps"][i], key=f"cap_{i}")))

    dem_vals = []
    with cd:
        st.caption("Demande de chaque client (dⱼ)")
        for j, cli in enumerate(cli_names):
            dem_vals.append(int(st.number_input(cli, 1, 999999,
                value=st.session_state["def_dem"][j], key=f"dem_{j}")))

    # ── Export Excel des données manuelles ──
    st.markdown("---")
    st.markdown('<div class="section-title">📤 Exporter vos données saisies en Excel</div>',
                unsafe_allow_html=True)
    st.info(
        "Générez un fichier Excel à partir de vos données ci-dessus. "
        "Vous pourrez le modifier librement puis le recharger en **Mode Fichier Excel** "
        "sans avoir à tout ressaisir."
    )
    manual_excel = make_excel_bytes(cost_vals, src_names, cli_names, cap_vals, dem_vals)
    st.download_button(
        label="📥 Télécharger mes données (Excel)",
        data=manual_excel,
        file_name="mes_donnees_transport.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        help="Exporte vos coûts, capacités et demandes — rechargeable en Mode Fichier Excel"
    )
    st.caption("💡 Astuce : modifiez ce fichier dans Excel, puis rechargez-le en "
               "**Mode Fichier Excel** pour tester différents scénarios rapidement.")

    # ── Résolution ──
    st.markdown("---")
    display_balance_and_solve(
        src_names, cli_names,
        np.array(cost_vals, dtype=float),
        cap_vals, dem_vals
    )


# ══════════════════════════════════════════════════════════════
#  MODE FICHIER EXCEL
# ══════════════════════════════════════════════════════════════
else:
    st.markdown('<div class="section-title">📁 Import du fichier Excel</div>',
                unsafe_allow_html=True)

    # Instructions claires avec aperçu du format
    with st.expander("📋 Format attendu du fichier Excel — cliquez pour voir", expanded=True):
        st.markdown("""
Le fichier doit contenir **exactement 3 feuilles** nommées :

| Feuille | Contenu | Exemple |
|---------|---------|---------|
| **`Couts`** | Matrice des coûts : sources en lignes, clients en colonnes | Ligne 1 = Source 1, Col A = Client 1 |
| **`Capacites`** | Une colonne `Capacite` avec la capacité de chaque source | Source 1 → 200 |
| **`Demandes`** | Une colonne `Demande` avec la demande de chaque client | Client 1 → 120 |

👉 **Téléchargez le fichier exemple** dans la barre latérale pour voir exactement le format,
ou **exportez vos données manuelles** depuis le Mode Manuel.
        """)

    st.markdown("---")

    # Zone d'upload mise en valeur
    st.markdown('<div class="upload-box">', unsafe_allow_html=True)
    uploaded = st.file_uploader(
        "📂 Glissez-déposez votre fichier Excel ici, ou cliquez pour parcourir",
        type=["xlsx"],
        help="Fichier .xlsx avec les feuilles Couts, Capacites, Demandes"
    )
    st.markdown('</div>', unsafe_allow_html=True)

    if uploaded is not None:
        try:
            xls      = pd.ExcelFile(uploaded)

            # Vérification des feuilles
            required = {"Couts", "Capacites", "Demandes"}
            missing  = required - set(xls.sheet_names)
            if missing:
                st.error(f"❌ Feuilles manquantes dans le fichier : **{', '.join(missing)}**\n\n"
                         f"Feuilles trouvées : {', '.join(xls.sheet_names)}")
                st.stop()

            costs_df = pd.read_excel(xls, sheet_name="Couts",     index_col=0)
            caps_df  = pd.read_excel(xls, sheet_name="Capacites", index_col=0)
            dem_df   = pd.read_excel(xls, sheet_name="Demandes",  index_col=0)

            source_names = list(costs_df.index.astype(str))
            client_names = list(costs_df.columns.astype(str))
            costs_arr    = costs_df.values.astype(float)
            capacities   = [float(v) for v in caps_df.iloc[:, 0].tolist()]
            demands      = [float(v) for v in dem_df.iloc[:, 0].tolist()]

            # Vérification cohérence dimensions
            if len(capacities) != len(source_names):
                st.error(f"❌ Incohérence : {len(source_names)} sources dans Couts "
                         f"mais {len(capacities)} lignes dans Capacites.")
                st.stop()
            if len(demands) != len(client_names):
                st.error(f"❌ Incohérence : {len(client_names)} clients dans Couts "
                         f"mais {len(demands)} lignes dans Demandes.")
                st.stop()

            st.success(f"✅ Fichier chargé avec succès — "
                       f"**{len(source_names)} sources** × **{len(client_names)} clients**")

            # Aperçu des données
            with st.expander("🔍 Aperçu des données importées"):
                st.markdown("**Matrice des coûts (Couts)**")
                st.dataframe(costs_df,
                             use_container_width=True)
                cc1, cc2 = st.columns(2)
                with cc1:
                    st.markdown("**Capacités des sources**")
                    st.dataframe(caps_df, use_container_width=True)
                with cc2:
                    st.markdown("**Demandes des clients**")
                    st.dataframe(dem_df, use_container_width=True)

            # ── Résolution ──
            st.markdown("---")
            display_balance_and_solve(
                source_names, client_names,
                costs_arr, capacities, demands
            )

        except Exception as e:
            st.error(f"❌ Erreur lors de la lecture du fichier : **{e}**\n\n"
                     "Vérifiez que le fichier respecte le format attendu.")


# ── FOOTER ──
st.markdown("---")
st.markdown("<div style='text-align:center;color:#9e9e9e;font-size:.85rem;'>"
            "Mini-projet Recherche Opérationnelle · Problème de Transport · "
            "Solveur <b>PuLP/CBC</b> · Interface <b>Streamlit</b></div>",
            unsafe_allow_html=True)
