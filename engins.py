import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO
from prophet import Prophet
import numpy as np
from sklearn.metrics import mean_absolute_error
from statsmodels.tsa.arima.model import ARIMA
from statsmodels.tools.sm_exceptions import ValueWarning
import warnings

# Configuration de la page
st.set_page_config(layout="wide", page_title="Analyse des Engins R1600")


# Charger les données
@st.cache_data
def load_data():
    df = pd.read_excel('engins2.xlsx', sheet_name='BASE DE DONNEE')
    
    # Nettoyage des données
    df['Desc_CA'] = df['Desc_CA'].str.replace('', '').str.strip()
    df['Desc_Cat'] = df['Desc_Cat'].str.strip()
    df['Montant'] = pd.to_numeric(df['Montant'], errors='coerce')
    df['Date'] = pd.to_datetime(df['Date'])
    
    # Liste des engins spécifiques à analyser
    engins_cibles = [
        'CHARGEUSE CATERPILLAR 10T   R1600 Nｰ14 DS',
        'CHARGEUSE CATERPILAR R 1600H Nｰ15 DS',
        'Chargeuse CATERPILAR 10T  R1600 Nｰ16',
        'Chargeuse Caterpillar 10T R1600 Nｰ17',
        'Chargeuse  CAT    R1600 10T  Nｰ18',                
        'CHARGEUSE CATERPILAR R 1600 Nｰ20',
        'CHARGEUSE CATERPILAR R 1600 Nｰ21',
        'CHARGEUSE CATERPILLARD R1600 Nｰ22',
        'CHARGEUSE CATERPILLARD R1600 Nｰ23'
    ]
    
    # Filtrer uniquement les engins cibles
    df = df[df['Desc_CA'].isin(engins_cibles)].copy()
    
    # Extraire le numéro de l'engin et créer un nom standardisé
    df['Numéro_Engin'] = df['Desc_CA'].str.extract(r'Nｰ(\d+)')
    df['Engin_Formaté'] = 'R1600-' + df['Numéro_Engin']
    
    # Mois en français
    months_fr = {
        'January': 'Janvier', 'February': 'Février', 'March': 'Mars',
        'April': 'Avril', 'May': 'Mai', 'June': 'Juin',
        'July': 'Juillet', 'August': 'Août', 'September': 'Septembre',
        'October': 'Octobre', 'November': 'Novembre', 'December': 'Décembre'
    }
    df['Mois'] = df['Date'].dt.month_name().map(months_fr)
    
    return df

# Chargement des données
df = load_data()

# =============================================
# SECTION 1 : EN-TÊTE AVEC KPI PRINCIPAUX
# =============================================
st.markdown("<h1 style='color:#F28C38;'>📊 Analyse des Engins R1600</h1>", unsafe_allow_html=True)

kpi1, kpi2, kpi3, kpi4 = st.columns(4)
with kpi1:
    st.metric("Coût total", f"{df['Montant'].sum():,.0f} MAD", 
             help="Somme totale des dépenses pour tous les engins")
with kpi2:
    top_engine = df.groupby('Engin_Formaté')['Montant'].sum().idxmax()
    st.metric("Engin le plus coûteux", top_engine.split('-')[-1],
             help="Numéro de l'engin avec le coût total le plus élevé")
with kpi3:
    top_category = df.groupby('Desc_Cat')['Montant'].sum().idxmax()
    st.metric("Catégorie principale", top_category,
             help="Catégorie de dépense la plus importante")
with kpi4:
    avg_cost = df['Montant'].mean()
    st.metric("Coût moyen par intervention", f"{avg_cost:,.0f} MAD",
             help="Moyenne des montants des interventions")

# =============================================
# SECTION 2 : FILTRES (SIDEBAR)
# =============================================
st.sidebar.markdown("<h2 style='color:#F28C38;'>🔍 Filtres</h2>", unsafe_allow_html=True)

selected_engin = st.sidebar.radio(
    'Sélectionner un engin',
    ['Tous'] + sorted(df['Engin_Formaté'].unique().tolist()),
    help="Filtrer par engin spécifique"
)

# Filtre par mois
selected_month = st.sidebar.multiselect(
    'Mois', 
    options=sorted(df['Mois'].unique()),
    default=sorted(df['Mois'].unique()),
    help="Sélectionner un ou plusieurs mois"
)

# Appliquer les filtres
filtered_data = df.copy()
if selected_engin != 'Tous':
    filtered_data = filtered_data[filtered_data['Engin_Formaté'] == selected_engin]

if selected_month:
    filtered_data = filtered_data[filtered_data['Mois'].isin(selected_month)]

# Statistiques sidebar
st.sidebar.markdown("<h2 style='color:#F28C38;'>📈 Statistiques filtres</h2>", unsafe_allow_html=True)
st.sidebar.metric("Consommation totale", f"{filtered_data['Montant'].sum():,.0f} MAD")
st.sidebar.metric("Nombre d'interventions", len(filtered_data))
st.sidebar.metric("Coût moyen", f"{filtered_data['Montant'].mean():,.0f} MAD")

# =============================================
# SECTION 3 : ONGLETS PRINCIPAUX
# =============================================
tab1, tab2, tab3, tab4 = st.tabs(["📋 Données", "📊 Analyse", "🔄 Comparaisons", "💡 Recommandations"])

# Onglet 1 : Données
with tab1:
    st.markdown("<h2 style='color:#F28C38;'>Données détaillées des consommations</h2>", unsafe_allow_html=True)
    
    # Afficher un échantillon avec possibilité de tout voir
    if st.checkbox("Afficher toutes les données (défilement)"):
        height = 600
    else:
        height = 300
    
    st.dataframe(
        filtered_data[['Date', 'Engin_Formaté', 'Desc_Cat', 'Montant', 'Mois']]
        .sort_values(['Engin_Formaté', 'Date'])
        .style.format({
            'Montant': '{:,.0f} MAD',
            'Date': lambda x: x.strftime('%d/%m/%Y') if not pd.isnull(x) else ''
        }),
        height=height,
        use_container_width=True
    )
    
    # Export des données
    st.markdown("<h3 style='color:#F28C38;'>Exporter les données</h3>", unsafe_allow_html=True)
    export_col1, export_col2 = st.columns([1, 3])
    with export_col1:
        export_format = st.radio("Format", ['CSV', 'Excel'])
    with export_col2:
        if export_format == 'CSV':
            csv = filtered_data.to_csv(index=False).encode('utf-8')
            st.download_button(
                label="📥 Télécharger CSV",
                data=csv,
                file_name='consommation_r1600.csv',
                mime='text/csv'
            )
        else:
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                filtered_data.to_excel(writer, index=False)
            st.download_button(
                label="📥 Télécharger Excel",
                data=output.getvalue(),
                file_name='consommation_r1600.xlsx',
                mime='application/vnd.ms-excel'
            )

# Onglet 2 : Analyse
with tab2:
    if selected_engin == 'Tous':
        selected = st.selectbox('Choisir un engin à analyser', sorted(df['Engin_Formaté'].unique()))
        engin_data = df[df['Engin_Formaté'] == selected]
    else:
        engin_data = filtered_data
    
    st.markdown(f"<h2 style='color:#F28C38;'>Analyse pour {engin_data['Engin_Formaté'].iloc[0].split('-')[-1]}</h2>", unsafe_allow_html=True)
    
    # Colonnes principales
    col1, col2 = st.columns([7, 3])
    
    with col1:
        # Graphique d'évolution temporelle avec projection
        fig = px.line(
            engin_data.groupby('Date')['Montant'].sum().reset_index(),
            x='Date', y='Montant',
            title='Évolution des dépenses avec projection',
            height=400
        )
        
        # Ajout de la projection si assez de données
        if len(engin_data) >= 3:
            dates = engin_data.groupby('Date')['Montant'].sum().index
            x = np.arange(len(dates))
            y = engin_data.groupby('Date')['Montant'].sum().values
            coeff = np.polyfit(x, y, 1)
            future_dates = [dates[-1] + pd.DateOffset(months=i) for i in range(1,4)]
            projection = np.polyval(coeff, [x[-1]+1, x[-1]+2, x[-1]+3])
            
            fig.add_scatter(
                x=future_dates,
                y=projection,
                mode='lines+markers',
                name='Projection (linéaire)',
                line=dict(color='red', dash='dot')
            )
            
            # Ajout des fourchettes
            fig.add_scatter(
                x=future_dates,
                y=projection * 1.3,
                mode='lines',
                name='Fourchette haute (+30%)',
                line=dict(color='orange', dash='dash')
            )
            
            fig.add_scatter(
                x=future_dates,
                y=projection * 0.7,
                mode='lines',
                name='Fourchette basse (-30%)',
                line=dict(color='green', dash='dash')
            )
        
        st.plotly_chart(fig, use_container_width=True)
        
        # Mini-graphiques d'analyse
        subcols = st.columns(3)
        with subcols[0]:
            st.plotly_chart(
                px.pie(
                    engin_data.groupby('Desc_Cat')['Montant'].sum().reset_index(),
                    values='Montant', names='Desc_Cat',
                    title='Répartition par catégorie',
                    height=250
                ),
                use_container_width=True
            )
        
        with subcols[1]:
            st.plotly_chart(
                px.bar(
                    engin_data.groupby('Mois')['Montant'].sum().reset_index(),
                    x='Mois', y='Montant',
                    title='Coût mensuel',
                    height=250
                ),
                use_container_width=True
            )
        
        with subcols[2]:
            st.plotly_chart(
                px.histogram(
                    engin_data, x='Montant',
                    title='Distribution des coûts',
                    height=250
                ),
                use_container_width=True
            )
    
    with col2:
        # Panneau de prévision compact
        st.markdown("""
        <div style='background-color:#f8f9fa; padding:15px; border-radius:10px; border-left:4px solid #4285f4; margin-bottom:20px;'>
            <h3 style='color:#F28C38; margin-top:0;'>🔮 Prévision</h3>
        </div>
        """, unsafe_allow_html=True)
        
        if len(engin_data) >= 3:
            last_3 = engin_data.groupby('Mois')['Montant'].sum().tail(3)
            avg = last_3.mean()
            
            st.metric("Estimation moyenne", f"{avg:,.0f} MAD/mois")
            st.metric("Fourchette probable", 
                      f"{last_3.min():,.0f}-{last_3.max():,.0f} MAD")
            
            st.progress(65, "Fiabilité des prévisions")
            st.caption("Basé sur les 3 derniers mois")
            
            # Bouton d'export
            if st.button("💾 Exporter prévisions", key="export_btn"):
                future_dates = pd.date_range(
                    start=engin_data['Date'].max() + pd.DateOffset(months=1),
                    periods=3,
                    freq='M'
                )
                projections = pd.DataFrame({
                    'Mois': future_dates.strftime('%Y-%m'),
                    'Estimation': [avg]*3,
                    'Minimum': [last_3.min()]*3,
                    'Maximum': [last_3.max()]*3
                })
                
                csv = projections.to_csv(index=False).encode('utf-8')
                st.download_button(
                    label="Télécharger CSV",
                    data=csv,
                    file_name='previsions_engin.csv',
                    mime='text/csv'
                )
        else:
            st.warning("Données insuffisantes (minimum 3 mois requis)")
        
        st.markdown("</div>", unsafe_allow_html=True)
        
        # Panneau d'indicateurs clés
        st.markdown("""
        <div style='background-color:#000000; padding:15px; border-radius:10px; border-left:4px solid #4285f4; margin-bottom:20px;'>
            <h3 style='color:#F28C38; margin-top:0;'>📊 Indicateurs</h3>
            <p>Dernier mois: <strong>{last_month:,.0f} MAD</strong></p>
            <p>Moyenne 3m: <strong>{avg_3m:,.0f} MAD</strong></p>
            <p>Maximum: <strong>{max_m:,.0f} MAD</strong></p>
            <p>Coût total: <strong>{total:,.0f} MAD</strong></p>
        </div>
        """.format(
            last_month=engin_data.groupby('Mois')['Montant'].sum().iloc[-1],
            avg_3m=engin_data.groupby('Mois')['Montant'].sum().tail(3).mean(),
            max_m=engin_data['Montant'].max(),
            total=engin_data['Montant'].sum()
        ), unsafe_allow_html=True)
    
    # Section d'analyse
    st.markdown("""
    **📝 Analyse des tendances :**
    - La projection (ligne rouge) montre l'évolution estimée
    - Les fourchettes donnent une plage de valeurs probables
    - Les données limitées réduisent la précision des prévisions
    
    **💡 Recommandations :**
    1. Surveiller particulièrement les catégories majoritaires
    2. Analyser les causes des pics de dépenses
    3. Collecter plus de données pour améliorer les prévisions
    """)

# Onglet 3 : Comparaisons
with tab3:
    st.markdown("<h2 style='color:#F28C38;'>Comparaisons entre engins</h2>", unsafe_allow_html=True)
    
    # Heatmap compacte
    st.plotly_chart(
        px.imshow(
            df.pivot_table(index='Engin_Formaté', columns='Desc_Cat', values='Montant', aggfunc='sum'),
            labels=dict(x="Catégorie", y="Engin", color="Coût"),
            title='Heatmap des coûts par engin et catégorie',
            aspect="auto",
            height=400
        ),
        use_container_width=True
    )
    
    # Graphiques de comparaison côte à côte
    col1, col2 = st.columns(2)
    with col1:
        st.plotly_chart(
            px.bar(
                df.groupby('Engin_Formaté')['Montant'].sum().reset_index().sort_values('Montant'),
                x='Montant', y='Engin_Formaté',
                title='Coût total par engin',
                height=400
            ),
            use_container_width=True
        )
    with col2:
        st.plotly_chart(
            px.box(
                df, x='Engin_Formaté', y='Montant',
                title='Distribution des coûts par engin',
                height=400
            ),
            use_container_width=True
        )
    st.markdown("""
    **📝 Analyse de la heatmap :**  
    - Les cases chaudes (rouges) révèlent des combinaisons engin/catégorie problématiques  
    - Patterns verticaux = problèmes communs à plusieurs engins  
    - Patterns horizontaux = engins particulièrement coûteux
    """)
    
    st.markdown("""
    **📝 Analyse des boîtes à moustaches :**  
    - Médiane élevée = coût de base important  
    - Longues moustaches = grande variabilité des coûts  
    - Points isolés = interventions exceptionnellement coûteuses
    """)

# =============================================
# SECTION RECOMMANDATIONS SIMPLIFIÉE (MAD)
# =============================================
with tab4:
    st.markdown("<h2 style='color:#F28C38;'>🚀 Plan d'Action Simplifié</h2>", unsafe_allow_html=True, help="Recommandations pratiques pour réduire les coûts")
    
    # Cartes visuelles avec recommandations
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown(f"""
        <div style='background-color:#e3f2fd; padding:20px; border-radius:10px; margin-bottom:20px; border-left:5px solid #1976d2;'>
            <h3 style='color:#F28C38; margin-top:0;'>🔍 Top 3 des Dépenses à Surveiller</h3>
            <ol style='color:#424242;'>
                <li style='margin-bottom:10px;'><b>{top_category}</b> <span style='color:#d32f2f; font-weight:bold;'>{df[df['Desc_Cat']==top_category]['Montant'].sum():,.0f} MAD</span></li>
                <li style='margin-bottom:10px;'><b>{df.groupby('Desc_Cat')['Montant'].sum().nlargest(2).index[1]}</b> <span style='color:#d32f2f; font-weight:bold;'>{df.groupby('Desc_Cat')['Montant'].sum().nlargest(2).values[1]:,.0f} MAD</span></li>
                <li><b>{df.groupby('Desc_Cat')['Montant'].sum().nlargest(3).index[2]}</b> <span style='color:#d32f2f; font-weight:bold;'>{df.groupby('Desc_Cat')['Montant'].sum().nlargest(3).values[2]:,.0f} MAD</span></li>
            </ol>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        problem_engines = df.groupby('Engin_Formaté')['Montant'].sum().nlargest(3)
        st.markdown(f"""
        <div style='background-color:#fff8e1; padding:20px; border-radius:10px; margin-bottom:20px; border-left:5px solid #ffa000;'>
            <h3 style='color:#F28C38; margin-top:0;'>🚜 Engins Prioritaires</h3>
            <div style='display:flex; align-items:center; margin-bottom:10px; color:#424242;'>
                <div style='background-color:#d32f2f; width:20px; height:20px; border-radius:50%; margin-right:10px;'></div>
                <div>Engin R1600-{problem_engines.index[0].split('-')[-1]} <span style='color:#d32f2f; font-weight:bold;'>- {problem_engines.values[0]:,.0f} MAD</span></div>
            </div>
            <div style='display:flex; align-items:center; margin-bottom:10px; color:#424242;'>
                <div style='background-color:#ffa000; width:20px; height:20px; border-radius:50%; margin-right:10px;'></div>
                <div>Engin R1600-{problem_engines.index[1].split('-')[-1]} <span style='color:#d32f2f; font-weight:bold;'>- {problem_engines.values[1]:,.0f} MAD</span></div>
            </div>
            <div style='display:flex; align-items:center; color:#424242;'>
                <div style='background-color:#fbc02d; width:20px; height:20px; border-radius:50%; margin-right:10px;'></div>
                <div>Engin R1600-{problem_engines.index[2].split('-')[-1]} <span style='color:#d32f2f; font-weight:bold;'>- {problem_engines.values[2]:,.0f} MAD</span></div>
            </div>
        </div>
        """, unsafe_allow_html=True)
    
    # Timeline d'actions
    st.markdown("""
    <div style='background-color:#e8f5e9; padding:20px; border-radius:10px; margin-bottom:20px; border-left:5px solid #388e3c;'>
        <h3 style='color:#F28C38; margin-top:0;'>📅 Plan d'Action sur 3 Mois</h3>
        <div style='border-left:4px solid #388e3c; padding-left:20px;'>
            <div style='margin-bottom:15px; color:#424242;'>
                <h4 style='margin-bottom:5px; color:#2e7d32;'>Mois 1 : Audit Initial</h4>
                <p>• Vérifier les 3 engins les plus coûteux<br>• Analyser les pièces les plus remplacées</p>
            </div>
            <div style='margin-bottom:15px; color:#424242;'>
                <h4 style='margin-bottom:5px; color:#2e7d32;'>Mois 2 : Négociations</h4>
                <p>• Contacter fournisseurs pour remises volume<br>• Standardiser les pièces communes</p>
            </div>
            <div style='color:#424242;'>
                <h4 style='margin-bottom:5px; color:#2e7d32;'>Mois 3 : Optimisation</h4>
                <p>• Mettre en place maintenance préventive<br>• Former les opérateurs aux bonnes pratiques</p>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    # Checklist interactive
    st.markdown("""
    <div style='background-color:#f3e5f5; padding:20px; border-radius:10px; margin-bottom:20px; border-left:5px solid #8e24aa;'>
        <h3 style='color:#F28C38; margin-top:0;'>✅ Checklist des Actions Clés</h3>
        <div style='margin-bottom:10px; color:#424242;'><input type='checkbox' style='margin-right:10px;'> Identifier les 5 pièces les plus changées</div>
        <div style='margin-bottom:10px; color:#424242;'><input type='checkbox' style='margin-right:10px;'> Comparer les coûts avec les standards du marché (en MAD)</div>
        <div style='margin-bottom:10px; color:#424242;'><input type='checkbox' style='margin-right:10px;'> Réaliser un diagnostic des engins prioritaires</div>
        <div style='margin-bottom:10px; color:#424242;'><input type='checkbox' style='margin-right:10px;'> Organiser une réunion avec les fournisseurs</div>
        <div style='color:#424242;'><input type='checkbox' style='margin-right:10px;'> Mettre en place un suivi mensuel des coûts (MAD)</div>
    </div>
    """, unsafe_allow_html=True)
    
    # Conseils pratiques
    st.markdown(f"""
    <div style='background-color:#ffebee; padding:20px; border-radius:10px; border-left:5px solid #d32f2f;'>
        <h3 style='color:#F28C38; text-align:center; margin-top:0;'>💡 3 Astuces pour Réduire les Coûts (MAD)</h3>
        <div style='display:flex; justify-content:space-between; text-align:center; margin-top:20px;'>
            <div style='width:30%;'>
                <div style='background-color:#ffcdd2; padding:15px; border-radius:10px; height:120px;'>
                    <h4 style='color:#c62828;'>🛠️ Maintenance</h4>
                    <p style='color:#424242;'>Économie potentielle :<br><b>{df['Montant'].mean()*0.25:,.0f} MAD/mois</b></p>
                </div>
            </div>
            <div style='width:30%;'>
                <div style='background-color:#c8e6c9; padding:15px; border-radius:10px; height:120px;'>
                    <h4 style='color:#2e7d32;'>🔄 Pièces Standard</h4>
                    <p style='color:#424242;'>Économie potentielle :<br><b>{df['Montant'].mean()*0.40:,.0f} MAD/mois</b></p>
                </div>
            </div>
            <div style='width:30%;'>
                <div style='background-color:#bbdefb; padding:15px; border-radius:10px; height:120px;'>
                    <h4 style='color:#1565c0;'>📊 Suivi Rigoureux</h4>
                    <p style='color:#424242;'>Économie potentielle :<br><b>{df['Montant'].mean()*0.15:,.0f} MAD/mois</b></p>
                </div>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)