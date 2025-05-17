import streamlit as st
import pandas as pd
import plotly.express as px
import numpy as np
from datetime import datetime
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from io import BytesIO
import plotly.io as pio
import json
import bcrypt
import os

# Configuration de la page
st.set_page_config(page_title="Tableau de bord de la consommation des équipements miniers", layout="wide")
st.markdown("""
    <style>
    .stApp { 
        background-image: url("https://img.freepik.com/premium-photo/underground-mining-truck_873668-11862.jpg"); 
        background-size: cover; 
        background-repeat: no-repeat; 
    }
    .stApp > div { 
        padding: 20px; 
        border-radius: 10px; 
    }
    h1, h2, h3 { 
        color: #003087; 
        font-family: Arial, sans-serif; 
    }
    .stMetric { 
        background-color: #Ff7f00; 
        border-left: 5px solid #FFC107; 
        padding: 10px; 
        border-radius: 5px; 
    }
    .stButton>button { 
        background-color: #003087; 
        color: white; 
        border-radius: 5px; 
    }
    .stButton>button:hover { 
        background-color: #FFC107; 
        color: black; 
    }
    </style>
""", unsafe_allow_html=True)

# Fonctions pour gérer les utilisateurs
def load_users():
    if os.path.exists("users.json"):
        with open("users.json", "r") as f:
            return json.load(f)
    return {}

def save_users(users):
    with open("users.json", "w") as f:
        json.dump(users, f)

def hash_password(password):
    return bcrypt.hashpw(password.encode('utf-8'), bcrypt.gensalt()).decode('utf-8')

def check_password(password, hashed):
    return bcrypt.checkpw(password.encode('utf-8'), hashed.encode('utf-8'))

# Chargement des données
@st.cache_data
def load_data():
    df = pd.read_excel("engins2.xlsx")
    if not all(col in df.columns for col in ['Date', 'CATEGORIE', 'Desc_Cat', 'Desc_CA', 'Montant']):
        st.error("Fichier engins2.xlsx manquant des colonnes requises.")
        st.stop()
    if pd.api.types.is_numeric_dtype(df['Date']):
        df['Date'] = pd.to_datetime(df['Date'], origin='1899-12-30', unit='D')
    elif not pd.api.types.is_datetime64_any_dtype(df['Date']):
        df['Date'] = pd.to_datetime(df['Date'])
    df = df.dropna(subset=['CATEGORIE', 'Desc_Cat', 'Desc_CA', 'Montant'])
    df['Montant'] = pd.to_numeric(df['Montant'], errors='coerce')
    df['Mois'] = df['Date'].dt.month_name()
    months_fr = {
        'January': 'Janvier', 'February': 'Février', 'March': 'Mars',
        'April': 'Avril', 'May': 'Mai', 'June': 'Juin',
        'July': 'Juillet', 'August': 'Août', 'September': 'Septembre',
        'October': 'Octobre', 'November': 'Novembre', 'December': 'Décembre'
    }
    df['Mois'] = df['Mois'].map(months_fr)
    return df

# Calculs en cache
@st.cache_data
def compute_monthly_costs(data):
    monthly_data = data.groupby('Mois')['Montant'].sum().reset_index()
    month_order = ['Janvier', 'Février', 'Mars', 'Avril', 'Mai', 'Juin',
                   'Juillet', 'Août', 'Septembre', 'Octobre', 'Novembre', 'Décembre']
    monthly_data['Mois'] = pd.Categorical(monthly_data['Mois'], categories=month_order, ordered=True)
    return monthly_data.sort_values('Mois')

@st.cache_data
def compute_category_breakdown(data):
    return data.groupby('Desc_Cat')['Montant'].sum().reset_index()

# Fonction pour générer le rapport Word
@st.cache_data
def generate_word_report(filtered_data, total_cost, global_avg, category_stats, most_consumed_per_cat, 
                        pivot_engine, selected_engine, table_df, total_montant, figures):
    doc = Document()
    
    # Titre et métadonnées
    title = doc.add_heading('Rapport Complet de Consommation des Équipements Miniers', 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph(f"Date de génération: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    doc.add_paragraph(f"Période couverte: du {filtered_data['Date'].min().strftime('%d/%m/%Y')} au {filtered_data['Date'].max().strftime('%d/%m/%Y')}")
    doc.add_paragraph(f"Nombre d'équipements analysés: {filtered_data['Desc_CA'].nunique()}")
    
    # Table des matières
    doc.add_heading('Table des Matières', level=1)
    doc.add_paragraph('1. Indicateurs Clés\n2. Analyse par Catégorie\n3. Analyse Comparative\n4. Données Détailées\n5. Recommandations', style='ListBullet')
    
    # Section 1: Indicateurs clés
    doc.add_heading('1. Indicateurs Clés', level=1)
    table = doc.add_table(rows=3, cols=2)
    table.style = 'LightShading'
    table.cell(0, 0).text = 'Indicateur'
    table.cell(0, 1).text = 'Valeur'
    table.cell(1, 0).text = 'Coût total'
    table.cell(1, 1).text = f"{total_cost:,.0f} DH"
    table.cell(2, 0).text = 'Moyenne globale par jour'
    table.cell(2, 1).text = f"{global_avg:,.0f} DH"
    
    # Indicateurs par catégorie
    doc.add_heading('Indicateurs par Catégorie', level=2)
    cat_table = doc.add_table(rows=category_stats.shape[0]+1, cols=4)
    cat_table.style = 'LightShading'
    cat_table.cell(0, 0).text = 'Catégorie'
    cat_table.cell(0, 1).text = 'Total (DH)'
    cat_table.cell(0, 2).text = 'Moyenne (DH)'
    cat_table.cell(0, 3).text = 'Type le plus consommé'
    
    for i, (_, row) in enumerate(category_stats.iterrows()):
        most_consumed = most_consumed_per_cat[most_consumed_per_cat['CATEGORIE'] == row['CATEGORIE']]
        most_consumed_desc = most_consumed['Desc_Cat'].iloc[0] if not most_consumed.empty else "N/A"
        
        cat_table.cell(i+1, 0).text = row['CATEGORIE']
        cat_table.cell(i+1, 1).text = f"{row['Total']:,.0f}"
        cat_table.cell(i+1, 2).text = f"{row['Moyenne']:,.0f}"
        cat_table.cell(i+1, 3).text = most_consumed_desc
    
    # Section 2: Graphiques et analyses
    doc.add_heading('2. Analyse par Catégorie', level=1)
    doc.add_paragraph('Cette section présente les analyses détaillées pour chaque catégorie d\'équipement.')
    
    progress_bar = st.progress(0)
    total_figures = sum(1 for fig_name in figures if "Consommation par équipement" in fig_name) + ("Coût total par catégorie" in figures)
    
    for i, (fig_name, fig) in enumerate(figures.items()):
        if "Consommation par équipement" in fig_name:
            doc.add_heading(fig_name, level=2)
            category = fig_name.split('(')[-1].replace(')', '')
            doc.add_paragraph(f"Ce graphique montre la répartition des coûts par équipement pour la catégorie {category}. "
                            "Il permet d'identifier les équipements les plus coûteux à maintenir.")
            
            img_bytes = pio.to_image(fig, format='png', scale=1)
            doc.add_picture(BytesIO(img_bytes), width=Inches(6))
            progress_bar.progress((i + 1) / total_figures)
    
    # Section 3: Analyse comparative
    doc.add_heading('3. Analyse Comparative', level=1)
    doc.add_paragraph('Comparaison des performances entre les différentes catégories d\'équipements.')
    
    if "Coût total par catégorie" in figures:
        doc.add_heading('Comparaison des coûts par catégorie', level=2)
        doc.add_paragraph("Ce graphique compare les coûts totaux entre les différentes catégories d'équipements. "
                        "Les catégories les plus à droite représentent les postes de dépenses les plus importants.")
        
        img_bytes = pio.to_image(figures["Coût total par catégorie"], format='png', scale=1)
        doc.add_picture(BytesIO(img_bytes), width=Inches(6))
        progress_bar.progress(1.0)
    
    # Section 4: Données détaillées
    doc.add_heading('4. Données Détailées', level=1)
    
    # Tableau pivot
    if not pivot_engine.empty:
        doc.add_heading(f'Détail des consommations pour {selected_engine}', level=2)
        doc.add_paragraph(f"Tableau détaillant les différents types de consommation pour chaque équipement de type {selected_engine}.")
        
        table = doc.add_table(rows=pivot_engine.shape[0]+1, cols=pivot_engine.shape[1]+1)
        table.style = 'Table Grid'
        
        table_rows = table.rows
        table_rows[0].cells[0].text = 'Équipement'
        for j, col in enumerate(pivot_engine.columns):
            table_rows[0].cells[j+1].text = str(col)
        
        for i, (index, row) in enumerate(pivot_engine.iterrows()):
            row_cells = table_rows[i+1].cells
            row_cells[0].text = str(index)
            for j, value in enumerate(row):
                row_cells[j+1].text = f"{value:,.2f} DH"
    
    # Tableau complet des équipements (limité à 100 lignes)
    doc.add_heading('Journal complet des consommations', level=2)
    doc.add_paragraph('Liste détaillée des consommations enregistrées (limité aux 100 premières entrées).')
    
    max_rows = min(table_df.shape[0], 100)
    table = doc.add_table(rows=max_rows+2, cols=table_df.shape[1])
    table.style = 'Table Grid'
    
    table_rows = table.rows
    for j, col in enumerate(table_df.columns):
        table_rows[0].cells[j].text = col
    
    for i in range(max_rows):
        row_cells = table_rows[i+1].cells
        for j, value in enumerate(table_df.iloc[i]):
            row_cells[j].text = str(value)
    
    table_rows[max_rows+1].cells[0].text = 'Total'
    table_rows[max_rows+1].cells[table_df.shape[1]-1].text = f"{total_montant:,.2f} DH"
    
    # Section 5: Recommandations
    doc.add_heading('5. Recommandations', level=1)
    
    top_categories = filtered_data.groupby('CATEGORIE')['Montant'].sum().nlargest(3).reset_index()
    doc.add_heading('Catégories prioritaires', level=2)
    for _, row in top_categories.iterrows():
        doc.add_paragraph(
            f"{row['CATEGORIE']}: {row['Montant']:,.0f} DH ({(row['Montant']/total_cost)*100:.1f}% du total)",
            style='ListBullet'
        )
    
    doc.add_heading('Actions recommandées', level=2)
    recommendations = [
        "Prioriser les analyses des équipements dans les catégories les plus coûteuses",
        "Mettre en place un suivi mensuel des consommations par catégorie",
        "Comparer les performances des équipements similaires pour identifier les anomalies",
        "Négocier avec les fournisseurs pour les pièces les plus fréquemment remplacées",
        "Étudier la possibilité de maintenance préventive pour réduire les coûts",
        "Former les opérateurs à une utilisation optimale des équipements"
    ]
    for rec in recommendations:
        doc.add_paragraph(rec, style='ListBullet')
    
    # Conclusion
    doc.add_heading('Conclusion', level=1)
    doc.add_paragraph(
        "Ce rapport fournit une analyse complète des coûts de consommation des équipements miniers. "
        "Les graphiques et tableaux présentés permettent d'identifier les principaux postes de dépenses "
        "et de prendre des décisions éclairées pour optimiser les coûts d'exploitation."
    )
    
    # Pied de page
    section = doc.sections[0]
    footer = section.footer
    footer_para = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    footer_para.text = f"Généré le {datetime.now().strftime('%d/%m/%Y')} - Tableau de bord de consommation des équipements miniers"
    footer_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    st.write("Buffer size:", buffer.getbuffer().nbytes)
    return buffer

# Initialiser l'état de la session
if 'logged_in' not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.username = None
    st.session_state.page = 'login'

# Interface de connexion/inscription
if not st.session_state.logged_in:
    st.markdown("""
    <div style='background-color:#424242; padding:20px; border-radius:10px; border-left:5px solid #1976d2; margin-bottom:20px;'>
        <h1 style='color:#F28C38; text-align:center; margin-top:0;'>Bienvenue</h1>
        <p style='color:#FFFFFF; text-align:center;'>Veuillez vous connecter ou créer un compte pour accéder au tableau de bord</p>
    </div>
    """, unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        if st.button("Connexion", key="show_login"):
            st.session_state.page = 'login'
    with col2:
        if st.button("Inscription", key="show_signup"):
            st.session_state.page = 'signup'

    users = load_users()

    if st.session_state.page == 'login':
        st.subheader("Connexion")
        with st.form("login_form"):
            username = st.text_input("Nom d'utilisateur")
            password = st.text_input("Mot de passe", type="password")
            submit = st.form_submit_button("Se connecter")

            if submit:
                if username in users and check_password(password, users[username]):
                    st.session_state.logged_in = True
                    st.session_state.username = username
                    st.success(f"Connecté en tant que {username}")
                    st.rerun()
                else:
                    st.error("Nom d'utilisateur ou mot de passe incorrect")

    elif st.session_state.page == 'signup':
        st.subheader("Inscription")
        with st.form("signup_form"):
            new_username = st.text_input("Nouveau nom d'utilisateur")
            new_password = st.text_input("Nouveau mot de passe", type="password")
            confirm_password = st.text_input("Confirmer le mot de passe", type="password")
            submit = st.form_submit_button("S'inscrire")

            if submit:
                if new_username in users:
                    st.error("Ce nom d'utilisateur existe déjà")
                elif new_password != confirm_password:
                    st.error("Les mots de passe ne correspondent pas")
                elif not new_username or not new_password:
                    st.error("Veuillez remplir tous les champs")
                else:
                    users[new_username] = hash_password(new_password)
                    save_users(users)
                    st.success("Compte créé avec succès ! Veuillez vous connecter.")
                    st.session_state.page = 'login'
                    st.rerun()

else:
    # Chargement des données
    df = load_data()

    # Barre latérale pour les filtres
    with st.sidebar:
        # Bouton de déconnexion
        if st.button("Déconnexion", key="logout"):
            st.session_state.logged_in = False
            st.session_state.username = None
            st.session_state.page = 'login'
            st.rerun()

        st.subheader("Filtres")
        st.subheader("Plage de dates")
        default_start = df['Date'].min().date()
        default_end = df['Date'].max().date()
        date_range = st.date_input(
            "Période",
            value=(default_start, default_end),
            min_value=default_start,
            max_value=default_end,
            help="Choisir une plage de dates pour filtrer les interventions"
        )
        st.subheader("Rechercher un équipement")
        equipment_search = st.text_input("Entrer le nom de l'équipement (correspondance partielle)", "").strip()
        if equipment_search:
            available_equipment = sorted(df[df['Desc_CA'].str.contains(equipment_search, case=False, na=False)]['Desc_CA'].unique())
        else:
            available_equipment = sorted(df['Desc_CA'].unique())
        equipment_options = ["Tous les équipements"] + available_equipment
        if not available_equipment:
            st.warning("Aucun équipement ne correspond au terme de recherche.")
        selected_equipment = st.selectbox("Sélectionner l'équipement", equipment_options)
        
        st.subheader("Type d'engin")
        engine_types = sorted(df['CATEGORIE'].unique())
        selected_engine = st.selectbox("Sélectionner le type d'engin", engine_types, key="engine_type_select")

    # Filtrer les données
    filtered_data = df.copy()
    if len(date_range) == 2:
        start_date, end_date = date_range
        filtered_data = filtered_data[(filtered_data['Date'].dt.date >= start_date) & 
                                    (filtered_data['Date'].dt.date <= end_date)]

    if selected_equipment != "Tous les équipements":
        filtered_data = filtered_data[filtered_data['Desc_CA'] == selected_equipment]

    if filtered_data.empty:
        st.warning("Aucune donnée disponible après filtrage. Veuillez ajuster les filtres.")
        st.stop()

    # Calculs pour les KPIs
    total_cost = filtered_data['Montant'].sum()
    global_avg = filtered_data['Montant'].mean()
    category_stats = filtered_data.groupby('CATEGORIE').agg(
        Total=('Montant', 'sum'),
        Moyenne=('Montant', 'mean')
    ).reset_index()
    most_consumed_per_cat = filtered_data.groupby(['CATEGORIE', 'Desc_Cat'])['Montant'].sum().reset_index()
    most_consumed_per_cat = most_consumed_per_cat.loc[most_consumed_per_cat.groupby('CATEGORIE')['Montant'].idxmax()]

    # Exportation dans la barre latérale
    with st.sidebar:
        st.subheader("Exportation")
        if st.button("📄 Générer un rapport Word complet"):
            with st.spinner("Génération du rapport en cours..."):
                # Préparer les figures
                figures = {}
                
                # Figure: Comparaison des catégories
                fig_comp = px.bar(
                    filtered_data.groupby('CATEGORIE')['Montant'].sum().reset_index(),
                    x='CATEGORIE',
                    y='Montant',
                    title='Coût total par catégorie',
                    height=400,
                    text='Montant'
                )
                fig_comp.update_traces(
                    texttemplate='%{text:,.0f} DH',
                    textposition='auto'
                )
                fig_comp.update_layout(
                    xaxis_title="Catégorie",
                    yaxis_title="Montant total (DH)",
                    template='plotly_white'
                )
                figures["Coût total par catégorie"] = fig_comp
                
                # Figures par catégorie
                for cat in filtered_data['CATEGORIE'].unique():
                    cat_data = filtered_data[filtered_data['CATEGORIE'] == cat]
                    equip_sum = cat_data.groupby('Desc_CA')['Montant'].sum().reset_index().sort_values('Montant', ascending=False)
                    fig_cat = px.bar(
                        equip_sum,
                        x='Desc_CA',
                        y='Montant',
                        title=f'Consommation par équipement ({cat})',
                        height=400,
                        text='Montant'
                    )
                    fig_cat.update_traces(
                        texttemplate='%{text:,.0f} DH',
                        textposition='auto'
                    )
                    fig_cat.update_layout(
                        xaxis_title="Équipement",
                        yaxis_title="Montant total (DH)",
                        template='plotly_white',
                        xaxis={'categoryorder':'total descending'}
                    )
                    figures[f"Consommation par équipement ({cat})"] = fig_cat
                
                # Préparer la table pivot
                pivot_engine = pd.DataFrame()
                if not filtered_data.empty and selected_engine:
                    pivot_engine = pd.pivot_table(
                        filtered_data[filtered_data['CATEGORIE'] == selected_engine],
                        values='Montant',
                        index='Desc_CA',
                        columns='Desc_Cat',
                        aggfunc='sum',
                        fill_value=0,
                        margins=True,
                        margins_name='Total'
                    ).round(2)
                
                # Préparer le tableau des équipements
                table_df = filtered_data[['Date', 'Desc_CA', 'Desc_Cat', 'Montant']].copy()
                table_df['Date'] = table_df['Date'].dt.strftime('%d/%m/%Y')
                table_df['Montant'] = table_df['Montant'].round(2)
                table_df = table_df.rename(columns={
                    'Date': 'Date',
                    'Desc_CA': 'Équipement',
                    'Desc_Cat': 'Type de consommation',
                    'Montant': 'Montant (DH)'
                })
                total_montant = table_df['Montant (DH)'].sum()
                
                # Générer le rapport
                report = generate_word_report(
                    filtered_data,
                    total_cost,
                    global_avg,
                    category_stats,
                    most_consumed_per_cat,
                    pivot_engine,
                    selected_engine,
                    table_df,
                    total_montant,
                    figures
                )
                
                # Téléchargement
                st.session_state['report_buffer'] = report
                st.write("Filtered data rows:", filtered_data.shape[0])
                st.write("Selected engine:", selected_engine)
                st.success("Rapport généré avec succès!")
        
        if 'report_buffer' in st.session_state:
            st.download_button(
                label="📥 Télécharger le rapport Word",
                data=st.session_state['report_buffer'],
                file_name=f"Rapport_Consommation_{datetime.now().strftime('%Y%m%d_%H%M')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key="download_button"
            )

    st.markdown("""
    <div style='background-color:#424242; padding:20px; border-radius:10px; border-left:5px solid #1976d2; margin-bottom:20px;'>
        <h1 style='color:#F28C38; text-align:center; margin-top:0;'>📊 Tableau De Bord De La Consommation Des Engins</h1>
        <p style='color:#FFFFFF; text-align:center;'>Suivre et optimiser la consommation des équipements</p>
    </div>
    """, unsafe_allow_html=True)

    # Section des indicateurs clés
    kpi_container = st.container()
    with kpi_container:
        st.markdown(f"""
        <div style='background-color:#424242; padding:15px; border-radius:10px; margin-bottom:20px;'>
            <h3 style='color:#F28C38; margin-top:0;'>Indicateurs globaux</h3>
            <div style='display:flex; justify-content:space-between;'>
                <div style='width:48%; background-color:#424242; padding:10px; border-radius:5px; border-left:4px solid #1976d2;'>
                    <p style='color:#FFFFFF; font-size:16px;'><b>Coût total</b></p>
                    <p style='color:#FFFFFF; font-size:24px; font-weight:bold;'>{total_cost:,.0f} DH</p>
                </div>
                <div style='width:48%; background-color:#424242; padding:10px; border-radius:5px; border-left:4px solid #388e3c;'>
                    <p style='color:#FFFFFF; font-size:16px;'><b>Moyenne globale des engins par jour</b></p>
                    <p style='color:#FFFFFF; font-size:24px; font-weight:bold;'>{global_avg:,.0f} DH</p>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        categories = category_stats['CATEGORIE'].unique()
        cols = st.columns(len(categories))

        for idx, (col, (_, row)) in enumerate(zip(cols, category_stats.iterrows())):
            with col:
                most_consumed = most_consumed_per_cat[most_consumed_per_cat['CATEGORIE'] == row['CATEGORIE']]
                most_consumed_desc = most_consumed['Desc_Cat'].iloc[0] if not most_consumed.empty else "Aucune"
                
                st.markdown(f"""
                <div style='background-color:#424242; padding:15px; border-radius:10px; border-left:4px solid #{'1976d2' if idx%2==0 else '388e3c'}; margin-bottom:10px;'>
                    <h4 style='color:#F28C38; margin-top:0; text-align:center;'>{row['CATEGORIE']}</h4>
                    <div style='display:flex; justify-content:space-between; margin-bottom:5px;'>
                        <span style='color:#FFFFFF;'>Total:</span>
                        <span style='color:#FFFFFF; font-weight:bold;'>{row['Total']:,.0f} DH</span>
                    </div>
                    <div style='display:flex; justify-content:space-between; margin-bottom:5px;'>
                        <span style='color:#FFFFFF;'>Moyenne des engins par jour:</span>
                        <span style='color:#FFFFFF; font-weight:bold;'>{row['Moyenne']:,.0f} DH</span>
                    </div>
                </div>
                """, unsafe_allow_html=True)

        # Pivot table
        st.markdown("#### Consommation des catégories par type de consommation")
        hist_data = filtered_data.groupby(['CATEGORIE', 'Desc_Cat'])['Montant'].sum().reset_index()
        fig_hist = px.bar(
            hist_data,
            x='CATEGORIE',
            y='Montant',
            color='Desc_Cat',
            barmode='group',
            title='Consommation par catégorie et type de consommation',
            height=500,
            text='Desc_Cat'
        )
        fig_hist.update_traces(
            texttemplate='%{text}',
            textposition='inside',
            textfont=dict(
                size=30,
                color='#000000',
                family='Gravitas One, sans-serif'
            )
        )
        fig_hist.update_layout(
            xaxis_title="Catégorie",
            yaxis_title="Montant total (DH)",
            template='plotly_white',
            legend_title="Type de consommation",
            xaxis={'tickangle': 45},
            showlegend=False
        )
        st.plotly_chart(fig_hist, use_container_width=True, key="category_consumption")
        
        # Pivot table for CATEGORIE vs Desc_Cat
        st.markdown("#### Consommation totale par type d'engin et catégorie de consommation")
        pivot_table = pd.pivot_table(
            filtered_data,
            values='Montant',
            index='CATEGORIE',
            columns='Desc_Cat',
            aggfunc='sum',
            fill_value=0,
            margins=True,
            margins_name='Total'
        ).round(2)
        st.dataframe(
            pivot_table.style.format("{:,.2f} DH").set_properties(**{
                'background-color': '#424242',
                'border': '1px solid #ddd',
                'text-align': 'center',
                'color': '#FFFFFF'
            }).set_table_styles([
                {'selector': 'th', 'props': [('background-color', '#424242'), ('color', '#F28C38'), ('font-weight', 'bold')]}
            ]),
            use_container_width=True
        )

        # Pivot table for selected CATEGORIE
        st.markdown("#### Consommation par équipement pour le type d'engin sélectionné")
        engine_data = filtered_data[filtered_data['CATEGORIE'] == selected_engine]
        if not engine_data.empty:
            pivot_engine = pd.pivot_table(
                engine_data,
                values='Montant',
                index='Desc_CA',
                columns='Desc_Cat',
                aggfunc='sum',
                fill_value=0,
                margins=True,
                margins_name='Total'
            ).round(2)
            st.dataframe(
                pivot_engine.style.format("{:,.2f} DH").set_properties(**{
                    'background-color': '#424242',
                    'border': '1px solid #ddd',
                    'text-align': 'center',
                    'color': '#FFFFFF'
                }).set_table_styles([
                    {'selector': 'th', 'props': [('background-color', '#424242'), ('color', '#F28C38'), ('font-weight', 'bold')]}
                ]),
                use_container_width=True
            )
        else:
            st.warning(f"Aucune donnée disponible pour {selected_engine}.")

    # Onglets
    tabs = st.tabs(
        [f"📋 {cat}" for cat in sorted(filtered_data['CATEGORIE'].unique())] + 
        ["📊 Analyse comparative", "💡 Recommandations", "📋 Tableau des équipements"]
    )

    # Category tabs
    figures = {}
    for i, cat in enumerate(sorted(filtered_data['CATEGORIE'].unique())):
        with tabs[i]:
            cat_data = filtered_data[filtered_data['CATEGORIE'] == cat]
            st.markdown(f"""
            <div style='background-color:#424242; padding:20px; border-radius:10px; border-left:5px solid #1976d2; margin-bottom:20px;'>
                <h2 style='color:#F28C38; margin-top:0;'>Analyse pour la catégorie {cat}</h2>
            </div>
            """, unsafe_allow_html=True)
            
            st.markdown("#### Consommation par équipement")
            equip_sum = cat_data.groupby('Desc_CA')['Montant'].sum().reset_index().sort_values('Montant', ascending=False)
            fig2 = px.bar(
                equip_sum,
                x='Desc_CA',
                y='Montant',
                title=f'Consommation totale par équipement ({cat})',
                height=400,
                text='Montant'
            )
            fig2.update_traces(
                texttemplate='%{text:,.0f} DH',
                textposition='auto'
            )
            fig2.update_layout(
                xaxis_title="Équipement",
                yaxis_title="Montant total (DH)",
                template='plotly_white',
                xaxis={'categoryorder':'total descending'}
            )
            st.plotly_chart(fig2, use_container_width=True, key=f"equip_sum_{cat}")
            figures[f"Consommation totale par équipement ({cat})"] = fig2

    # Analyse comparative
    with tabs[-3]:
        st.markdown("""
        <div style='background-color:#424242; padding:20px; border-radius:10px; border-left:5px solid #388e3c; margin-bottom:20px;'>
            <h2 style='color:#F28C38; margin-top:0;'>Analyse comparative</h2>
        </div>
        """, unsafe_allow_html=True)
        
        st.markdown("#### Comparaison des catégories")
        fig_comp = px.bar(
            filtered_data.groupby('CATEGORIE')['Montant'].sum().reset_index(),
            x='CATEGORIE',
            y='Montant',
            title='Coût total par catégorie',
            height=400,
            text='Montant'
        )
        fig_comp.update_traces(
            texttemplate='%{text:,.0f} DH',
            textposition='auto'
        )
        fig_comp.update_layout(
            xaxis_title="Catégorie",
            yaxis_title="Montant total (DH)",
            template='plotly_white'
        )
        st.plotly_chart(fig_comp, use_container_width=True, key="category_comparison")
        figures['Coût total par catégorie'] = fig_comp

    # Recommandations
    with tabs[-2]:
        st.markdown("""
        <div style='background-color:#424242; padding:20px; border-radius:10px; border-left:5px solid #8e24aa; margin-bottom:20px;'>
            <h2 style='color:#F28C38; margin-top:0;'>Recommandations</h2>
        </div>
        """, unsafe_allow_html=True)
        
        st.markdown("#### Catégories prioritaires")
        top_categories = filtered_data.groupby('CATEGORIE')['Montant'].sum().nlargest(3).reset_index()
        cols = st.columns(3)
        colors = ['#d32f2f', '#ffa000', '#388e3c']
        for i, (col, (_, row)) in enumerate(zip(cols, top_categories.iterrows())):
            with col:
                st.markdown(f"""
                <div style='background-color:#424242; padding:15px; border-radius:10px; border-left:5px solid {colors[i]};'>
                    <h4 style='color:#F28C38; text-align:center;'>{row['CATEGORIE']}</h4>
                    <p style='color:#FFFFFF; text-align:center; font-size:24px; font-weight:bold;'>{row['Montant']:,.0f} DH</p>
                    <p style='color:#FFFFFF; text-align:center;'>{(row['Montant']/total_cost)*100:.1f}% du total</p>
                </div>
                """, unsafe_allow_html=True)

        st.markdown("""
        <div style='background-color:#424242; padding:20px; border-radius:10px; margin-top:20px;'>
            <h3 style='color:#F28C38;'>Actions recommandées</h3>
            <ul style='color:#FFFFFF;'>
                <li>Prioriser les analyses des équipements dans les catégories les plus coûteuses</li>
                <li>Mettre en place un suivi mensuel des consommations par catégorie</li>
                <li>Comparer les performances des équipements similaires pour identifier les anomalies</li>
                <li>Négocier avec les fournisseurs pour les pièces les plus fréquemment remplacées</li>
                <li>Étudier la possibilité de maintenance préventive pour réduire les coûts</li>
                <li>Former les opérateurs à une utilisation optimale des équipements</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)

    # Tableau des équipements
    with tabs[-1]:
        st.markdown("""
        <div style='background-color:#424242; padding:20px; border-radius:10px; border-left:5px solid #388e3c; margin-bottom:20px;'>
            <h2 style='color:#F28C38; margin-top:0;'>Tableau de la consommation des équipements</h2>
            <p style='color:#FFFFFF;'>Consommation détaillée par équipement pour la catégorie sélectionnée</p>
        </div>
        """, unsafe_allow_html=True)
        
        table_df = filtered_data[['Date', 'Desc_CA', 'Desc_Cat', 'Montant']].copy()
        table_df['Date'] = table_df['Date'].dt.strftime('%d/%m/%Y')
        table_df['Montant'] = table_df['Montant'].round(2)
        table_df = table_df.rename(columns={
            'Date': 'Date',
            'Desc_CA': 'Équipement',
            'Desc_Cat': 'Type de consommation',
            'Montant': 'Montant (DH)'
        })
        
        total_montant = table_df['Montant (DH)'].sum()
        
        st.dataframe(
            table_df.style.format({
                'Montant (DH)': '{:,.2f} DH',
                'Date': lambda x: x if x else ''
            }).set_properties(**{
                'background-color': '#424242',
                'border': '1px solid #ddd',
                'text-align': 'center',
                'color': '#FFFFFF'
            }).set_table_styles([
                {'selector': 'th', 'props': [('background-color', '#424242'), ('color', '#F28C38'), ('font-weight', 'bold')]}
            ]),
            height=600,
            use_container_width=True
        )
        
        st.markdown(f"""
        <div style='background-color:#424242; padding:10px; border-radius:10px; text-align:right; margin-top:10px;'>
            <p style='color:#FFFFFF; font-size:16px; font-weight:bold;'>Total : {total_montant:,.2f} DH</p>
        </div>
        """, unsafe_allow_html=True)