import streamlit as st
from streamlit_option_menu import option_menu
from Login import login_page
from Coaching import coaching_page
from Recolt import recolt_page
from Equipes_Recol import *
from datetime import datetime

# 🎨 Initialisation de l'état de session
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
if "user_info" not in st.session_state:
    st.session_state.user_info = None

# 🎨 Interface de connexion
if not st.session_state.logged_in:
    login_page()
else:
    # Structure de la page en trois colonnes
    col1, col2, col3 = st.columns([1, 2, 1])

    with col1:
        # Image en haut à gauche (visible sur toutes les pages)
        st.image("Images/Logo_Axa_Noir.png", width=200)  # Ajustez la largeur selon vos besoins

    with col2:
        # Nom de la page sélectionnée au milieu
        if "selected" not in st.session_state:
            st.session_state.selected = "Coaching"  # Par défaut
        st.title(st.session_state.selected)  # Affiche le nom de la page sélectionnée

    with col3:
        # Message de bienvenue et nom de la personne connectée à droite
        st.markdown(
            f"<h3 style='text-align: right;'>Bienvenue, <strong>{st.session_state.user_info['NomP']}</strong>!</h3>",
            unsafe_allow_html=True,
        )
        # Bouton de déconnexion en haut à droite
        if st.button("Se déconnecter", key="logout_button"):
            st.session_state.logged_in = False
            st.session_state.user_info = None
            st.rerun()

    # Menu horizontal en bas
    menu_items = ["Coaching", "Équipe", "Recolt"]
    selected = option_menu(
        menu_title=None,  # Pas de titre pour le menu
        options=menu_items,
        icons=["calendar", "list-task", "calendar-event"],
        menu_icon="menu",
        default_index=0,  # Sélectionner le premier onglet par défaut
        orientation="horizontal",  # Menu horizontal
        styles={
            "container": {"padding": "0!important", "background-color": "#f0f2f6"},
            "nav-link": {"font-size": "16px", "text-align": "center", "margin": "0px", "padding": "10px"},
            "nav-link-selected": {"background-color": "#4CAF50", "color": "white"},
        }
    )

    # Mettre à jour la page sélectionnée dans l'état de session
    st.session_state.selected = selected

    # Affichage du contenu de la page sélectionnée
    st.write("---")  # Ligne de séparation
    if selected == "Coaching":
        coaching_page()
    elif selected == "Équipe":
        evaluation_result = st.radio(
                        "Résultat de l'évaluation",
                        options=["Oui", "Non", "Partielle"],
                        horizontal=True
                    )
        st.write(f"Résultat sélectionné : **{evaluation_result}**")
        commentaire = st.text_area(
                    "Commentaire :", 
                    placeholder="Entrez votre commentaire ici...",  
                    height=150, 
                )
    elif selected == "Recolt":
        recolt_page()