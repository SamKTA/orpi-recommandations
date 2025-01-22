import streamlit as st
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import os

# Configuration de la page
st.set_page_config(page_title="Recommandations internes", layout="wide")

# Titre et sous-titre
st.title("Recommandations internes")
st.subheader("Agence Orpi Panazol, Arcades, La Souterraine et Saint-Yrieix")

# Définition des projets et leurs couleurs associées
PROJETS = {
    "Vente": "#FF6B6B",
    "Achat": "#4ECDC4",
    "Location": "#45B7D1",
    "Gestion": "#96CEB4",
    "Syndic": "#FFEEAD",
    "Orpi PRO": "#D4A5A5",
    "Location + Gestion": "#9AC1D9",
    "Recrutement": "#FFD93D",
    "CGI": "#6C5B7B"
}

# Création du formulaire
with st.form("recommandation_form"):
    # Date automatique
    date_recommandation = datetime.now().strftime("%d/%m/%Y")
    
    # Informations de l'expéditeur et receveur
    prescripteur = st.text_input("Votre nom complet *", key="prescripteur")
    email_receveur = st.text_input("E-mail du receveur de la recommandation *", key="email_receveur")
    
    # Informations du client
    col1, col2 = st.columns(2)
    with col1:
        nom_client = st.text_input("Nom client *", key="nom_client")
        telephone_client = st.text_input("Téléphone client *", key="telephone_client")
    with col2:
        email_client = st.text_input("Mail client *", key="email_client")
        projet = st.selectbox("Projet concerné *", 
                            options=list(PROJETS.keys()),
                            key="projet")
    
    # Détails et adresse
    details_projet = st.text_area("Détails du projet", key="details_projet")
    adresse_projet = st.text_input("Adresse du projet *", key="adresse_projet")
    
    # Bouton de validation
    submitted = st.form_submit_button("Je valide ma recommandation")

def sauvegarder_dans_sheets(donnees):
    SHEET_ID = "1dkjKAvwlALjo8RHkIm-6PQlAhEkorrA5T9d5CnrMBIA"
    scope = ['https://spreadsheets.google.com/feeds',
             'https://www.googleapis.com/auth/drive']
    
    # Utiliser les credentials stockés dans Streamlit Secrets
    creds_dict = st.secrets["google_credentials"]
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    client = gspread.authorize(creds)
    
    # Ouvrir la feuille de calcul
    sheet = client.open_by_key(SHEET_ID).sheet1
    
    # Ajouter les données dans l'ordre exact des colonnes
    # Création de la date au format jour/mois/année
    date_ajout = datetime.now().strftime("%d/%m/%Y")
    
    # Ajout des données dans l'ordre exact des colonnes du Google Sheet avec des cellules vides pour les colonnes à ne pas remplir
    sheet.append_row([
        date_ajout,                  # Date
        donnees["prescripteur"],     # Nom complet du prescripteur
        donnees["email_receveur"],   # Mail du receveur
        donnees["nom_client"],       # Nom client
        donnees["telephone_client"], # Tél client
        donnees["email_client"],     # Mail client
        donnees["projet"],          # Projet concerné
        donnees["details_projet"],   # Détails du projet
        donnees["adresse_projet"],   # Adresse du projet
        "",                         # Prise en charge
        "",                         # Où en sommes-nous ?
        "",                         # Date transfert
        "",                         # CA généré
        "",                         # Validat° Prime OS
        "",                         # DATE RÈGLEMENT
        "",                         # Ajout sens
        "",                         # Nature du service prescripteur
        "",                         # CA instantané
        "",                         # CA/an
        "",                         # Mois de la reco
        "",                         # prnom
        "",                         # mail
        "",                         # Remarque(s)
    ])
    
    return "https://docs.google.com/spreadsheets/d/1dkjKAvwlALjo8RHkIm-6PQlAhEkorrA5T9d5CnrMBIA/edit?usp=sharing"

def envoyer_email(prescripteur, email_receveur, projet, lien):
    try:
        # Configuration de l'email
        msg = MIMEMultipart()
        msg['From'] = st.secrets["email"]["username"]
        msg['To'] = email_receveur
        msg['Subject'] = f"Nouvelle recommandation - {projet}"
        
        body = f"""
        Hello !

        Tu as reçu une nouvelle recommandation de la part de {prescripteur}.

        Cette recommandation concerne un projet de {projet}

        Pour avoir accès à ta recommandation, voici le lien : {lien}
        """
        
        msg.attach(MIMEText(body, 'plain'))
        
        # Configuration du serveur SMTP
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(st.secrets["email"]["username"], st.secrets["email"]["password"])
        server.send_message(msg)
        server.quit()
        return True
    except Exception as e:
        st.error(f"Erreur lors de l'envoi de l'email: {str(e)}")
        return False

# Traitement du formulaire
if submitted:
    # Vérification des champs obligatoires
    required_fields = {
        "prescripteur": prescripteur,
        "email_receveur": email_receveur,
        "nom_client": nom_client,
        "telephone_client": telephone_client,
        "email_client": email_client,
        "projet": projet,
        "adresse_projet": adresse_projet
    }
    
    missing_fields = [field for field, value in required_fields.items() if not value]
    
    if missing_fields:
        st.error("Veuillez remplir tous les champs obligatoires.")
    else:
        # Préparer les données
        donnees = {
            "date": date_recommandation,
            "prescripteur": prescripteur,
            "email_receveur": email_receveur,
            "nom_client": nom_client,
            "telephone_client": telephone_client,
            "email_client": email_client,
            "projet": projet,
            "details_projet": details_projet,
            "adresse_projet": adresse_projet
        }
        
        # Sauvegarder et envoyer
        try:
            lien = sauvegarder_dans_sheets(donnees)
            if envoyer_email(prescripteur, email_receveur, projet, lien):
                st.success("Recommandation envoyée avec succès !")
            else:
                st.error("Erreur lors de l'envoi de l'email.")
        except Exception as e:
            st.error(f"Erreur lors de la sauvegarde : {str(e)}")

# Style CSS pour les couleurs des projets
css = """
<style>
"""
for projet, couleur in PROJETS.items():
    css += f"""
    div[data-value="{projet}"] {{
        background-color: {couleur} !important;
        color: white !important;
    }}
    """
css += "</style>"

st.markdown(css, unsafe_allow_html=True)
