import streamlit as st
import requests
import json
import os
import base64
from datetime import datetime
import pandas as pd
import PyPDF2  # Bibliothèque pour manipuler les PDF
import io  # Pour manipuler les fichiers en mémoire
import time  # Pour ajouter des délais si nécessaire
import zipfile  # Pour créer des archives ZIP
import gc  # Pour le garbage collector
import sys  # Pour accéder aux références d'objets

st.set_page_config(
    page_title="Payfit - Récupération des bulletins de paie",
    page_icon="📄",
    layout="wide"
)

# Fonction pour télécharger un fichier
def get_binary_file_downloader_html(bin_file, file_label):
    with open(bin_file, 'rb') as f:
        data = f.read()
    b64 = base64.b64encode(data).decode()
    href = f'<a href="data:application/pdf;base64,{b64}" download="{os.path.basename(bin_file)}">Télécharger {file_label}</a>'
    return href

# Fonction pour créer un ZIP en mémoire avec tous les bulletins
def create_zip_in_memory(payslip_data):
    # Créer un buffer en mémoire pour le ZIP
    zip_buffer = io.BytesIO()
    
    # Créer le ZIP directement en mémoire
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for name, data in payslip_data.items():
            # Ajouter chaque bulletin au ZIP
            zipf.writestr(data["file_name"], data["content"])
    
    # Revenir au début du buffer
    zip_buffer.seek(0)
    return zip_buffer.getvalue()

# Fonction pour extraire la 2ème page d'un PDF et la renvoyer comme contenu binaire
def extract_second_page(pdf_content):
    try:
        # Utiliser un objet BytesIO pour manipuler le PDF en mémoire sans fichier temporaire
        input_pdf = io.BytesIO(pdf_content)
        
        # Lire le PDF depuis la mémoire
        pdf_reader = PyPDF2.PdfReader(input_pdf)
        
        # Vérifier si le PDF a au moins 2 pages
        if len(pdf_reader.pages) < 2:
            st.warning("Le fichier n'a qu'une seule page.")
            return None
        
        # Créer un nouveau PDF avec seulement la 2ème page
        pdf_writer = PyPDF2.PdfWriter()
        pdf_writer.add_page(pdf_reader.pages[1])  # Ajouter la 2ème page (index 1)
        
        # Écrire le nouveau PDF dans un BytesIO
        output_pdf = io.BytesIO()
        pdf_writer.write(output_pdf)
        output_pdf.seek(0)  # Revenir au début du buffer
        
        return output_pdf.getvalue()  # Retourner le contenu binaire
    except Exception as e:
        st.error(f"Erreur lors de l'extraction de la 2ème page: {str(e)}")
        return None

# Fonction pour créer un bouton de téléchargement CSV (à l'extérieur du form)
def create_csv_download(df, filename):
    csv = df.to_csv(index=False).encode('utf-8')
    return csv, filename

# Fonction pour vérifier la présence de la clé API en mémoire
def check_api_key_in_memory():
    # Forcer le garbage collector à nettoyer les objets non référencés
    gc.collect()
    
    # Vérifier si 'api_key' est dans les variables locales ou globales
    api_key_in_locals = 'api_key' in locals()
    api_key_in_globals = 'api_key' in globals()
    
    # Rechercher dans toutes les références d'objets (plus approfondi)
    api_key_refs = []
    for obj in gc.get_objects():
        try:
            if isinstance(obj, str) and len(obj) > 20 and 'api' in obj.lower():
                api_key_refs.append(obj)
        except:
            pass
    
    results = {
        "API Key dans variables locales": api_key_in_locals,
        "API Key dans variables globales": api_key_in_globals,
        "Nombre de références potentielles à API Key": len(api_key_refs)
    }
    
    return results

# Fonction principale pour récupérer les bulletins de paie
def get_payslips(api_key, target_year, target_month):
    # Validation des entrées
    if not api_key or not target_year or not target_month:
        st.error("Tous les champs sont obligatoires!")
        return
    
    BASE_URL = "https://partner-api.payfit.com"
    
    headers_json = {
        'Content-Type': 'application/json',
        'Authorization': f'Bearer {api_key}'
    }
    headers_auth = {
        'Authorization': f'Bearer {api_key}'
    }
    
    # Afficher une barre de progression
    progress_bar = st.progress(0)
    status_placeholder = st.empty()
    
    # 1. Vérification de la clé API
    status_placeholder.info("🔎 Vérification de la clé API...")
    resp = requests.post("https://oauth.payfit.com/introspect", headers=headers_json, data=json.dumps({"token": api_key}))
    
    try:
        data = resp.json()
        if not data.get("active"):
            st.error("❌ Clé API invalide ou expirée.")
            return
        
        company_id = data["company_id"]
        status_placeholder.success(f"✅ Clé valide. Entreprise ID : {company_id}")
        progress_bar.progress(10)
        
        # 2. Infos détaillées de l'entreprise
        company_info = requests.get(f"{BASE_URL}/companies/{company_id}", headers=headers_auth).json()
        
        with st.expander("🏢 Informations détaillées de l'entreprise", expanded=True):
            col1, col2 = st.columns(2)
            with col1:
                st.write(f"**Nom de l'entreprise:** {company_info['name']}")
                st.write(f"**ID de l'entreprise:** {company_id}")
                st.write(f"**Contrats actifs:** {company_info['nbActiveContracts']}")
            with col2:
                if 'countryCode' in company_info:
                    st.write(f"**Pays:** {company_info['countryCode']}")
                if 'city' in company_info:
                    st.write(f"**Ville:** {company_info['city']}")
                if 'postalCode' in company_info:
                    st.write(f"**Code postal:** {company_info['postalCode']}")
            
            # Autres informations disponibles dans company_info
            st.subheader("Informations supplémentaires")
            remaining_info = {k: v for k, v in company_info.items() 
                             if k not in ['name', 'countryCode', 'city', 'postalCode', 'nbActiveContracts']}
            st.json(remaining_info)
        
        progress_bar.progress(20)
        
        # 3. Récupération des collaborateurs avec pagination
        status_placeholder.info("📥 Récupération des collaborateurs...")
        all_collabs = []
        next_page_token = None
        
        page_count = 0
        while True:
            page_count += 1
            params = {"nextPageToken": next_page_token} if next_page_token else {}
            response = requests.get(f"{BASE_URL}/companies/{company_id}/collaborators", headers=headers_auth, params=params)
            
            if response.status_code != 200:
                st.error(f"❌ Erreur lors de la récupération des collaborateurs: {response.status_code}")
                break
            
            collabs_response = response.json()
            page_collabs = collabs_response.get("collaborators", [])
            all_collabs.extend(page_collabs)
            
            # Mise à jour du statut avec la pagination
            status_placeholder.info(f"📥 Récupération des collaborateurs... Page {page_count} ({len(all_collabs)} collaborateurs jusqu'à présent)")
            
            next_page_token = collabs_response.get("meta", {}).get("nextPageToken")
            if not next_page_token:
                break  # Sortie de boucle si plus de pages
        
        collabs = all_collabs
        status_placeholder.success(f"✅ {len(collabs)} collaborateurs récupérés.")
        progress_bar.progress(30)
        
        # Affichage de tous les collaborateurs avec plus d'informations
        with st.expander("👥 Liste complète des collaborateurs", expanded=True):
            if collabs:
                # Préparation des données pour le tableau
                collabs_data = []
                for collab in collabs:
                    collab_data = {
                        "ID": collab.get("id", ""),
                        "Prénom": collab.get("firstName", ""),
                        "Nom": collab.get("lastName", ""),
                        "Email": collab.get("email", ""),
                        "Statut": "Actif" if collab.get("status") == "active" else "Inactif"
                    }
                    # Ajouter d'autres champs s'ils existent
                    if "startDate" in collab:
                        collab_data["Date de début"] = collab["startDate"]
                    if "endDate" in collab:
                        collab_data["Date de fin"] = collab["endDate"]
                    
                    collabs_data.append(collab_data)
                
                # Création du DataFrame pour affichage
                df = pd.DataFrame(collabs_data)
                st.dataframe(df, use_container_width=True)
                
                # Préparation des données pour le bouton de téléchargement (mais ne pas créer le bouton ici)
                csv_data, csv_filename = create_csv_download(
                    df, 
                    f"collaborateurs_{company_info['name']}_{datetime.now().strftime('%Y%m%d')}.csv"
                )
                
                # Stockage des données dans session_state pour utilisation externe au formulaire
                st.session_state.csv_data = csv_data
                st.session_state.csv_filename = csv_filename
                st.session_state.show_download_button = True
            else:
                st.warning("Aucun collaborateur trouvé.")
                st.session_state.show_download_button = False
        
        # Création du dossier de sortie si nécessaire (pour le ZIP)
        os.makedirs("bulletins_paie", exist_ok=True)
        
        # 4. Filtrage + extraction
        status_placeholder.info("🔍 Recherche des bulletins de paie...")
        
        collabs_with_payslip = []
        collabs_without_payslip = []
        payslip_data = {}  # Dictionnaire pour stocker les données des PDF au lieu de les télécharger automatiquement
        
        with st.expander("Détails du traitement", expanded=False):
            details_placeholder = st.empty()
            detail_text = ""
        
        total_collabs = len(collabs)
        for i, collab in enumerate(collabs):
            # Mise à jour de la progression
            progress_value = 30 + (i / total_collabs * 60)
            progress_bar.progress(int(progress_value))
            
            collaborator_id = collab["id"]
            full_name = f"{collab.get('firstName', '')} {collab.get('lastName', '')}".strip()
            
            detail_text += f"Traitement de {full_name}...\n"
            details_placeholder.text_area("Logs", detail_text, height=400)
            
            payslips_url = f"{BASE_URL}/companies/{company_id}/collaborators/{collaborator_id}/payslips/"
            payslip_resp = requests.get(payslips_url, headers=headers_auth).json()
            
            if "payslips" not in payslip_resp or not payslip_resp["payslips"]:
                collabs_without_payslip.append({
                    "name": full_name,
                    "id": collaborator_id,
                    "reason": "Aucun bulletin disponible"
                })
                detail_text += f"  → Aucun bulletin disponible\n"
                details_placeholder.text_area("Logs", detail_text, height=400)
                continue
            
            # Recherche d'un bulletin correspondant à la période demandée
            target_payslip = next(
                (p for p in payslip_resp["payslips"] if p["year"] == target_year and p["month"] == target_month),
                None
            )
            
            if target_payslip:
                collabs_with_payslip.append({
                    "name": full_name,
                    "id": collaborator_id,
                    "payslip_info": target_payslip
                })
                detail_text += f"  → ✅ Bulletin trouvé pour {target_month}/{target_year}\n"
                details_placeholder.text_area("Logs", detail_text, height=400)
                
                contract_id = target_payslip["contractId"]
                payslip_id = target_payslip["payslipId"]
                
                pdf_url = f"{BASE_URL}/companies/{company_id}/collaborators/{collaborator_id}/contracts/{contract_id}/payslips/{payslip_id}"
                pdf_response = requests.get(pdf_url, headers={**headers_auth, 'accept': 'application/pdf'})
                
                if pdf_response.status_code == 200:
                    file_safe_name = f"{collab.get('firstName', 'collaborateur')}_{collab.get('lastName', '')}".replace(" ", "_")
                    
                    # Nom du fichier sans "page2"
                    file_name = f"{file_safe_name}_{target_year}_{target_month}.pdf"
                    
                    # Extraire la 2ème page et la stocker en mémoire
                    extracted_content = extract_second_page(pdf_response.content)
                    if extracted_content:
                        # Stockage des données de page 2 dans un dictionnaire
                        payslip_data[full_name] = {
                            "file_name": file_name,
                            "content": extracted_content
                        }
                        detail_text += f"  → ✅ 2ème page extraite et prête pour téléchargement\n"
                    else:
                        detail_text += f"  → ⚠️ Impossible d'extraire la 2ème page\n"
                else:
                    detail_text += f"  → ❌ Erreur lors du téléchargement du bulletin de paie (code {pdf_response.status_code})\n"
                
                details_placeholder.text_area("Logs", detail_text, height=400)
            else:
                collabs_without_payslip.append({
                    "name": full_name,
                    "id": collaborator_id,
                    "reason": f"Pas de bulletin pour {target_month}/{target_year}",
                    "available_periods": [f"{p['month']}/{p['year']}" for p in payslip_resp["payslips"]]
                })
                detail_text += f"  → Pas de bulletin pour {target_month}/{target_year}\n"
                details_placeholder.text_area("Logs", detail_text, height=400)
        
        progress_bar.progress(100)
        status_placeholder.success("✅ Traitement terminé!")
        
        # Stockage des données des bulletins dans la session_state
        st.session_state.payslip_data = payslip_data
        st.session_state.collabs_with_payslip = collabs_with_payslip
        st.session_state.collabs_without_payslip = collabs_without_payslip
        st.session_state.target_year = target_year
        st.session_state.target_month = target_month
        st.session_state.show_results = True
        
        # Préparer le ZIP en mémoire pour le téléchargement en une étape
        if payslip_data:
            # Créer le contenu ZIP en mémoire
            zip_content = create_zip_in_memory(payslip_data)
            # Stocker dans session_state
            st.session_state.zip_content = zip_content
            st.session_state.zip_filename = f"bulletins_paie_{target_year}_{target_month}.zip"
        
    except Exception as e:
        st.error(f"Une erreur est survenue: {str(e)}")
        import traceback
        st.error(traceback.format_exc())
        st.session_state.traitement_termine = False
        st.session_state.show_results = False

# Initialisation des variables de session si elles n'existent pas
if 'show_download_button' not in st.session_state:
    st.session_state.show_download_button = False
if 'traitement_termine' not in st.session_state:
    st.session_state.traitement_termine = False
if 'payslip_data' not in st.session_state:
    st.session_state.payslip_data = {}
if 'show_results' not in st.session_state:
    st.session_state.show_results = False
if 'collabs_with_payslip' not in st.session_state:
    st.session_state.collabs_with_payslip = []
if 'collabs_without_payslip' not in st.session_state:
    st.session_state.collabs_without_payslip = []
if 'zip_content' not in st.session_state:
    st.session_state.zip_content = None
if 'zip_filename' not in st.session_state:
    st.session_state.zip_filename = ""
if 'show_debug_mode' not in st.session_state:
    st.session_state.show_debug_mode = False
if 'memory_check_results' not in st.session_state:
    st.session_state.memory_check_results = None

# Interface utilisateur
st.title("📄 Récupération des bulletins de paie Payfit")
st.write("Cet outil vous permet de télécharger la deuxième page des bulletins de paie de vos collaborateurs pour un mois spécifique et d'obtenir des informations détaillées sur votre entreprise et vos collaborateurs.")

p="""
# Activer/désactiver le mode debug
with st.expander("🧪 Mode Test - Sécurité API Key", expanded=False):
    st.write("Activez ce mode pour vérifier si la clé API est correctement supprimée de la mémoire après utilisation.")
    debug_mode = st.checkbox("Activer le mode test de sécurité", value=st.session_state.show_debug_mode)
    st.session_state.show_debug_mode = debug_mode
    
    if st.session_state.memory_check_results:
        st.subheader("Résultats du test de sécurité")
        for key, value in st.session_state.memory_check_results.items():
            st.write(f"**{key}:** {value}")
            
        if not any(v == True for v in st.session_state.memory_check_results.values() if isinstance(v, bool)):
            st.success("✅ La clé API semble avoir été correctement supprimée de la mémoire!")
        else:
            st.error("❌ La clé API pourrait encore être présente en mémoire.")
"""
with st.form(key="payslip_form"):
    api_key = st.text_input("🔐 Clé API Payfit", type="password", help="Vous pouvez obtenir une clé API depuis votre compte Payfit")
    
    col1, col2 = st.columns(2)
    with col1:
        current_year = datetime.now().year
        years = list(range(current_year-3, current_year+1))
        target_year = st.selectbox("📅 Année", options=years, index=len(years)-1)
    
    with col2:
        months = [(str(i).zfill(2), datetime(2000, i, 1).strftime("%B")) for i in range(1, 13)]
        current_month = datetime.now().month - 1  # Mois précédent par défaut
        target_month = st.selectbox(
            "📅 Mois", 
            options=[m[0] for m in months],
            format_func=lambda x: months[int(x)-1][1],
            index=current_month - 1 if current_month > 0 else 0
        )
    
    submit_button = st.form_submit_button(label="📥 Récupérer les bulletins")
    
    if submit_button:
        get_payslips(api_key, str(target_year), target_month)
        # Suppression sécurisée de la clé API de la mémoire
        del api_key
        
        # Vérification de la mémoire si le mode debug est activé
        if st.session_state.show_debug_mode:
            st.session_state.memory_check_results = check_api_key_in_memory()
            st.experimental_rerun()  # Recharger pour afficher les résultats

# Bouton de téléchargement placé en dehors du formulaire
if st.session_state.show_download_button and 'csv_data' in st.session_state and 'csv_filename' in st.session_state:
    st.download_button(
        label="📥 Télécharger la liste des collaborateurs (CSV)",
        data=st.session_state.csv_data,
        file_name=st.session_state.csv_filename,
        mime="text/csv",
    )

# Affichage des résultats en dehors du formulaire
if st.session_state.show_results:
    # 5. Résumé des bulletins récupérés
    st.subheader("📄 Bulletins récupérés")
    
    target_year = st.session_state.target_year
    target_month = st.session_state.target_month
    payslip_data = st.session_state.payslip_data
    collabs_with_payslip = st.session_state.collabs_with_payslip
    collabs_without_payslip = st.session_state.collabs_without_payslip
    
    # Statistiques des bulletins
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Total des collaborateurs", len(collabs_with_payslip) + len(collabs_without_payslip))
    with col2:
        st.metric("Bulletins trouvés", len(collabs_with_payslip))
    with col3:
        st.metric("Bulletins manquants", len(collabs_without_payslip))
    
    # Affichage des collaborateurs avec bulletins
    if collabs_with_payslip:
        # Bouton pour télécharger tous les bulletins en une seule étape
        if 'zip_content' in st.session_state and st.session_state.zip_content:
            col1, col2 = st.columns([3, 1])
            with col1:
                st.info(f"**{len(collabs_with_payslip)} bulletins trouvés pour la période {target_month}/{target_year}**")
            with col2:
                st.download_button(
                    label="📥 Télécharger tous les bulletins",
                    data=st.session_state.zip_content,
                    file_name=st.session_state.zip_filename,
                    mime="application/zip",
                )
        else:
            st.write(f"**{len(collabs_with_payslip)} bulletins trouvés pour la période {target_month}/{target_year}:**")
        
        st.write("---")
        st.subheader("Bulletins individuels")
        
        # Tableau avec les boutons de téléchargement individuels
        for name in payslip_data.keys():
            col1, col2 = st.columns([3, 1])
            with col1:
                st.write(f"**{name}**")
            with col2:
                # Bouton de téléchargement direct depuis la mémoire via Streamlit
                key = f"dl_btn_{name.replace(' ', '_')}"
                st.download_button(
                    label="Télécharger bulletin de paie",
                    data=payslip_data[name]["content"],
                    file_name=payslip_data[name]["file_name"],
                    mime="application/pdf",
                    key=key
                )
        
        # Option avancée pour créer le ZIP sur disque si nécessaire
        with st.expander("⚙️ Options avancées", expanded=False):
            st.subheader("Création d'une archive ZIP sur disque")
            st.write("Utilisez cette option si vous avez besoin d'une archive ZIP enregistrée localement.")
            
            if st.button("Préparer l'archive ZIP sur disque"):
                with st.spinner("Préparation de l'archive ZIP..."):
                    # Création d'un ZIP avec tous les bulletins
                    zip_filename = f"bulletins_paie_{target_year}_{target_month}.zip"
                    
                    # Enregistrer temporairement les fichiers pour les ajouter au ZIP
                    temp_files = []
                    for name, data in payslip_data.items():
                        file_path = f"bulletins_paie/{data['file_name']}"
                        with open(file_path, "wb") as f:
                            f.write(data["content"])
                        temp_files.append(file_path)
                    
                    # Créer le ZIP
                    with zipfile.ZipFile(zip_filename, 'w') as zipf:
                        for file_path in temp_files:
                            zipf.write(file_path, arcname=os.path.basename(file_path))
                    
                    # Afficher le lien de téléchargement
                    st.markdown(get_binary_file_downloader_html(zip_filename, "tous les bulletins (ZIP)"), unsafe_allow_html=True)
    else:
        st.warning("Aucun bulletin trouvé pour ce mois.")
    
    # 6. Informations sur les collaborateurs sans bulletin
    with st.expander("ℹ️ Collaborateurs sans bulletin pour cette période", expanded=True):
        if collabs_without_payslip:
            # Préparation des données pour le tableau
            missing_data = []
            for collab in collabs_without_payslip:
                missing_item = {
                    "Nom": collab["name"],
                    "Raison": collab["reason"]
                }
                if "available_periods" in collab:
                    missing_item["Périodes disponibles"] = ", ".join(collab["available_periods"])
                
                missing_data.append(missing_item)
            
            # Création du DataFrame pour affichage
            df_missing = pd.DataFrame(missing_data)
            st.dataframe(df_missing, use_container_width=True)
        else:
            st.success("Tous les collaborateurs ont un bulletin pour cette période!")
    
    # 7. Historique des bulletins disponibles est intégré directement dans la fonction get_payslips pour simplifier

# Aide et instructions
with st.expander("ℹ️ Comment utiliser cet outil"):
    st.write("""
    1. **Obtenir une clé API Payfit** : Connectez-vous à votre compte administrateur Payfit et générez une clé API.
    2. **Sélectionner la période** : Choisissez l'année et le mois pour lesquels vous souhaitez récupérer les bulletins.
    3. **Lancer la récupération** : Cliquez sur le bouton pour commencer le processus.
    4. **Explorer les résultats** : Une fois le traitement terminé, vous pourrez :
       - Voir les informations détaillées de votre entreprise
       - Consulter la liste complète de vos collaborateurs
       - Télécharger tous les bulletins en une seule fois avec un clic
       - Télécharger les bulletins individuellement
       - Voir quels collaborateurs n'ont pas de bulletin pour la période sélectionnée
    """)

st.divider()
st.write("**Note :** Seules les 2èmes pages des bulletins de paie sont extraites et disponibles pour téléchargement. Les bulletins complets ne sont pas stockés localement.")