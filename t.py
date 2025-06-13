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

# ==================== FONCTIONS COMMUNES ====================

def get_binary_file_downloader_html(bin_file, file_label):
    with open(bin_file, 'rb') as f:
        data = f.read()
    b64 = base64.b64encode(data).decode()
    href = f'<a href="data:application/pdf;base64,{b64}" download="{os.path.basename(bin_file)}">Télécharger {file_label}</a>'
    return href

def create_zip_in_memory(payslip_data):
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for name, data in payslip_data.items():
            zipf.writestr(data["file_name"], data["content"])
    zip_buffer.seek(0)
    return zip_buffer.getvalue()

def extract_second_page(pdf_content):
    try:
        input_pdf = io.BytesIO(pdf_content)
        pdf_reader = PyPDF2.PdfReader(input_pdf)
        
        if len(pdf_reader.pages) < 2:
            st.warning("Le fichier n'a qu'une seule page.")
            return None
        
        pdf_writer = PyPDF2.PdfWriter()
        pdf_writer.add_page(pdf_reader.pages[1])
        
        output_pdf = io.BytesIO()
        pdf_writer.write(output_pdf)
        output_pdf.seek(0)
        
        return output_pdf.getvalue()
    except Exception as e:
        st.error(f"Erreur lors de l'extraction de la 2ème page: {str(e)}")
        return None

def create_csv_download(df, filename):
    csv = df.to_csv(index=False).encode('utf-8')
    return csv, filename

def check_api_key_in_memory():
    gc.collect()
    api_key_in_locals = 'api_key' in locals()
    api_key_in_globals = 'api_key' in globals()
    
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

# ==================== FONCTIONS BULLETINS ANNUELS ====================

def get_company_and_collaborators(api_key):
    """Fonction commune pour récupérer les infos de l'entreprise et les collaborateurs"""
    BASE_URL = "https://partner-api.payfit.com"
    
    headers_json = {
        'Content-Type': 'application/json',
        'Authorization': f'Bearer {api_key}'
    }
    headers_auth = {
        'Authorization': f'Bearer {api_key}'
    }
    
    # Vérification de la clé API
    resp = requests.post("https://oauth.payfit.com/introspect", headers=headers_json, data=json.dumps({"token": api_key}))
    
    try:
        data = resp.json()
        if not data.get("active"):
            return None, None, None, None
        
        company_id = data["company_id"]
        
        # Infos de l'entreprise
        company_info = requests.get(f"{BASE_URL}/companies/{company_id}", headers=headers_auth).json()
        
        # Récupération des collaborateurs avec pagination
        all_collabs = []
        next_page_token = None
        
        while True:
            params = {"nextPageToken": next_page_token} if next_page_token else {}
            response = requests.get(f"{BASE_URL}/companies/{company_id}/collaborators", headers=headers_auth, params=params)
            
            if response.status_code != 200:
                break
            
            collabs_response = response.json()
            page_collabs = collabs_response.get("collaborators", [])
            all_collabs.extend(page_collabs)
            
            next_page_token = collabs_response.get("meta", {}).get("nextPageToken")
            if not next_page_token:
                break
        
        return company_id, company_info, all_collabs, headers_auth
        
    except Exception as e:
        st.error(f"Erreur lors de la récupération des données: {str(e)}")
        return None, None, None, None

def get_yearly_payslips(api_key, target_year):
    if not api_key or not target_year:
        st.error("Tous les champs sont obligatoires!")
        return
    
    progress_bar = st.progress(0)
    status_placeholder = st.empty()
    
    # Récupération des données communes
    result = get_company_and_collaborators(api_key)
    if result[0] is None:
        st.error("❌ Clé API invalide ou expirée.")
        return
    
    company_id, company_info, collabs, headers_auth = result
    BASE_URL = "https://partner-api.payfit.com"
    
    status_placeholder.success(f"✅ Clé valide. Entreprise : {company_info['name']}")
    progress_bar.progress(10)
    
    # Structure pour organiser les bulletins par collaborateur
    collaborator_payslips = {}  # {collaborator_name: {month: pdf_content}}
    yearly_stats = {
        'total_collaborators': len(collabs),
        'collaborators_with_payslips': 0,
        'total_payslips_found': 0,
        'months_processed': set()
    }
    
    with st.expander("📊 Détails du traitement", expanded=False):
        details_placeholder = st.empty()
        detail_text = "Début du traitement des bulletins annuels...\n\n"
    
    total_collabs = len(collabs)
    for i, collab in enumerate(collabs):
        progress_value = 10 + (i / total_collabs * 80)
        progress_bar.progress(int(progress_value))
        
        collaborator_id = collab["id"]
        full_name = f"{collab.get('firstName', '')} {collab.get('lastName', '')}".strip()
        
        status_placeholder.info(f"Traitement de {full_name}... ({i+1}/{total_collabs})")
        detail_text += f"📝 Traitement de {full_name}...\n"
        details_placeholder.text_area("Logs de traitement", detail_text, height=300)
        
        # Récupération de tous les bulletins du collaborateur
        payslips_url = f"{BASE_URL}/companies/{company_id}/collaborators/{collaborator_id}/payslips/"
        payslip_resp = requests.get(payslips_url, headers=headers_auth).json()
        
        if "payslips" not in payslip_resp or not payslip_resp["payslips"]:
            detail_text += f"  ❌ Aucun bulletin disponible\n"
            details_placeholder.text_area("Logs de traitement", detail_text, height=300)
            continue
        
        # Filtrer les bulletins pour l'année cible (conversion en int pour éviter les erreurs de type)
        yearly_payslips = [p for p in payslip_resp["payslips"] if int(p["year"]) == int(target_year)]
        
        if not yearly_payslips:
            detail_text += f"  ❌ Aucun bulletin pour l'année {target_year}\n"
            details_placeholder.text_area("Logs de traitement", detail_text, height=300)
            continue
        
        # Initialiser le dictionnaire pour ce collaborateur
        collaborator_payslips[full_name] = {}
        collaborator_has_payslips = False
        
        detail_text += f"  ✅ {len(yearly_payslips)} bulletin(s) trouvé(s) pour {target_year}\n"
        
        # Télécharger chaque bulletin du collaborateur pour l'année
        for payslip in yearly_payslips:
            month = str(payslip["month"]).zfill(2)  # Conversion en string avec format 2 chiffres
            contract_id = payslip["contractId"]
            payslip_id = payslip["payslipId"]
            
            yearly_stats['months_processed'].add(month)
            
            pdf_url = f"{BASE_URL}/companies/{company_id}/collaborators/{collaborator_id}/contracts/{contract_id}/payslips/{payslip_id}"
            pdf_response = requests.get(pdf_url, headers={**headers_auth, 'accept': 'application/pdf'})
            
            if pdf_response.status_code == 200:
                # Extraire la 2ème page
                extracted_content = extract_second_page(pdf_response.content)
                if extracted_content:
                    collaborator_payslips[full_name][month] = {
                        'content': extracted_content,
                        'file_name': f"{full_name.replace(' ', '_')}_{target_year}_{month.zfill(2)}.pdf"
                    }
                    collaborator_has_payslips = True
                    yearly_stats['total_payslips_found'] += 1
                    detail_text += f"    ✅ Mois {month} - bulletin récupéré\n"
                else:
                    detail_text += f"    ⚠️ Mois {month} - impossible d'extraire la 2ème page\n"
            else:
                detail_text += f"    ❌ Mois {month} - erreur téléchargement (code {pdf_response.status_code})\n"
        
        if collaborator_has_payslips:
            yearly_stats['collaborators_with_payslips'] += 1
        
        detail_text += "\n"
        details_placeholder.text_area("Logs de traitement", detail_text, height=300)
    
    progress_bar.progress(100)
    status_placeholder.success("✅ Traitement terminé!")
    
    # Stockage des résultats
    st.session_state.yearly_payslip_data = collaborator_payslips
    st.session_state.yearly_stats = yearly_stats
    st.session_state.yearly_target_year = target_year
    st.session_state.yearly_company_info = company_info
    st.session_state.yearly_show_results = True
    
    # Création du ZIP global
    if collaborator_payslips:
        create_yearly_zip(collaborator_payslips, target_year)

def create_yearly_zip(collaborator_payslips, target_year):
    """Créer un ZIP organisé avec tous les bulletins de l'année"""
    zip_buffer = io.BytesIO()
    
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for collaborator_name, months_data in collaborator_payslips.items():
            # Créer un dossier par collaborateur
            safe_name = collaborator_name.replace(' ', '_').replace('/', '_')
            
            for month, payslip_data in months_data.items():
                # Chemin dans le ZIP : Collaborateur/fichier.pdf
                file_path_in_zip = f"{safe_name}/{payslip_data['file_name']}"
                zipf.writestr(file_path_in_zip, payslip_data['content'])
    
    zip_buffer.seek(0)
    st.session_state.yearly_zip_content = zip_buffer.getvalue()
    st.session_state.yearly_zip_filename = f"bulletins_paie_annee_{target_year}.zip"

def display_yearly_results():
    if not st.session_state.yearly_show_results:
        return
    
    collaborator_payslips = st.session_state.yearly_payslip_data
    yearly_stats = st.session_state.yearly_stats
    target_year = st.session_state.yearly_target_year
    company_info = st.session_state.yearly_company_info
    
    st.subheader(f"📊 Résultats pour l'année {target_year}")
    
    # Statistiques globales
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Total collaborateurs", yearly_stats['total_collaborators'])
    with col2:
        st.metric("Avec bulletins", yearly_stats['collaborators_with_payslips'])
    with col3:
        st.metric("Total bulletins", yearly_stats['total_payslips_found'])
    with col4:
        st.metric("Mois couverts", len(yearly_stats['months_processed']))
    
    # Bouton de téléchargement global
    if 'yearly_zip_content' in st.session_state and st.session_state.yearly_zip_content:
        st.download_button(
            label=f"📥 Télécharger tous les bulletins {target_year} (ZIP)",
            data=st.session_state.yearly_zip_content,
            file_name=st.session_state.yearly_zip_filename,
            mime="application/zip",
            help=f"Archive contenant {yearly_stats['total_payslips_found']} bulletins organisés par collaborateur"
        )
    
    # Détail par collaborateur
    st.subheader("👥 Détail par collaborateur")
    
    if collaborator_payslips:
        # Créer un tableau récapitulatif
        summary_data = []
        for collab_name, months_data in collaborator_payslips.items():
            months_list = sorted([int(m) for m in months_data.keys()])
            months_str = ", ".join([datetime(2000, m, 1).strftime("%B") for m in months_list])
            
            summary_data.append({
                "Collaborateur": collab_name,
                "Nombre de bulletins": len(months_data),
                "Mois disponibles": months_str
            })
        
        df_summary = pd.DataFrame(summary_data)
        st.dataframe(df_summary, use_container_width=True)
        
        # Téléchargements individuels par collaborateur
        with st.expander("📁 Téléchargements individuels par collaborateur", expanded=False):
            for collab_name, months_data in collaborator_payslips.items():
                st.write(f"**{collab_name}** - {len(months_data)} bulletin(s)")
                
                # Créer un ZIP spécifique pour ce collaborateur
                collab_zip_buffer = io.BytesIO()
                with zipfile.ZipFile(collab_zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:
                    for month, payslip_data in months_data.items():
                        zipf.writestr(payslip_data['file_name'], payslip_data['content'])
                
                collab_zip_buffer.seek(0)
                
                col1, col2 = st.columns([3, 1])
                with col1:
                    months_list = sorted([int(m) for m in months_data.keys()])
                    months_str = ", ".join([datetime(2000, m, 1).strftime("%b") for m in months_list])
                    st.write(f"Mois : {months_str}")
                with col2:
                    st.download_button(
                        label="📥 ZIP collaborateur",
                        data=collab_zip_buffer.getvalue(),
                        file_name=f"{collab_name.replace(' ', '_')}_bulletins_{target_year}.zip",
                        mime="application/zip",
                        key=f"collab_{collab_name.replace(' ', '_')}"
                    )
    else:
        st.warning(f"Aucun bulletin trouvé pour l'année {target_year}")

# ==================== FONCTIONS BULLETINS MENSUELS ====================

def get_payslips(api_key, target_year, target_month):
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
            
            status_placeholder.info(f"📥 Récupération des collaborateurs... Page {page_count} ({len(all_collabs)} collaborateurs jusqu'à présent)")
            
            next_page_token = collabs_response.get("meta", {}).get("nextPageToken")
            if not next_page_token:
                break
        
        collabs = all_collabs
        status_placeholder.success(f"✅ {len(collabs)} collaborateurs récupérés.")
        progress_bar.progress(30)
        
        # Affichage de tous les collaborateurs
        with st.expander("👥 Liste complète des collaborateurs", expanded=True):
            if collabs:
                collabs_data = []
                for collab in collabs:
                    collab_data = {
                        "ID": collab.get("id", ""),
                        "Prénom": collab.get("firstName", ""),
                        "Nom": collab.get("lastName", ""),
                        "Email": collab.get("email", ""),
                        "Statut": "Actif" if collab.get("status") == "active" else "Inactif"
                    }
                    if "startDate" in collab:
                        collab_data["Date de début"] = collab["startDate"]
                    if "endDate" in collab:
                        collab_data["Date de fin"] = collab["endDate"]
                    
                    collabs_data.append(collab_data)
                
                df = pd.DataFrame(collabs_data)
                st.dataframe(df, use_container_width=True)
                
                csv_data, csv_filename = create_csv_download(
                    df, 
                    f"collaborateurs_{company_info['name']}_{datetime.now().strftime('%Y%m%d')}.csv"
                )
                
                st.session_state.csv_data = csv_data
                st.session_state.csv_filename = csv_filename
                st.session_state.show_download_button = True
            else:
                st.warning("Aucun collaborateur trouvé.")
                st.session_state.show_download_button = False
        
        os.makedirs("bulletins_paie", exist_ok=True)
        
        # 4. Filtrage + extraction
        status_placeholder.info("🔍 Recherche des bulletins de paie...")
        
        collabs_with_payslip = []
        collabs_without_payslip = []
        payslip_data = {}
        
        with st.expander("Détails du traitement", expanded=False):
            details_placeholder = st.empty()
            detail_text = ""
        
        total_collabs = len(collabs)
        for i, collab in enumerate(collabs):
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
                    file_name = f"{file_safe_name}_{target_year}_{target_month}.pdf"
                    
                    extracted_content = extract_second_page(pdf_response.content)
                    if extracted_content:
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
        
        # Stockage des données dans la session_state
        st.session_state.payslip_data = payslip_data
        st.session_state.collabs_with_payslip = collabs_with_payslip
        st.session_state.collabs_without_payslip = collabs_without_payslip
        st.session_state.target_year = target_year
        st.session_state.target_month = target_month
        st.session_state.show_results = True
        
        if payslip_data:
            zip_content = create_zip_in_memory(payslip_data)
            st.session_state.zip_content = zip_content
            st.session_state.zip_filename = f"bulletins_paie_{target_year}_{target_month}.zip"
        
    except Exception as e:
        st.error(f"Une erreur est survenue: {str(e)}")
        import traceback
        st.error(traceback.format_exc())
        st.session_state.traitement_termine = False
        st.session_state.show_results = False

# ==================== INITIALISATION DES VARIABLES DE SESSION ====================

# Variables pour bulletins mensuels
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

# Variables pour bulletins annuels
if 'yearly_payslip_data' not in st.session_state:
    st.session_state.yearly_payslip_data = {}
if 'yearly_show_results' not in st.session_state:
    st.session_state.yearly_show_results = False
if 'yearly_target_year' not in st.session_state:
    st.session_state.yearly_target_year = None
if 'yearly_stats' not in st.session_state:
    st.session_state.yearly_stats = {}
if 'yearly_zip_content' not in st.session_state:
    st.session_state.yearly_zip_content = None
if 'yearly_zip_filename' not in st.session_state:
    st.session_state.yearly_zip_filename = ""
if 'yearly_company_info' not in st.session_state:
    st.session_state.yearly_company_info = {}

# ==================== INTERFACE PRINCIPALE ====================

st.title("📄 Payfit - Récupération des bulletins de paie")
st.write("Application complète pour télécharger les bulletins de paie de vos collaborateurs.")

# Navigation par onglets
tab1, tab2 = st.tabs(["📅 Bulletins par mois", "📆 Bulletins par année"])

# ==================== ONGLET 1: BULLETINS MENSUELS ====================

with tab1:
    st.header("📅 Récupération des bulletins par mois")
    st.write("Téléchargez la deuxième page des bulletins de paie de vos collaborateurs pour un mois spécifique et obtenez des informations détaillées sur votre entreprise et vos collaborateurs.")
    
    with st.form(key="payslip_form"):
        api_key = st.text_input("🔐 Clé API Payfit", type="password", help="Vous pouvez obtenir une clé API depuis votre compte Payfit")
        
        col1, col2 = st.columns(2)
        with col1:
            current_year = datetime.now().year
            years = list(range(current_year-3, current_year+1))
            target_year = st.selectbox("📅 Année", options=years, index=len(years)-1)
        
        with col2:
            months = [(str(i).zfill(2), datetime(2000, i, 1).strftime("%B")) for i in range(1, 13)]
            current_month = datetime.now().month - 1
            target_month = st.selectbox(
                "📅 Mois", 
                options=[m[0] for m in months],
                format_func=lambda x: months[int(x)-1][1],
                index=current_month - 1 if current_month > 0 else 0
            )
        
        submit_button = st.form_submit_button(label="📥 Récupérer les bulletins")
        
        if submit_button:
            get_payslips(api_key, str(target_year), target_month)
            del api_key
    
    # Bouton de téléchargement CSV
    if st.session_state.show_download_button and 'csv_data' in st.session_state and 'csv_filename' in st.session_state:
        st.download_button(
            label="📥 Télécharger la liste des collaborateurs (CSV)",
            data=st.session_state.csv_data,
            file_name=st.session_state.csv_filename,
            mime="text/csv",
        )
    
    # Affichage des résultats mensuels
    if st.session_state.show_results:
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
            
            st.write("---")
            st.subheader("Bulletins individuels")
            
            for name in payslip_data.keys():
                col1, col2 = st.columns([3, 1])
                with col1:
                    st.write(f"**{name}**")
                with col2:
                    key = f"dl_btn_{name.replace(' ', '_')}"
                    st.download_button(
                        label="Télécharger bulletin de paie",
                        data=payslip_data[name]["content"],
                        file_name=payslip_data[name]["file_name"],
                        mime="application/pdf",
                        key=key
                    )
        else:
            st.warning("Aucun bulletin trouvé pour ce mois.")
        
        # Informations sur les collaborateurs sans bulletin
        with st.expander("ℹ️ Collaborateurs sans bulletin pour cette période", expanded=True):
            if collabs_without_payslip:
                missing_data = []
                for collab in collabs_without_payslip:
                    missing_item = {
                        "Nom": collab["name"],
                        "Raison": collab["reason"]
                    }
                    if "available_periods" in collab:
                        missing_item["Périodes disponibles"] = ", ".join(collab["available_periods"])
                    
                    missing_data.append(missing_item)
                
                df_missing = pd.DataFrame(missing_data)
                st.dataframe(df_missing, use_container_width=True)
            else:
                st.success("Tous les collaborateurs ont un bulletin pour cette période!")

# ==================== ONGLET 2: BULLETINS ANNUELS ====================

with tab2:
    st.header("📆 Récupération des bulletins de paie par année")
    st.write("Récupérez tous les bulletins de paie de vos collaborateurs pour une année complète.")
    
    # Initialisation des variables de session pour la page annuelle
    if 'yearly_show_results' not in st.session_state:
        st.session_state.yearly_show_results = False
    if 'yearly_payslip_data' not in st.session_state:
        st.session_state.yearly_payslip_data = {}
    
    with st.form(key="yearly_payslip_form"):
        api_key = st.text_input("🔐 Clé API Payfit", type="password", help="Vous pouvez obtenir une clé API depuis votre compte Payfit")
        
        current_year = datetime.now().year
        years = list(range(current_year-5, current_year+1))
        target_year = st.selectbox("📅 Année", options=years, index=len(years)-2)  # Année précédente par défaut
        
        submit_button = st.form_submit_button(label="📥 Récupérer tous les bulletins de l'année")
        
        if submit_button:
            get_yearly_payslips(api_key, target_year)
            del api_key
    
    # Affichage des résultats
    display_yearly_results()

# ==================== SECTION D'AIDE ====================

with st.expander("ℹ️ Guide d'utilisation", expanded=False):
    st.write("""
    ### 🚀 Comment utiliser cette application
    
    #### 📅 Onglet "Bulletins par mois"
    1. **Clé API** : Obtenez votre clé API depuis votre compte administrateur Payfit
    2. **Sélection de la période** : Choisissez l'année et le mois pour lesquels vous souhaitez récupérer les bulletins
    3. **Lancement** : Cliquez sur le bouton pour commencer le processus
    4. **Exploration des résultats** : Une fois le traitement terminé, vous pourrez :
       - Voir les informations détaillées de votre entreprise
       - Consulter la liste complète de vos collaborateurs
       - Télécharger tous les bulletins en une seule fois
       - Télécharger les bulletins individuellement
       - Voir quels collaborateurs n'ont pas de bulletin pour la période sélectionnée
    
    ### 📁 Organisation des fichiers
    
    - **Bulletins mensuels** : Fichiers nommés `Prenom_Nom_ANNEE_MOIS.pdf`
    - **Extraction** : Seules les 2èmes pages des bulletins sont incluses
    
    ### 🔒 Sécurité
    
    - Les clés API sont automatiquement supprimées de la mémoire après utilisation
    - Aucun stockage permanent des données sensibles
    - Traitement entièrement en mémoire pour les PDF
    """)

# Pied de page
st.divider()
col1, col2 = st.columns(2)
with col1:
    st.caption("📄 **Note :** Seules les 2èmes pages des bulletins sont extraites")
with col2:
    st.caption("🔒 **Sécurité :** Les clés API sont automatiquement supprimées de la mémoire")