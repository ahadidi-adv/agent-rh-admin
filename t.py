import streamlit as st
import requests
import json
import os
import base64
import pandas as pd
from datetime import datetime

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

# Fonction pour afficher des informations dans un expander
def display_dict_info(title, data_dict, icon="ℹ️"):
    with st.expander(f"{icon} {title}", expanded=False):
        if isinstance(data_dict, list):
            for item in data_dict:
                for key, value in item.items():
                    if isinstance(value, (dict, list)):
                        st.write(f"**{key}** :")
                        st.json(value)
                    else:
                        st.write(f"**{key}** : {value}")
                st.markdown("---")
        else:
            for key, value in data_dict.items():
                if isinstance(value, (dict, list)):
                    st.write(f"**{key}** :")
                    st.json(value)
                else:
                    st.write(f"**{key}** : {value}")

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
        
        # Afficher les informations d'authentification
        display_dict_info("Informations d'authentification", data, "🔑")
        
        progress_bar.progress(10)
        
        # 2. Infos de l'entreprise
        st.subheader("🏢 Informations de l'entreprise")
        company_info = requests.get(f"{BASE_URL}/companies/{company_id}", headers=headers_auth).json()
        
        # Afficher un résumé des informations principales
        company_summary = {
            "Nom": company_info.get('name', 'N/A'),
            "ID": company_id,
            "Contrats actifs": company_info.get('nbActiveContracts', 'N/A'),
            "Adresse": f"{company_info.get('address', {}).get('street', '')}, {company_info.get('address', {}).get('zipCode', '')} {company_info.get('address', {}).get('city', '')}",
            "Pays": company_info.get('address', {}).get('country', 'N/A'),
            "SIRET": company_info.get('siret', 'N/A'),
            "NAF": company_info.get('naf', 'N/A')
        }
        
        # Afficher dans des colonnes les informations principales
        col1, col2 = st.columns(2)
        with col1:
            st.write(f"**Nom :** {company_summary['Nom']}")
            st.write(f"**SIRET :** {company_summary['SIRET']}")
            st.write(f"**Code NAF :** {company_summary['NAF']}")
        with col2:
            st.write(f"**Adresse :** {company_summary['Adresse']}")
            st.write(f"**Pays :** {company_summary['Pays']}")
            st.write(f"**Contrats actifs :** {company_summary['Contrats actifs']}")
        
        # Afficher toutes les informations détaillées dans un expander
        display_dict_info("Détails complets de l'entreprise", company_info, "🏢")
        
        progress_bar.progress(20)
        
        # 3. Récupération des collaborateurs
        st.subheader("👥 Collaborateurs")
        status_placeholder.info("📥 Récupération des collaborateurs...")
        
        # Récupération de tous les collaborateurs avec pagination
        all_collabs = []
        next_page_token = None
        page_count = 0
        
        # Ajout de logs pour déboguer la pagination
        debug_logs = []
        
        try:
            # Première page - récupération initiale sans token
            page_count += 1
            collaborators_url = f"{BASE_URL}/companies/{company_id}/collaborators"
            debug_logs.append(f"URL (page 1): {collaborators_url}")
            
            # Récupérer la première page
            response = requests.get(collaborators_url, headers=headers_auth)
            
            if response.status_code != 200:
                st.error(f"❌ Erreur lors de la récupération des collaborateurs. Code: {response.status_code}")
                return
            
            # Analyser la première réponse
            collabs_response = response.json()
            debug_logs.append(f"Réponse API - Clés: {list(collabs_response.keys())}")
            
            if "collaborators" in collabs_response:
                all_collabs.extend(collabs_response["collaborators"])
                debug_logs.append(f"Collaborateurs trouvés sur la page 1: {len(collabs_response['collaborators'])}")
            else:
                st.error("❌ Format de réponse inattendu - pas de collaborateurs trouvés")
                st.json(collabs_response)
                return
            
            # Vérifier si la pagination est présente - basé sur les logs précédents
            has_more_pages = False
            if "meta" in collabs_response and "nextPageToken" in collabs_response["meta"] and collabs_response["meta"]["nextPageToken"]:
                next_page_token = collabs_response["meta"]["nextPageToken"]
                debug_logs.append(f"Token pour la page suivante: {next_page_token}")
                has_more_pages = True
            
            # Si plus de pages sont disponibles, essayer d'autres méthodes de pagination basées sur l'erreur précédente
            if has_more_pages:
                # Méthode 1: essayer avec le paramètre 'nextPageToken'
                page_count += 1
                pagination_methods = [
                    {"name": "nextPageToken", "param": "nextPageToken"},
                    {"name": "page_token", "param": "page_token"},
                    {"name": "pageToken dans l'URL", "url_param": True}
                ]
                
                # Tester différentes méthodes de pagination basées sur l'erreur précédente
                for method in pagination_methods:
                    status_placeholder.info(f"📥 Test de pagination avec la méthode '{method['name']}'...")
                    
                    if method.get("url_param", False):
                        # Méthode 3: Essayer d'ajouter le token directement dans l'URL
                        next_url = f"{collaborators_url}?pageToken={next_page_token}"
                        debug_logs.append(f"Test méthode '{method['name']}' - URL: {next_url}")
                        response = requests.get(next_url, headers=headers_auth)
                    else:
                        # Méthode 1 et 2: Utiliser les paramètres de requête
                        param_name = method["param"]
                        params = {param_name: next_page_token}
                        debug_logs.append(f"Test méthode '{method['name']}' - Params: {params}")
                        response = requests.get(collaborators_url, headers=headers_auth, params=params)
                    
                    # Vérifier si cette méthode a fonctionné
                    if response.status_code == 200:
                        try:
                            page_response = response.json()
                            if "collaborators" in page_response and len(page_response["collaborators"]) > 0:
                                # Bingo! Cette méthode fonctionne
                                debug_logs.append(f"🎯 La méthode '{method['name']}' fonctionne!")
                                all_collabs.extend(page_response["collaborators"])
                                debug_logs.append(f"Collaborateurs supplémentaires trouvés: {len(page_response['collaborators'])}")
                                st.success(f"✅ Méthode de pagination '{method['name']}' trouvée et fonctionnelle!")
                                break
                            else:
                                debug_logs.append(f"La méthode '{method['name']}' n'a pas renvoyé de collaborateurs.")
                        except Exception as e:
                            debug_logs.append(f"Erreur en essayant la méthode '{method['name']}': {str(e)}")
                    else:
                        debug_logs.append(f"La méthode '{method['name']}' a échoué avec le code {response.status_code}")
                        debug_logs.append(response.text)
                
                # Avertissement si aucune méthode n'a fonctionné et s'il y a potentiellement plus de pages
                st.warning(f"⚠️ La pagination a été interrompue. {len(all_collabs)} collaborateurs ont été récupérés, mais il pourrait y en avoir d'autres. Consultez les logs pour plus de détails.")
            
            # Définir collabs pour la suite du traitement
            collabs = all_collabs
            status_placeholder.success(f"✅ {len(collabs)} collaborateurs récupérés.")
            
            # Option pour voir les logs de débogage
            with st.expander("Logs de débogage de la pagination", expanded=False):
                for log in debug_logs:
                    st.write(log)
            
            # DOCUMENTATION POUR UNE IMPLÉMENTATION FUTURE
            with st.expander("🛠️ Note pour le développeur sur la pagination", expanded=False):
                st.markdown("""
                ## Problème de pagination
                
                L'API Payfit retourne un token `nextPageToken` dans la réponse, mais le paramètre exact à utiliser pour la pagination n'est pas clair:
                
                1. Essai avec `pageToken` → Erreur 400: "Unknown query parameter 'pageToken'"
                2. Essai avec d'autres paramètres potentiels → Résultats variables
                
                ### Options pour résoudre ce problème:
                
                1. **Consulter la documentation spécifique de l'API Payfit** pour comprendre le mécanisme exact de pagination
                2. **Contacter le support technique de Payfit** pour obtenir des informations précises sur la pagination
                3. **Implémenter une méthode robuste** qui teste automatiquement différentes approches de pagination
                
                ### Référence:
                
                La documentation de référence est disponible à: https://developers.payfit.io/reference/get_contracts
                """)
        
        except Exception as e:
            st.error(f"Une erreur s'est produite lors de la récupération des collaborateurs: {str(e)}")
            with st.expander("Détails de l'erreur", expanded=True):
                for log in debug_logs:
                    st.write(log)
                import traceback
                st.code(traceback.format_exc())
            
            # Définir collabs comme une liste vide en cas d'erreur pour éviter les erreurs suivantes
            collabs = all_collabs if 'all_collabs' in locals() else []
        
        # Création d'un dataframe pour afficher les collaborateurs
        collabs_data = []
        for collab in collabs:
            collabs_data.append({
                "ID": collab.get('id', 'N/A'),
                "Prénom": collab.get('firstName', 'N/A'),
                "Nom": collab.get('lastName', 'N/A'),
                "Email": collab.get('email', 'N/A'),
                "Statut": collab.get('status', 'N/A'),
                "Date d'embauche": collab.get('hireDate', 'N/A')
            })
        
        df_collabs = pd.DataFrame(collabs_data)
        
        # Afficher un résumé des collaborateurs
        st.write(f"**Total des collaborateurs :** {len(collabs)}")
        
        # Filtres pour les collaborateurs
        with st.expander("🔍 Filtrer les collaborateurs", expanded=False):
            search_term = st.text_input("Rechercher par nom ou prénom")
            status_filter = st.multiselect("Filtrer par statut", options=df_collabs["Statut"].unique())
            
            filtered_df = df_collabs
            
            if search_term:
                filtered_df = filtered_df[
                    filtered_df["Prénom"].str.contains(search_term, case=False) | 
                    filtered_df["Nom"].str.contains(search_term, case=False)
                ]
            
            if status_filter:
                filtered_df = filtered_df[filtered_df["Statut"].isin(status_filter)]
            
            st.dataframe(filtered_df, use_container_width=True)
        
        # Afficher tous les collaborateurs dans un expander
        with st.expander("👥 Liste complète des collaborateurs", expanded=True):
            st.dataframe(df_collabs, use_container_width=True)
        
        # Informations détaillées sur chaque collaborateur
        with st.expander("🧑‍💼 Détails des collaborateurs", expanded=False):
            selected_collab = st.selectbox(
                "Sélectionnez un collaborateur pour voir ses détails",
                options=range(len(collabs)),
                format_func=lambda i: f"{collabs[i].get('firstName', '')} {collabs[i].get('lastName', '')}"
            )
            st.json(collabs[selected_collab])
        
        status_placeholder.success(f"✅ {len(collabs)} collaborateurs récupérés.")
        progress_bar.progress(30)
        
        # Création du dossier de sortie
        os.makedirs("bulletins_paie", exist_ok=True)
        
        # 4. Filtrage + extraction
        st.subheader(f"📄 Bulletins de paie - {datetime(2000, int(target_month), 1).strftime('%B')} {target_year}")
        status_placeholder.info("🔍 Recherche des bulletins de paie...")
        
        collabs_with_payslip = []
        collabs_without_payslip = []
        download_links = []
        payslip_details = []
        
        # Expander pour les détails du traitement
        with st.expander("⚙️ Détails du traitement", expanded=False):
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
                detail_text += f"  → Aucun bulletin disponible\n"
                collabs_without_payslip.append({
                    "Nom": full_name,
                    "ID": collaborator_id,
                    "Raison": "Aucun bulletin disponible"
                })
                details_placeholder.text_area("Logs", detail_text, height=400)
                continue
            
            # Recherche d'un bulletin correspondant à la période demandée
            target_payslip = next(
                (p for p in payslip_resp["payslips"] if p["year"] == target_year and p["month"] == target_month),
                None
            )
            
            if target_payslip:
                collabs_with_payslip.append(full_name)
                
                # Collecter les détails du bulletin pour affichage
                payslip_details.append({
                    "Nom": full_name,
                    "ID Collaborateur": collaborator_id,
                    "ID Contrat": target_payslip.get("contractId", "N/A"),
                    "ID Bulletin": target_payslip.get("payslipId", "N/A"),
                    "Période": f"{target_month}/{target_year}",
                    "Date de création": target_payslip.get("creationDate", "N/A"),
                    "Statut": target_payslip.get("status", "N/A")
                })
                
                detail_text += f"  → ✅ Bulletin trouvé pour {target_month}/{target_year}\n"
                details_placeholder.text_area("Logs", detail_text, height=400)
                
                contract_id = target_payslip["contractId"]
                payslip_id = target_payslip["payslipId"]
                
                pdf_url = f"{BASE_URL}/companies/{company_id}/collaborators/{collaborator_id}/contracts/{contract_id}/payslips/{payslip_id}"
                pdf_response = requests.get(pdf_url, headers={**headers_auth, 'accept': 'application/pdf'})
                
                if pdf_response.status_code == 200:
                    file_safe_name = f"{collab.get('firstName', 'collaborateur')}_{collab.get('lastName', '')}".replace(" ", "_")
                    file_path = f"bulletins_paie/{file_safe_name}_{target_year}_{target_month}.pdf"
                    
                    with open(file_path, "wb") as f:
                        f.write(pdf_response.content)
                    
                    download_links.append((full_name, file_path))
                    detail_text += f"  → Bulletin téléchargé : {file_path}\n"
                else:
                    detail_text += f"  → ❌ Erreur lors du téléchargement du bulletin de paie (code {pdf_response.status_code})\n"
                
                details_placeholder.text_area("Logs", detail_text, height=400)
            else:
                collabs_without_payslip.append({
                    "Nom": full_name,
                    "ID": collaborator_id,
                    "Raison": f"Aucun bulletin pour {target_month}/{target_year}"
                })
                detail_text += f"  → ❌ Aucun bulletin trouvé pour {target_month}/{target_year}\n"
                details_placeholder.text_area("Logs", detail_text, height=400)
        
        progress_bar.progress(100)
        status_placeholder.success("✅ Traitement terminé!")
        
        # 5. Résumé statistique
        st.subheader("📊 Statistiques")
        stats_col1, stats_col2, stats_col3 = st.columns(3)
        
        with stats_col1:
            st.metric(
                label="Total des collaborateurs", 
                value=len(collabs)
            )
        
        with stats_col2:
            st.metric(
                label="Bulletins trouvés", 
                value=len(collabs_with_payslip),
                delta=f"{int(len(collabs_with_payslip)/len(collabs)*100)}%" if len(collabs) > 0 else "0%"
            )
        
        with stats_col3:
            st.metric(
                label="Sans bulletin", 
                value=len(collabs_without_payslip),
                delta=f"-{int(len(collabs_without_payslip)/len(collabs)*100)}%" if len(collabs) > 0 else "0%",
                delta_color="inverse"
            )
        
        # 6. Résumé des bulletins
        st.subheader("📑 Bulletins récupérés")
        
        if collabs_with_payslip:
            st.write(f"**{len(collabs_with_payslip)} bulletins trouvés pour la période {target_month}/{target_year}:**")
            
            # Afficher les détails des bulletins
            if payslip_details:
                with st.expander("🔍 Détails des bulletins", expanded=True):
                    st.dataframe(pd.DataFrame(payslip_details), use_container_width=True)
            
            # Création d'un ZIP avec tous les bulletins
            if len(download_links) > 1:
                import zipfile
                zip_filename = f"bulletins_paie_{target_year}_{target_month}.zip"
                with zipfile.ZipFile(zip_filename, 'w') as zipf:
                    for _, file_path in download_links:
                        zipf.write(file_path, arcname=os.path.basename(file_path))
                
                st.success(f"✅ Archive ZIP créée contenant {len(download_links)} bulletins de paie.")
                st.markdown(get_binary_file_downloader_html(zip_filename, "tous les bulletins (ZIP)"), unsafe_allow_html=True)
                st.write("---")
            
            # Tableau avec les liens de téléchargement individuels
            st.write("### 📥 Téléchargements individuels")
            for name, file_path in download_links:
                col1, col2 = st.columns([3, 1])
                with col1:
                    st.write(f"**{name}**")
                with col2:
                    st.markdown(get_binary_file_downloader_html(file_path, "bulletin"), unsafe_allow_html=True)
            
            # Afficher la liste des collaborateurs sans bulletin
            if collabs_without_payslip:
                st.subheader("⚠️ Collaborateurs sans bulletin")
                with st.expander("Afficher les détails", expanded=True):
                    st.dataframe(pd.DataFrame(collabs_without_payslip), use_container_width=True)
                    
        else:
            st.warning("⚠️ Aucun bulletin trouvé pour ce mois.")
            
            # Afficher la liste des collaborateurs sans bulletin
            if collabs_without_payslip:
                st.subheader("⚠️ Collaborateurs sans bulletin")
                with st.expander("Afficher les détails", expanded=True):
                    st.dataframe(pd.DataFrame(collabs_without_payslip), use_container_width=True)
            
    except Exception as e:
        st.error(f"Une erreur est survenue: {str(e)}")
        import traceback
        st.error(traceback.format_exc())

# Interface utilisateur
st.title("📄 Récupération des bulletins de paie Payfit")
st.write("Cet outil vous permet de télécharger les bulletins de paie de vos collaborateurs pour un mois spécifique.")

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

with st.expander("ℹ️ À propos de cet outil", expanded=False):
    st.write("""
    ### Fonctionnalités
    - Récupération des bulletins de paie depuis l'API Payfit
    - Téléchargement individuel ou groupé (format ZIP)
    - Visualisation des informations de l'entreprise
    - Liste détaillée des collaborateurs
    - Statistiques sur les bulletins disponibles
    
    ### Confidentialité
    - Les bulletins de paie sont téléchargés localement dans le dossier 'bulletins_paie'
    - Votre clé API n'est jamais stockée
    - Aucune donnée n'est envoyée à des serveurs tiers
    
    ### Prérequis
    - Une clé API Payfit valide
    - Des droits d'accès suffisants pour récupérer les bulletins de paie
    
    ### Limitations connues
    - La pagination des collaborateurs peut être limitée - l'application récupère au moins la première page
    - Pour les grandes entreprises, consultez les logs de pagination si tous les collaborateurs ne sont pas visibles
    """)

st.divider()
st.write("**Note :** Les bulletins de paie sont téléchargés localement dans le dossier 'bulletins_paie'. Cette application ne stocke pas vos données.")