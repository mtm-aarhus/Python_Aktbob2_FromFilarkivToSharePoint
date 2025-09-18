import requests
import uuid
import re
import html
import unicodedata
import pandas as pd
from datetime import datetime
import os
import json
import os
import time
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File
from office365.runtime.auth.user_credential import UserCredential
from urllib.parse import quote
import re
from SendSMTPMail import send_email
from SharePointUploader import sharepoint_client, download_file_from_sharepoint
def sanitize_sagstitel(sagstitel):
    try:
        sagstitel = html.unescape(sagstitel)
        sagstitel = unicodedata.normalize('NFKC', sagstitel)
        sagstitel = sagstitel.replace('"', '')
        sagstitel = re.sub(r'[.:>#<*\?/%&{}\[\]\$!"@+\|\'=€]+', '', sagstitel)
        sagstitel = sagstitel.replace('\n', '').replace('\r', '')
        sagstitel = re.sub(r'[^a-zA-Z0-9ÆØÅæøå ]', '', sagstitel)
        sagstitel = re.sub(r' {2,}', ' ', sagstitel)
        sagstitel = sagstitel.strip()
                # Check length and truncate if necessary
        if len(sagstitel) > 49:
            sagstitel = sagstitel[:50].strip()
        
        return sagstitel

    except Exception as e:
        raise Exception(f"Error during sanitization: {str(e)}")
    
def get_document_list(Sagsnummer, GeoSag, NovaSag, KMD_access_token, KMDNovaURL, SharePointUrl, Overmappe, Undermappe, MailModtager, tenant, client_id, thumbprint, cert_path, go_session): 

    sagstitel = ""  # Default value if no title is retrieved

    # --- Check if it's a Geo-sag ---
    if GeoSag:
        print("Sagen er en Geo-sag, henter derfor sagstitel i GO")
        url = f"https://ad.go.aarhuskommune.dk/_goapi/Cases/Metadata/{Sagsnummer}"

        try:
            response = go_session.get(url)
            print(f"Geo API Response Status Code: {response.status_code}")

            response_data = response.json()
            metadata = response_data.get("Metadata")

            sagstitel = metadata.split('ows_Title="')[1].split('"')[0]
            print("Sagstitel (Geo):", sagstitel)
        except Exception as e:
            raise Exception("Failed to extract Sagstitel (Geo):", str(e))

    # --- Check if it's a Nova-sag ---
    elif NovaSag:
        print("Sagen er en Novasag, henter Sagstitel i NOVA")
        TransactionID = str(uuid.uuid4())
        url = f"{KMDNovaURL}/Case/GetList?api-version=2.0-Case"

        headers = {
            "Authorization": f"Bearer {KMD_access_token}",
            "Content-Type": "application/json"
        }

        payload = {
            "common": {
                "transactionId": TransactionID
            },
            "paging": {
                "startRow": 1,
                "numberOfRows": 100
            },
            "caseAttributes": {
                "userFriendlyCaseNumber": Sagsnummer
            },
            "caseGetOutput": {
                "caseAttributes": {
                    "title": True,
                    "userFriendlyCaseNumber": True
                }
            }
        }

        try:
            response = requests.put(url, headers=headers, json=payload)
            response.raise_for_status()
            
            sagstitel = response.json()['cases'][0]['caseAttributes']['title']
            print("Sagstitel (Nova):", sagstitel)
        except Exception as e:
            raise e

    # Sanitize sagstitel regardless of source or failure
    sagstitel = sanitize_sagstitel(sagstitel)
    print(f"Final Sanitized Sagstitel: {sagstitel}")

    # ---- Henter dokumentlisten fra Sharepoint ----
    site_relative_path = "/Teams/tea-teamsite10506/Delte Dokumenter"
    
    # Main logic
    try:
        client = sharepoint_client(tenant, client_id, thumbprint, cert_path, SharePointUrl)

        # Construct paths for Overmappe and Undermappe without over-encoding
        overmappe_url = f"{site_relative_path}/Dokumentlister/{Overmappe}"
        overmappe_folder = client.web.get_folder_by_server_relative_url(overmappe_url)
        client.load(overmappe_folder)
        client.execute_query()

        undermappe_url = f"{overmappe_url}/{Undermappe}"
        print(f'undermappe {undermappe_url}')
        undermappe_folder = client.web.get_folder_by_server_relative_url(undermappe_url)
        client.load(undermappe_folder)
        client.execute_query()


        # Fetch files in the Undermappe folder
        print("Fetching files from the folder...")
        files = undermappe_folder.files
        client.load(files)
        client.execute_query()

        # Print and process file names
        date_re = re.compile(r"(\d{2}-\d{2}-\d{4})")
        data_table = []
        for file in files:
            file_name = file.properties["Name"]
            date_from_name = date_re.search(file_name)
            if not date_from_name:
                raise ValueError(f"Mangler dato i filnavn: {file_name}")

            date_str = date_from_name.group(1)
            try:
                dt = datetime.strptime(date_str, "%d-%m-%Y")
            except ValueError as e:
                raise ValueError(f"Ugyldig dato '{date_str}' i filnavn: {file_name}") from e

            data_table.append({
                "FileName": file_name,
                "DocumentDate": dt.strftime("%d-%m-%Y")})

        # Sort files by date in descending order
        data_table = sorted(
            data_table, 
            key=lambda x: datetime.strptime(x["DocumentDate"], "%d-%m-%Y"), 
            reverse=True
        )

        # Download the newest file if available
        if data_table:
            newest_file = data_table[0]
            newest_file_name = newest_file["FileName"]
            DokumentlisteDatoString = newest_file["DocumentDate"] 
            sharepoint_file_url = f"{undermappe_url}/{newest_file_name}"
            local_file_path = download_file_from_sharepoint(client, sharepoint_file_url, download_path= os.getcwd())

        if local_file_path.endswith('.xlsx'):
            # Read Excel file into a Pandas DataFrame
            dt_DocumentList = pd.read_excel(local_file_path)
            os.remove(local_file_path)
        else:
            print(f"Downloaded file is not an Excel file: {local_file_path}")
    except Exception as e:
        # --- HANDLE ERROR HERE ---
        print("Failed - sender mail til sagsbehandler")  # Replace with email logic if needed
    
        # Define email details
        sender = "aktbob@aarhus.dk" 
        subject = f"{Sagsnummer} mangler dokumentliste"
        body = f"""Kære sagsbehandler,<br><br>
        Sagen: {Sagsnummer} mangler at få oprettet dokumentlisten. <br><br>
        Få oprettet denne først, inden du forsøger steppet 'Overfør dokumenter til udleveringsmappe (Sharepoint)'.<br><br>
        Det anbefales at følge <a href="https://aarhuskommune.atlassian.net/wiki/spaces/AB/pages/64979049/AKTBOB+--+Vejledning">vejledningen</a>, 
        hvor du også finder svar på de fleste spørgsmål og fejltyper.
        """
        smtp_server = "smtp.adm.aarhuskommune.dk"   
        smtp_port = 25               

        # Call the send_email function
        send_email(
            receiver=MailModtager,
            sender=sender,
            subject=subject,
            body=body,
            smtp_server=smtp_server,
            smtp_port=smtp_port,
            html_body=True
        )

        return None

    return sagstitel, dt_DocumentList, DokumentlisteDatoString

