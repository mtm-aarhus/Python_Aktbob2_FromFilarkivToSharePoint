import pandas as pd
import requests
import json
import os
import time
from datetime import datetime
from SharePointUploader import upload_file_to_sharepoint
from SendSMTPMail import send_email
from office365.sharepoint.client_context import ClientContext
import sys

def download_files(FilarkivURL, FilarkivCaseID, Filarkiv_access_token, Sagsnummer, MailModtager, dt_AktIndex ):
    url = f"{FilarkivURL}/Documents/CaseDocumentOverview?caseId={FilarkivCaseID}&pageIndex=1&pageSize=500"
    
    headers = {
        "Authorization": f"Bearer {Filarkiv_access_token}",
        "Content-Type": "application/xml"
    }
    try:
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        document_list = response.json()      
    except Exception as e:
        print("Exception occurred:", str(e))

        # Define email details
        sender = "aktbob@aarhus.dk"
        subject = f"{Sagsnummer} mangler dokumentliste"
        body = f"""Kære sagsbehandler,<br><br>
        Sagen: {Sagsnummer} mangler at få overført dokumenter til screeningsmappen. <br><br>
        Få overfør dokumenterne først, inden du forsøger steppet 'Overfør dokumenter til udleveringsmappe (Sharepoint)'.<br><br>
        Det anbefales at følge <a href="https://aarhuskommune.atlassian.net/wiki/spaces/AB/pages/64979049/AKTBOB+--+Vejledning">vejledningen</a>, 
        hvor du også finder svar på de fleste spørgsmål og fejltyper.
        """

        smtp_server = "smtp.adm.aarhuskommune.dk"
        smtp_port = 25

        # Send the error notification
        send_email(
            receiver=MailModtager,
            sender=sender,
            subject=subject,
            body=body,
            smtp_server=smtp_server,
            smtp_port=smtp_port,
            html_body=True
        )
    
    downloaded_files = []
    
    for document in document_list:
        documentreference = str(document.get('documentReference', None)) #Checking for doc id, which we gave the document at upload
        for file in document.get("files", []):
            file_id = file["id"]
            file_name = file["fileName"]
            
            file_basename = file_name.rsplit(".", 1)[0]
            
            # akt_id_from_file = str(int(file_name[:4]))
            # matching_rows = dt_AktIndex[dt_AktIndex["Akt ID"].astype(str) == akt_id_from_file]
            if documentreference:
                matching_rows = dt_AktIndex[dt_AktIndex["Dok ID"].str.contains(documentreference, na=False, regex=False)]
            if not documentreference: #matching at name if there is no doc id to avoid errors in the transition period
                matching_rows = dt_AktIndex[dt_AktIndex["Filnavn"].str.contains(file_basename, na=False, regex=False)]
            if matching_rows.empty: 
                parts = file_basename.split(" - ")
                if len(parts) >= 2:
                    value_between = parts[1]
                    matching_rows = dt_AktIndex[dt_AktIndex["Dok ID"].str.contains(value_between, na=False, regex=False)]

            if matching_rows.empty: 
                raise ValueError(f"No matching row for dokumenttitle {file_basename}")
            
            for index, row in matching_rows.iterrows():
                if "Ja" in row["Gives der aktindsigt?"] or "Delvis" in row["Gives der aktindsigt?"]:
                    dt_AktIndex.at[index, "Filnavn"] = f"{file_basename}.pdf"
                    
                    download_url = f"{FilarkivURL}/FileIO/Download?fileId={file_id}"
                    file_path = os.path.join("C:\\Users", os.getlogin(), "Downloads", file_name)
                    print(f'Getting {file_name}')
                    response = requests.get(download_url, headers=headers)
                    if response.status_code == 404:
                        continue
                    response.raise_for_status()

                    with open(file_path, "wb") as f:
                        f.write(response.content)
                        downloaded_files.append(file_path)  
    return downloaded_files

def DownloadFilesFromFilarkivAndUploadToSharePoint( FilarkivURL, Filarkiv_access_token, dt_AktIndex, FilarkivCaseID, SharePointURL, Overmappe, Undermappe, MailModtager, Sagsnummer, tenant, client_id, thumbprint, cert_path, orchestrator_connection):
    
    cert_credentials = {
        "tenant": tenant,
        "client_id": client_id,
        "thumbprint": thumbprint,
        "cert_path": cert_path
    }
    ctx = ClientContext(SharePointURL).with_client_certificate(**cert_credentials)

    # Load the SharePoint web to test the connection
    web = ctx.web
    ctx.load(web)
    ctx.execute_query()

    downloaded_files = download_files(FilarkivURL, FilarkivCaseID, Filarkiv_access_token, Sagsnummer, MailModtager, dt_AktIndex )
    if downloaded_files:
        for file_path in downloaded_files:
            upload_file_to_sharepoint(
                    site_url=SharePointURL,
                    Overmappe=Overmappe,
                    Undermappe=Undermappe,
                    file_path=file_path,
                    ctx = ctx
                )
            try:
                os.remove(file_path)
            except Exception as e:
                raise Exception(f"Error deleting {file_path}: {str(e)}")
    else:
        orchestrator_connection.log_info('Ingen filer er blevet downloadet, slutter processen')
        return dt_AktIndex

    return dt_AktIndex
