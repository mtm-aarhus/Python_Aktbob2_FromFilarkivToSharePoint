def invoke_DownloadFilesFromFilarkivAndUploadToSharePoint(Arguments_DownloadFilesFromFilarkivAndUploadToSharePoint):
    import pandas as pd
    import requests
    import json
    import os
    import time
    from datetime import datetime
    from SharePointUploader import upload_file_to_sharepoint
    from SendSMTPMail import send_email

    FilarkivURL = Arguments_DownloadFilesFromFilarkivAndUploadToSharePoint.get("in_FilarkivURL")
    Filarkiv_access_token = Arguments_DownloadFilesFromFilarkivAndUploadToSharePoint.get("in_Filarkiv_access_token")
    dt_AktIndex = Arguments_DownloadFilesFromFilarkivAndUploadToSharePoint.get("in_dt_AktIndex")
    FilarkivCaseID = Arguments_DownloadFilesFromFilarkivAndUploadToSharePoint.get("in_FilarkivCaseID")
    SharePointAppID = Arguments_DownloadFilesFromFilarkivAndUploadToSharePoint.get("in_SharePointAppID")
    SharePointTenant = Arguments_DownloadFilesFromFilarkivAndUploadToSharePoint.get("in_SharePointTenant")
    SharePointURL = Arguments_DownloadFilesFromFilarkivAndUploadToSharePoint.get("in_SharePointURL")
    Overmappe = Arguments_DownloadFilesFromFilarkivAndUploadToSharePoint.get("in_Overmappe")
    Undermappe = Arguments_DownloadFilesFromFilarkivAndUploadToSharePoint.get("in_Undermappe")
    RobotUserName = Arguments_DownloadFilesFromFilarkivAndUploadToSharePoint.get("in_RobotUserName")
    RobotPassword = Arguments_DownloadFilesFromFilarkivAndUploadToSharePoint.get("in_RobotPassword")
    MailModtager= Arguments_DownloadFilesFromFilarkivAndUploadToSharePoint.get("in_MailModtager")
    Sagsnummer = Arguments_DownloadFilesFromFilarkivAndUploadToSharePoint.get("in_Sagsnummer")

    def download_files():
        url = f"{FilarkivURL}/Documents/CaseDocumentOverview?caseId={FilarkivCaseID}&pageIndex=1&pageSize=500"
        
        headers = {
            "Authorization": f"Bearer {Filarkiv_access_token}",
            "Content-Type": "application/xml"
        }
        
        try:
            response = requests.get(url, headers=headers)
            if response.status_code == 200:
                    document_list = response.json()
            else:
                error_msg = f"Failed to fetch document list from Filarkiv. Status code: {response.status_code}"
                raise Exception(error_msg)
                
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
            return []
        
        downloaded_files = []
        
        for document in document_list:
            for file in document.get("files", []):
                file_id = file["id"]
                file_name = file["fileName"]
                
                file_basename = file_name.rsplit(".", 1)[0]
                # akt_id_from_file = str(int(file_name[:4]))
                # matching_rows = dt_AktIndex[dt_AktIndex["Akt ID"].astype(str) == akt_id_from_file]
                matching_rows = dt_AktIndex[dt_AktIndex["Filnavn"].str.contains(file_basename, na=False, regex=False)]
                if matching_rows.empty: raise ValueError(f"No matching row for dokumenttitle {file_basename}")
                
                for index, row in matching_rows.iterrows():
                    if "Ja" in row["Gives der aktindsigt?"] or "Delvis" in row["Gives der aktindsigt?"]:
                        dt_AktIndex.at[index, "Filnavn"] = f"{file_basename}.pdf"
                        
                        download_url = f"{FilarkivURL}/FileIO/Download?fileId={file_id}"
                        file_path = os.path.join("C:\\Users", os.getlogin(), "Downloads", file_name)
                        
                        try:
                            response = requests.get(download_url, headers=headers)
                            if response.status_code == 200:
                                with open(file_path, "wb") as f:
                                    f.write(response.content)
                                downloaded_files.append(file_path)
                            else:
                                print(f"Failed to download {file_name}: {response.status_code}")
                        except Exception as e:
                            raise Exception(f"Error downloading {file_name}: {str(e)}")
        
        return downloaded_files

    def upload_files_to_sharepoint(files):
        for file_path in files:
            try:
                upload_file_to_sharepoint(
                    site_url=SharePointURL,
                    Overmappe=Overmappe,
                    Undermappe=Undermappe,
                    file_path=file_path,
                    RobotUserName=RobotUserName,
                    RobotPassword=RobotPassword
                )

            except Exception as e:
                raise Exception(f"Error uploading {file_path}: {e}")

    def delete_local_files(files):
        for file_path in files:
            try:
                os.remove(file_path)
            except Exception as e:
                raise Exception(f"Error deleting {file_path}: {str(e)}")

    # Execute the workflow

    downloaded_files = download_files()
    if downloaded_files:
        upload_files_to_sharepoint(downloaded_files)
        delete_local_files(downloaded_files)


    return {"out_Text": "Alle filer er downloaded og oploaded til Sharepoint",
            "out_dt_AktIndex": dt_AktIndex
            }
