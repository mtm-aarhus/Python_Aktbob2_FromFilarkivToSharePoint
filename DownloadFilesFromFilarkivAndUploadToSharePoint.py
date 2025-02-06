def invoke_DownloadFilesFromFilarkivAndUploadToSharePoint(Arguments_DownloadFilesFromFilarkivAndUploadToSharePoint):
    import pandas as pd
    import requests
    import json
    import os
    import time
    from datetime import datetime
    from SharePointUploader import upload_file_to_sharepoint

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
                print("Failed to fetch document list from Filarkiv:", response.status_code)
                return []
        except Exception as e:
            print("Could not fetch document list from Filarkiv:", str(e))
            return []
        
        downloaded_files = []
        
        for document in document_list:
            for file in document.get("files", []):
                file_id = file["id"]
                file_name = file["fileName"]
                
                file_basename = file_name.replace(".pdf", "")
                matching_rows = dt_AktIndex[dt_AktIndex["Filnavn"].str.contains(file_basename, na=False, regex=False)]
                
                for _, row in matching_rows.iterrows():
                    if "Ja" in row["Gives der aktindsigt?"] or "Delvis" in row["Gives der aktindsigt?"]:
                        download_url = f"{FilarkivURL}/FileIO/Download?fileId={file_id}"
                        file_path = os.path.join("C:\\Users", os.getlogin(), "Downloads", file_name)
                        
                        try:
                            response = requests.get(download_url, headers=headers)
                            if response.status_code == 200:
                                with open(file_path, "wb") as f:
                                    f.write(response.content)
                                print(f"Downloaded: {file_name} to {file_path}")
                                downloaded_files.append(file_path)
                            else:
                                print(f"Failed to download {file_name}: {response.status_code}")
                        except Exception as e:
                            print(f"Error downloading {file_name}: {str(e)}")
        
        return downloaded_files

    def upload_files_to_sharepoint(files):
        for file_path in files:
            try:
                upload_file_to_sharepoint(
                    site_url=SharePointURL,
                    overmappe=Overmappe,
                    undermappe=Undermappe,
                    file_path=file_path,
                    sharepoint_app_id=SharePointAppID,
                    sharepoint_tenant=SharePointTenant,
                    robot_username=RobotUserName,
                    robot_password=RobotPassword
                )
            except Exception as e:
                print(f"Error uploading {file_path}: {str(e)}")

    def delete_local_files(files):
        for file_path in files:
            try:
                os.remove(file_path)
            except Exception as e:
                print(f"Error deleting {file_path}: {str(e)}")

    # Execute the workflow

    downloaded_files = download_files()
    if downloaded_files:
        upload_files_to_sharepoint(downloaded_files)
        #delete_local_files(downloaded_files)


    return {"out_Text": "Alle filer er downloaded og oploaded til Sharepoint"}
