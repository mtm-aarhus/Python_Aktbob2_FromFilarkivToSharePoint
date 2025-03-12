def invoke_GetDocumentsForAktliste(Arguments_GetDocumentsForAktliste):
    import pandas as pd
    import re
    import requests
    from requests_ntlm import HttpNtlmAuth
    import json
    import os
    import time
    from datetime import datetime
    from msal import PublicClientApplication
    from office365.sharepoint.client_context import ClientContext
    from office365.sharepoint.sharing.links.kind import SharingLinkKind
    from office365.sharepoint.webs.web import Web
    from office365.runtime.auth.user_credential import UserCredential
    import uuid
    import json
    from datetime import datetime
    import mimetypes
    
    # henter in_argumenter:
    dt_DocumentList = Arguments_GetDocumentsForAktliste.get("in_dt_Documentlist")
    CloudConvertAPI = Arguments_GetDocumentsForAktliste.get("in_CloudConvertAPI")
    MailModtager = Arguments_GetDocumentsForAktliste.get("in_MailModtager")
    UdviklerMailAktbob = Arguments_GetDocumentsForAktliste.get("in_UdviklerMail")
    RobotUserName = Arguments_GetDocumentsForAktliste.get("in_RobotUserName")
    RobotPassword = Arguments_GetDocumentsForAktliste.get("in_RobotPassword")
    FilarkivCaseID = Arguments_GetDocumentsForAktliste.get("in_FilarkivCaseID")
    SharePointAppID = Arguments_GetDocumentsForAktliste.get("in_SharePointAppID")
    SharePointTenant = Arguments_GetDocumentsForAktliste.get("in_SharePointTenant")
    SharePointURL = Arguments_GetDocumentsForAktliste.get("in_SharePointUrl")
    Overmappe = Arguments_GetDocumentsForAktliste.get("in_Overmappe")
    Undermappe = Arguments_GetDocumentsForAktliste.get("in_Undermappe")
    Sagsnummer = Arguments_GetDocumentsForAktliste.get("in_Sagsnummer")
    GeoSag = Arguments_GetDocumentsForAktliste.get("in_GeoSag")
    NovaSag = Arguments_GetDocumentsForAktliste.get("in_NovaSag")
    FilarkivURL = Arguments_GetDocumentsForAktliste.get("in_FilarkivURL")
    Filarkiv_access_token = Arguments_GetDocumentsForAktliste.get("in_Filarkiv_access_token")
    KMDNovaURL = Arguments_GetDocumentsForAktliste.get("in_KMDNovaURL")
    KMD_access_token = Arguments_GetDocumentsForAktliste.get("in_NovaToken")
    GoUsername = Arguments_GetDocumentsForAktliste.get("in_GoUsername")
    GoPassword = Arguments_GetDocumentsForAktliste.get("in_GoPassword")

    # Define the structure of the data table
    dt_AktIndex = {
        "Akt ID": pd.Series(dtype="int32"),
        "Filnavn": pd.Series(dtype="string"),
        "Dokumentkategori": pd.Series(dtype="string"),
        "Dokumentdato": pd.Series(dtype="datetime64[ns]"),
        "Dok ID": pd.Series(dtype="string"),
        "Bilag til Dok ID": pd.Series(dtype="string"),
        "Bilag": pd.Series(dtype="string"),
        "Omfattet af aktindsigt?": pd.Series(dtype="string"),
        "Gives der aktindsigt?": pd.Series(dtype="string"),
        "Begrundelse hvis Nej/Delvis": pd.Series(dtype="string"),
    }
    dt_AktIndex = pd.DataFrame(dt_AktIndex)
    #Functions: 
    def sanitize_title(Titel):
        # 1. Replace double quotes with an empty string
        Titel = Titel.replace("\"", "")

        # 2. Remove special characters with regex
        Titel = re.sub(r"[.:>#<*\?/%&{}\$!\"@+\|'=]+", "", Titel)

        # 3. Remove any newline characters
        Titel = Titel.replace("\n", "").replace("\r", "")

        # 4. Trim leading and trailing whitespace
        Titel = Titel.strip()

        # 5. Remove non-alphanumeric characters except spaces and Danish letters
        Titel = re.sub(r"[^a-zA-Z0-9ÆØÅæøå ]", "", Titel)

        # 6. Replace multiple spaces with a single space
        Titel = re.sub(r" {2,}", " ", Titel)

        return Titel
    
    def calculate_available_title_length(base_path, Overmappe, Undermappe, AktID, DokumentID, Titel, max_path_length=400):
        overmappe_length = len(Overmappe)
        undermappe_length = len(Undermappe)
        aktID_length = len(str(AktID))
        dokID_length = len(str(DokumentID))

        fixed_length = len(base_path) + overmappe_length + undermappe_length + aktID_length + dokID_length + 7
        available_title_length = max_path_length - fixed_length

        if len(Titel) > available_title_length:
            return Titel[:available_title_length]
        
        return Titel
    
    def fetch_document_info(DokumentID, session, AktID, Titel):
        url = f"https://ad.go.aarhuskommune.dk/_goapi/Documents/Data/{DokumentID}"
        response = session.get(url)
        DocumentData = response.text
        data = json.loads(DocumentData)
        item_properties = data.get("ItemProperties", "")
        file_type_match = re.search(r'ows_File_x0020_Type="([^"]+)"', item_properties)
        version_ui_match = re.search(r'ows__UIVersionString="([^"]+)"', item_properties)
        DokumentType = file_type_match.group(1) if file_type_match else "Not found"
        VersionUI = version_ui_match.group(1) if version_ui_match else "Not found"
        Feedback = " "
        file_path = os.path.join(
            "C:\\Users",
            os.getenv("USERNAME"),
            "Downloads",
            f"{AktID:04} - {DokumentID} - {Titel}"
        )
        return {"DokumentType": DokumentType, "VersionUI": VersionUI, "Feedback": Feedback, "file_path": file_path}
    
    with requests.Session() as session:
            session.auth = HttpNtlmAuth(GoUsername, GoPassword)
            session.headers.update({"Content-Type": "application/json"}) 

    dt_DocumentList['Dokumentdato'] = pd.to_datetime(dt_DocumentList['Dokumentdato'], errors='coerce',format='%d-%m-%Y')

    
    for index, row in dt_DocumentList.iterrows():
        # Convert items to strings unless they are explicitly integers
        Omfattet = str(row["Omfattet af ansøgningen? (Ja/Nej)"])
        DokumentID = str(row["Dok ID"])
        
        # Handle AktID conversion
        AktID = row['Akt ID']
        if isinstance(AktID, str):  
            AktID = int(AktID.replace('.', ''))
        elif isinstance(AktID, int):  
            AktID = AktID

        Titel = str(row["Dokumenttitel"])
        mimetypes.add_type("application/x-msmetafile", ".emz")
        # Split title into name and extension
        parts = Titel.rsplit('.', 1)  # Splits at the last dot
        if len(parts) == 2:
            name, ext = parts
            # Check if it's a known file extension
            if mimetypes.guess_type(f"file.{ext}")[0]:  
                Titel = name  # Remove extension
                print(f"Updated Titel (without filetype): {Titel}")
            else:
                print("No known filetype detected.")
        else:
            print("No file extension detected.")




        BilagTilDok = str(row["Bilag til Dok ID"])
        DokBilag = str(row["Bilag"])
        Dokumentkategori = str(row["Dokumentkategori"])
        Aktstatus = str(row["Gives der aktindsigt i dokumentet? (Ja/Nej/Delvis)"])
        Begrundelse = str(row["Begrundelse hvis nej eller delvis"])
        Dokumentdato =row['Dokumentdato']
        if isinstance(Dokumentdato, pd.Timestamp):
            Dokumentdato = Dokumentdato.strftime("%d-%m-%Y")
        else:
            Dokumentdato = datetime.strptime(Dokumentdato, "%Y-%m-%d").strftime("%d-%m-%Y")
        # Dokumentdato = datetime.strptime(row["Dokumentdato"], "%d-%m-%Y").strftime("%d-%m-%Y")
    
        print(f"AktID til debug: {AktID}")

        # Declare the necessary variables
        base_path = "Teams/tea-teamsite10506/Delte dokumenter/Aktindsigter/"

        # Sanitize the title
        Titel = sanitize_title(Titel)

        Titel = calculate_available_title_length(base_path, Overmappe, Undermappe, AktID, DokumentID, Titel)


        #Henter Dokumentinfo: 

        if GeoSag == True: 
            Metadata = fetch_document_info(DokumentID, session, AktID, Titel)
            
            # Extracting variables for further use in the loop
            DokumentType = Metadata["DokumentType"]
            VersionUI = Metadata["VersionUI"]
            Feedback = Metadata["Feedback"]
            file_path = Metadata["file_path"]

        else: #Det er en NovaSag
            TransactionID = str(uuid.uuid4())
            url = f"{KMDNovaURL}/Document/GetList?api-version=2.0-Case"

            headers = {
                "Authorization": f"Bearer {KMD_access_token}",
                "Content-Type": "application/json"
            }

            payload = {
                "common": {
                    "transactionId": TransactionID,
                    #"uuid": DokumentID ## skal hente dokument id og ikke dokumentnr, find ud af hvornår den skal hentes.
                },
                "paging": {
                    "startRow": 1,
                    "numberOfRows": 100
                },
                "documentNumber": DokumentID,
                "caseNumber": Sagsnummer,
                "getOutput": {
                    "documentDate": True,
                    "title": True,
                    "fileExtension": True
                    }
                }

            try:
                response = requests.put(url, headers=headers, json=payload)
                if response.status_code == 200:
                    print(response.status_code)
                else:
                    print("Failed to fetch Sagstitel from NOVA. Status Code:", response.status_code)
            except Exception as e:
                raise Exception("Failed to fetch Sagstitel (Nova):", str(e))

            DokumentType = response.json()["documents"][0]["fileExtension"]
            DocumentUuid = response.json()["documents"][0]["documentUuid"]

        

        Titel = f"{AktID:04} - {DokumentID} - {Titel}.{DokumentType}"

        # Parse and prepare data for the row
        row_to_add = {
            "Akt ID": int(AktID),
            "Filnavn": Titel,
            "Dokumentkategori": Dokumentkategori,
            "Dokumentdato": datetime.strptime(Dokumentdato, "%d-%m-%Y"),
            "Dok ID": DokumentID,
            "Bilag til Dok ID": BilagTilDok,
            "Bilag": DokBilag,
            "Omfattet af aktindsigt?": Omfattet,
            "Gives der aktindsigt?": Aktstatus,
            "Begrundelse hvis Nej/Delvis": Begrundelse,
        }

        # Append the row to the DataFrame
        dt_AktIndex = pd.concat([dt_AktIndex, pd.DataFrame([row_to_add])], ignore_index=True)

        # Sort and reset index
        dt_AktIndex = dt_AktIndex.sort_values(by="Akt ID", ascending=True).reset_index(drop=True)
        
    return {
    "out_dt_AktIndex": dt_AktIndex,
    }
