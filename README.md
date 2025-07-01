
# 📄 README

## Aktbob Delivery Robot

**Aktbob Delivery** is an automation for **Teknik og Miljø, Aarhus Kommune**. It retrieves documents from Filarkiv and SharePoint, generates indexes (aktlister), uploads files to delivery folders, and notifies stakeholders when access-to-records deliveries are ready.

---

## 🚀 Features

✅ **Document Retrieval**
- Fetches document lists from SharePoint and Filarkiv
- Validates presence of documents before processing

📄 **Aktliste Generation**
- Creates Excel and PDF indexes summarizing the delivered documents
- Auto-formats headers, tables, and date fields

📤 **Automated Uploads**
- Uploads all prepared files to SharePoint delivery folders
- Handles large files via chunked uploads

🔗 **Link Sharing**
- Generates public and secure SharePoint links to delivery folders
- Updates Deskpro tickets and Podio records with generated links

📧 **Notifications**
- Sends emails to caseworkers if document lists are empty or errors occur
- Notifies recipients when deliveries are complete

🔐 **Credential Management**
- Fetches and refreshes API tokens (KMD and Filarkiv)
- All credentials are stored securely in OpenOrchestrator

---

## 🧭 Process Flow

1. **Token Management**
   - Fetches or refreshes KMD and Filarkiv access tokens (`GetKMDAcessToken.py`, `GetFilarkivToken.py`)
2. **Document List Retrieval**
   - Downloads metadata and file lists for the selected case (`GetDocumentList.py`)
   - Stops processing if no documents are found
3. **Document Preparation**
   - Downloads and filters documents from Filarkiv (`DownloadFilesFromFilarkivAndUploadToSharePoint.py`)
   - Generates the Excel Aktliste (`GenerateAndUploadAktliste.py`)
4. **Upload to SharePoint**
   - Uploads all processed documents and indexes to the target delivery folder
5. **Share Link Generation**
   - Creates public and password-protected SharePoint links (`SendShareLinkToDeskpro.py`)
   - Updates Deskpro tickets and Podio records
6. **Notifications**
   - Sends confirmation emails with links to stakeholders
   - Logs all actions in OpenOrchestrator

---

## 🔐 Privacy & Security

- All APIs use HTTPS
- Credentials and tokens are managed securely in OpenOrchestrator
- No personal data is stored locally after processing
- Temporary files are removed after uploads

---

## ⚙️ Dependencies

- Python 3.10+
- `requests`
- `requests-ntlm`
- `pandas`
- `pyodbc`
- `openpyxl`
- `reportlab`
- `office365-rest-python-client`
- `msal`
- `smtplib`

---

## 👷 Maintainer

Gustav Chatterton  
*Digital udvikling, Teknik og Miljø, Aarhus Kommune*
