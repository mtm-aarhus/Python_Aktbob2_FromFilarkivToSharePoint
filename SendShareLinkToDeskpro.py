from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
import requests
import json
from SendSMTPMail import send_email
from SharePointUploader import sharepoint_client, get_sharepoint_folder_links
def upload_sharepoint_link_to_podio(PodioID: str, ApiURL: str, ApiKey: str, SharePointLink: str):
    API_start = ApiURL.rsplit('/', 1)[0]

    url = f"{API_start}/Api/Podio/{PodioID}/SharepointmappeField"
    headers = {
        "ApiKey": ApiKey,
        "Content-Type": "application/json"
    }
    json_body = {"value": SharePointLink}

    response = requests.put(url, json=json_body, headers=headers)
    response.raise_for_status()

def send_LinkToDeskpro(secure_link, password, deskpro_id):
    # Define the URL
    url = "https://aarhuskommune4.deskpro.com/api/v2/webhooks/A7O1H3HKEW76MAXA/invocation"
    
    # Calculate expiration date (current date + 30 days) and format it
    # expiration_date = (datetime.now(timezone.utc) + timedelta(days=60)).strftime("%Y-%m-%d")
    # JSON payload
    payload = {
        "sharePointShareUrl": secure_link,
        "Password": password,
        #"sharePointExpirationDate": expiration_date,
        "deskproTicketId": deskpro_id
    }
    
    # Headers
    headers = {
        "Content-Type": "application/json"
    }
    
    # Make the request
    response = requests.post(url, headers=headers, data=json.dumps(payload))
    response.raise_for_status()   
    
def SendShareLinkToDeskpro(SharePointURL, Overmappe, PodioID, AktbobAPIKey, DeskProID, MailModtager, Sagsnummer, tenant, client_id, thumbprint, cert_path, orchestrator_connection: OrchestratorConnection):

    DeskProAPI = orchestrator_connection.get_credential("DeskProAPI") 
    DeskProAPIKey = DeskProAPI.password  

    Deskprourl = f"https://mtmsager.aarhuskommune.dk/api/v2/tickets/{DeskProID}"

    headers = {
        'Authorization': DeskProAPIKey,
        'Cookie': 'dp_last_lang=da'
    }

    response = requests.get(Deskprourl, headers=headers)
    response.raise_for_status()
    data = response.json()
    fields = data.get("data", {}).get("fields", {})
    # Get values from field 110 and 134
    sharepoint_link = fields.get("110", {}).get("value", "")

    SendEmailUdleveringsmappe = False

    # If no SharePoint link exists → we must generate one
    if not isinstance(sharepoint_link, str) or not sharepoint_link.strip():
        SendEmailUdleveringsmappe = True

    # Fetch SharePoint folder link
    site_relative_path = "/Teams/tea-teamsite10506/Delte Dokumenter"
    client = sharepoint_client(tenant, client_id, thumbprint, cert_path, SharePointURL)

    # Retrieve both links
    public_link, secure_link, password = get_sharepoint_folder_links(client, Overmappe, site_relative_path)
    API_url  = orchestrator_connection.get_credential('AktbobAPIKey').username

    upload_sharepoint_link_to_podio(PodioID, API_url, AktbobAPIKey, public_link)
    
    send_LinkToDeskpro(secure_link, password, DeskProID)
    
    if SendEmailUdleveringsmappe:
        # # Define email details
        sender = "aktbob@aarhus.dk" # Replace with actual sender
        subject = f"{Sagsnummer}: Udleveringsmappe klar"

        body = f"""
            
            Sag: <a href="https://mtmsager.aarhuskommune.dk/app#/t/ticket/{DeskProID}">{Overmappe}</a><br><br>
            Du kan se udleveringsmappen her: <a href="{public_link}">SharePoint</a>.<br><br>
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
