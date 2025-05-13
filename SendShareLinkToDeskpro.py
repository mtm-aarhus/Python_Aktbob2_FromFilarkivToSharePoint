def invoke_SendShareLinkToDeskpro(Arguments_SendShareLinkToDeskpro):
    from office365.sharepoint.client_context import ClientContext
    from office365.sharepoint.sharing.role import Role
    from office365.sharepoint.sharing.user_role_assignment import UserRoleAssignment
    from office365.runtime.auth.user_credential import UserCredential
    from office365.sharepoint.sharing.links.kind import SharingLinkKind
    import uuid
    import requests
    import json
    from datetime import datetime, timedelta, timezone
    from SendSMTPMail import send_email
    import random
    import string


    SharePointAppID = Arguments_SendShareLinkToDeskpro.get("in_SharePointAppID")
    SharePointTenant = Arguments_SendShareLinkToDeskpro.get("in_SharePointTenant")
    SharePointURL = Arguments_SendShareLinkToDeskpro.get("in_SharePointURL")
    Overmappe = Arguments_SendShareLinkToDeskpro.get("in_Overmappe")
    Undermappe = Arguments_SendShareLinkToDeskpro.get("in_Undermappe")
    RobotUserName = Arguments_SendShareLinkToDeskpro.get("in_RobotUserName")
    RobotPassword = Arguments_SendShareLinkToDeskpro.get("in_RobotPassword")
    PodioID = Arguments_SendShareLinkToDeskpro.get("in_PodioID")
    AktbobAPIKey = Arguments_SendShareLinkToDeskpro.get("in_AktbobAPIKey")
    DeskProID = Arguments_SendShareLinkToDeskpro.get("in_DeskProID")
    MailModtager = Arguments_SendShareLinkToDeskpro.get("in_MailModtager")
    Sagsnummer = Arguments_SendShareLinkToDeskpro.get("in_Sagsnummer")


    def sharepoint_client(RobotUserName, RobotPassword, SharePointUrl) -> ClientContext:
        try:
            credentials = UserCredential(RobotUserName, RobotPassword)
            ctx = ClientContext(SharePointUrl).with_credentials(credentials)

            # Load the SharePoint web to test the connection
            web = ctx.web
            ctx.load(web)
            ctx.execute_query()

            return ctx
        except Exception as e:
            print(f"Authentication failed: {e}")
            raise

    def get_sharepoint_folder_links(client: ClientContext, Overmappe: str, site_relative_path):
        """ Generates both public and password-protected SharePoint folder links. """
        try:
            folder_url = f"{site_relative_path}/Aktindsigter/{Overmappe}"
            folder = client.web.get_folder_by_server_relative_path(folder_url)
            client.load(folder)
            client.execute_query()

            # Generate a public anonymous view-only link
            public_link_result = folder.share_link(SharingLinkKind.AnonymousView).execute_query()
            public_link = public_link_result.value.sharingLinkInfo.Url

            # Generate a password-protected link
            #expiration_date = (datetime.utcnow() + timedelta(days=60)).isoformat() + "Z"
            expiration_date = (datetime.now(timezone.utc) + timedelta(days=30)).strftime("%Y-%m-%dT%H:%M:%S%z")
            characters = string.ascii_letters + string.digits 
            password = ''.join(random.choices(characters, k=6))
            secure_link_result = folder.share_link(
                link_kind=SharingLinkKind.Flexible,
                expiration=expiration_date,
                password=password,
                role=Role.View
            ).execute_query()
            secure_link = secure_link_result.value.sharingLinkInfo.Url

            return public_link, secure_link, password
        except Exception as e:
            raise Exception(f"Error generating shareable links: {e}")


    def upload_sharepoint_link_to_podio(PodioID: str, ApiKey: str, SharePointLink: str):
        try:
            url = f"https://aktbob-external-api.grayglacier-2d22de15.northeurope.azurecontainerapps.io/Api/Podio/{PodioID}/SharepointmappeField"
            headers = {
                "ApiKey": ApiKey,
                "Content-Type": "application/json"
            }
            json_body = {"value": SharePointLink}

            response = requests.put(url, json=json_body, headers=headers)
            
            print("Response Status:", response.status_code)
            print("Response:", response.text)
        except requests.exceptions.RequestException as e:
            print("Failed")
            #raise Exception(f"Request to API failed: {e}")    
    
    def send_LinkToDeskpro(secure_link, password, deskpro_id):
        try:
            # Define the URL
            url = "https://aarhuskommune4.deskpro.com/api/v2/webhooks/A7O1H3HKEW76MAXA/invocation"
            
            # Calculate expiration date (current date + 30 days) and format it
            expiration_date = (datetime.now(timezone.utc) + timedelta(days=60)).strftime("%Y-%m-%d")
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
            
            print("Response Status:", response.status_code)
        except requests.exceptions.RequestException as e:
            raise Exception(f"Request to API failed: {e} with status: {response.status_code}")    
    
    
    # Fetch SharePoint folder link
    site_relative_path = "/Teams/tea-teamsite10506/Delte Dokumenter"
    client = sharepoint_client(RobotUserName, RobotPassword, SharePointURL)


    # Retrieve both links
    public_link, secure_link, password = get_sharepoint_folder_links(client, Overmappe, site_relative_path)
    print(f"Public Shareable Link: {public_link}")
    print(f"Password-Protected Shareable Link: {secure_link}")


    upload_sharepoint_link_to_podio(PodioID, AktbobAPIKey, public_link)
    
    send_LinkToDeskpro(secure_link, password, DeskProID)


    # # Define email details
    sender = "Aktbob<rpamtm001@aarhus.dk>" # Replace with actual sender
    subject = f"{Sagsnummer}: Udleveringsmappe klar"

    body = f"""
        Sag: {Overmappe} <br><br>
        Du kan se udleveringsmappen her: <a href="{public_link}">SharePoint</a>.<br><br>
        Det anbefales at følge <a href="https://aarhuskommune.sharepoint.com/:w:/t/tea-teamsite10506/EVjuZhmtsHRGi6H7-COs26AB6afOXvReKSnWJ1XK1mKxZw?e=n03h0t/">vejledningen</a>, 
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


    return {"out_Text": "Delinger er blevet oprettet"}



