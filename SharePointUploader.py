import os
from office365.sharepoint.client_context import ClientContext
from datetime import datetime, timedelta, timezone
import string
import random
from office365.sharepoint.sharing.links.kind import SharingLinkKind
from office365.sharepoint.sharing.role import Role

def sharepoint_client(tenant, client_id, thumbprint, cert_path, SharePointUrl) -> ClientContext:
        try:
            cert_credentials = {
                "tenant": tenant,
                "client_id": client_id,
                "thumbprint": thumbprint,
                "cert_path": cert_path
            }
            ctx = ClientContext(SharePointUrl).with_client_certificate(**cert_credentials)

            # Load the SharePoint web to test the connection
            web = ctx.web
            ctx.load(web)
            ctx.execute_query()

            return ctx
        except Exception as e:
            raise Exception(f"Authentication failed: {e}")
    # File downloading logic from SharePoint
def download_file_from_sharepoint(client: ClientContext, sharepoint_file_url: str, download_path) -> str:
    """
    Downloads a file from SharePoint and returns the local file path.
    """
    file_name = sharepoint_file_url.split("/")[-1]  # Extract file name from URL
    local_file_path = os.path.join(download_path, file_name)  # Define local path

    try:
        # Ensure the download directory exists
        if not os.path.exists(download_path):
            os.makedirs(download_path)

        # Download the file
        with open(local_file_path, "wb") as local_file:
            client.web.get_file_by_server_relative_path(sharepoint_file_url).download(local_file).execute_query()

        return local_file_path
    except Exception as e:
        raise Exception(f"Error downloading file from SharePoint: {e}")
        
def upload_file_to_sharepoint(site_url,Overmappe, Undermappe, file_path, ctx):
    try:

        # Construct folder path
        folder_url = f"/Teams/tea-teamsite10506/Delte Dokumenter/Aktindsigter/{Overmappe}/{Undermappe}"
        folder = ctx.web.get_folder_by_server_relative_path(folder_url)
        ctx.load(folder)
        ctx.execute_query()

        file_name = os.path.basename(file_path)
        file_size = os.path.getsize(file_path)

        # Attempt normal upload for files ≤ 4MB
        if file_size <= 4 * 1024*1024:  
            try:
                with open(file_path, "rb") as f:
                    file = folder.files.add(file_name, f.read(), overwrite=True).execute_query()
                print(f"File uploaded successfully using normal upload.")
                return  # 
            except Exception as e:
                print(f"Normal upload failed for '{file_name}', switching to chunked upload. Error: {e}")

        size_chunk = 1000000  # 1MB chunks
        def print_upload_progress(offset):
            print(f"Uploaded {offset} bytes of {file_size} ({round(offset / file_size * 100, 2)}%)")

        with open(file_path, "rb") as f:
            uploaded_file = folder.files.create_upload_session(f, size_chunk, print_upload_progress).execute_query()

        print(f"File uploaded successfully using chunked upload.")

    except Exception as e:
        raise Exception(f"Failed to upload file '{file_path}': {e}")
    
def get_sharepoint_folder_links(client: ClientContext, Overmappe: str, site_relative_path):
    """ Generates both public and password-protected SharePoint folder links. """

    folder_url = f"{site_relative_path}/Aktindsigter/{Overmappe}"
    folder = client.web.get_folder_by_server_relative_path(folder_url)
    client.load(folder)
    client.execute_query()

    # Generate a public anonymous view-only link
    public_link_result = folder.share_link(SharingLinkKind.AnonymousView).execute_query()
    public_link = public_link_result.value.sharingLinkInfo.Url

    # Generate a password-protected link
    expiration_date = (datetime.now(timezone.utc) + timedelta(days=60)).strftime("%Y-%m-%dT%H:%M:%S%z")
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

