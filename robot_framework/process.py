from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
from OpenOrchestrator.database.queues import QueueElement
import requests
from requests_ntlm import HttpNtlmAuth
import json
from GetKMDAcessToken import GetKMDToken
from GetFilarkivToken import GetFilarkivToken
from GetDocumentList import get_document_list
from GetDocumentsForAktliste import GetDocumentsForAktliste
from DownloadFilesFromFilarkivAndUploadToSharePoint import DownloadFilesFromFilarkivAndUploadToSharePoint
from GenerateAndUploadAktliste import GenerateAndUploadAktliste
from SendShareLinkToDeskpro import SendShareLinkToDeskpro
from SendSMTPMail import send_email # Import the function and dataclass
import os
import sys

# pylint: disable-next=unused-argument
def process(orchestrator_connection: OrchestratorConnection, queue_element: QueueElement | None = None) -> None:

    """Do the primary process of the robot."""
    certification = orchestrator_connection.get_credential("SharePointCert")
    api = orchestrator_connection.get_credential("SharePointAPI")

    GraphAppIDAndTenant = orchestrator_connection.get_credential("GraphAppIDAndTenant")
    SharePointAppID = GraphAppIDAndTenant.username
    SharePointTenant = GraphAppIDAndTenant.password
    SharePointURL = orchestrator_connection.get_constant("AktbobSharePointURL").value
    CloudConvert = orchestrator_connection.get_credential("CloudConvertAPI")
    CloudConvertAPI = CloudConvert.password
    UdviklerMailAktbob =  orchestrator_connection.get_constant("UdviklerMailAktbob").value
    RobotCredentials = orchestrator_connection.get_credential("RobotCredentials")
    RobotUserName = RobotCredentials.username
    RobotPassword = RobotCredentials.password
    GOAPILIVECRED = orchestrator_connection.get_credential("GOAktApiUser")
    GoUsername = GOAPILIVECRED.username
    GoPassword = GOAPILIVECRED.password
    KMDNovaURL = orchestrator_connection.get_constant("KMDNovaURL").value
    FilarkivURL = orchestrator_connection.get_constant("FilarkivURL").value
    AktbobAPI = orchestrator_connection.get_credential("AktbobAPIKey")
    AktbobAPIKey = AktbobAPI.password
    tenant = api.username
    client_id = api.password
    thumbprint = certification.username
    cert_path = certification.password
    # ---- Henter access tokens ----
    KMD_access_token = GetKMDToken(orchestrator_connection)
    Filarkiv_access_token = GetFilarkivToken(orchestrator_connection)

    # ---- Deffinerer Go-session ----
    def GO_Session(GoUsername, GoPassword):
        session = requests.Session()
        session.auth = HttpNtlmAuth(GoUsername, GoPassword)
        session.headers.update({
            "Content-Type": "application/json"
        })
        return session

    # ---- Initialize Go-Session ----
    go_session = GO_Session(GoUsername, GoPassword)

    queue = json.loads(queue_element.data)
    Sagsnummer = queue.get("Sagsnummer")
    MailModtager = queue.get("MailModtager")
    DeskProID = queue.get("DeskProID")
    DeskProTitel = queue.get("DeskProTitel")
    PodioID = queue.get("PodioID")
    Overmappe = queue.get("Overmappe")
    Undermappe = queue.get("Undermappe")
    GeoSag = queue.get("GeoSag")
    NovaSag = queue.get("NovaSag")
    FilarkivCaseID = queue.get("FilarkivCaseID")

    # ---- Run "GetDocumentList" ----
    Sagstitel, dt_DocumentList,  DokumentlisteDatoString = get_document_list(Sagsnummer, GeoSag, NovaSag, KMD_access_token, KMDNovaURL, SharePointURL, Overmappe, Undermappe, MailModtager, tenant, client_id, thumbprint, cert_path, go_session)
    
    if any(x is None for x in [Sagstitel, dt_DocumentList, DokumentlisteDatoString]):
        orchestrator_connection.log_info('None returned from doclist. Maybe file transfer is tried before doclist is made.')
        return

    if dt_DocumentList.empty:
        sender = "aktbob@aarhus.dk" 
        subject = f"{Sagsnummer} er en tom sag"
        body = f"""Sagen: {Sagsnummer} er en tom sag. Vær opmærksom på, at processen ikke kan behandle tomme sager.<br><br>
        Det anbefales at følge <a href="https://aarhuskommune.atlassian.net/wiki/spaces/AB/pages/64979049/AKTBOB+--+Vejledning">vejledningen</a>, 
        hvor du også finder svar på de fleste spørgsmål og fejltyper.
        """
        smtp_server = "smtp.adm.aarhuskommune.dk"   
        smtp_port = 25               

        send_email(
            receiver=MailModtager,
            sender=sender,
            subject=subject,
            body=body,
            smtp_server=smtp_server,
            smtp_port=smtp_port,
            html_body=True
        )
        orchestrator_connection.log_info('Dokumentlisten er tom. Processen afsluttes')
        return

    dt_AktIndex = GetDocumentsForAktliste(dt_DocumentList = dt_DocumentList, Overmappe = Overmappe, Undermappe = Undermappe, Sagsnummer = Sagsnummer, GeoSag = GeoSag, KMDNovaURL = KMDNovaURL, KMD_access_token = KMD_access_token, go_session = go_session)
    dt_AktIndex = DownloadFilesFromFilarkivAndUploadToSharePoint(FilarkivURL, Filarkiv_access_token, dt_AktIndex, FilarkivCaseID, SharePointURL, Overmappe, Undermappe, MailModtager, Sagsnummer, tenant, client_id, thumbprint, cert_path, orchestrator_connection )
    GenerateAndUploadAktliste(dt_AktIndex, Sagsnummer, DokumentlisteDatoString,SharePointURL, Overmappe, Undermappe, tenant, client_id, thumbprint, cert_path )
    SendShareLinkToDeskpro(SharePointURL, Overmappe, PodioID, AktbobAPIKey, DeskProID, MailModtager, Sagsnummer, tenant, client_id, thumbprint, cert_path, orchestrator_connection)
    