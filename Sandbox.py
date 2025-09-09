from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
#Packages
import os
import requests
from requests_ntlm import HttpNtlmAuth

#Functions or scripts
from GetKMDAcessToken import GetKMDToken
from GetFilarkivToken import GetFilarkivToken
import GetDocumentList
import GetDocumentsForAktliste
import GenerateAndUploadAktliste
import DownloadFilesFromFilarkivAndUploadToSharePoint
import SendShareLinkToDeskpro
import json
from SendSMTPMail import send_email # Import the function and dataclass


#   ---- Henter Assets ----
orchestrator_connection = OrchestratorConnection("Henter Assets", os.getenv('OpenOrchestratorSQL'),os.getenv('OpenOrchestratorKey'), None)

# pylint: disable-next=unused-argument

"""Do the primary process of the robot."""
orchestrator_connection.log_trace("Running process.")

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

queue_element = """insert queue element here as json string"""
    #---- Henter kø-elementer ----
queue = json.loads(queue_element)
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


# ---- Run "GetDokumentlist" ----sort
Arguments = {
    "in_RobotUserName": RobotUserName,
    "in_RobotPassword": RobotPassword,
    "in_Sagsnummer": Sagsnummer,
    "in_SharePointUrl": SharePointURL,
    "in_Overmappe": Overmappe,
    "in_Undermappe": Undermappe,
    "in_GeoSag": GeoSag, 
    "in_NovaSag": NovaSag,
    "GoUsername": GoUsername,
    "GoPassword":  GoPassword,
    "KMD_access_token": KMD_access_token,
    "KMDNovaURL": KMDNovaURL,
    "in_MailModtager": MailModtager,
    "tenant": tenant,
    "client_id": client_id,
    "thumbprint": thumbprint,
    "cert_path": cert_path
}


# ---- Run "GetDocumentList" ----
GetDocumentList_Output_arguments = GetDocumentList.invoke(Arguments, go_session)
Sagstitel = GetDocumentList_Output_arguments.get("sagstitel")
dt_DocumentList = GetDocumentList_Output_arguments.get("dt_DocumentList")
DokumentlisteDatoString = GetDocumentList_Output_arguments.get("out_DokumentlisteDatoString")

if dt_DocumentList.empty:
    print("Number of rows:",len(dt_DocumentList))
        ###---- Send mail til sagsansvarlig ----####

    # Define email details
    sender = "aktbob@aarhus.dk" 
    subject = f"{Sagsnummer} er en tom sag"
    body = f"""Sagen: {Sagsnummer} er en tom sag. Vær opmærksom på, at processen ikke kan behandle tomme sager.<br><br>
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
    raise ValueError("Dokumentlisten inderholder ikke nogen data - Processen fejler")
else:
    print("Number of rows:",len(dt_DocumentList))






# ---- Run "GetDocumentsForAktliste" ----
Arguments_GetDocumentsForAktliste = {
    "in_dt_Documentlist": dt_DocumentList,
    "in_CloudConvertAPI": CloudConvertAPI,
    "in_MailModtager": MailModtager,
    "in_UdviklerMail": UdviklerMailAktbob,
    "in_RobotUserName": RobotUserName,
    "in_RobotPassword": RobotPassword,
    "in_SharePointAppID": SharePointAppID,
    "in_SharePointTenant":SharePointTenant,
    "in_SharePointUrl": SharePointURL,
    "in_Overmappe": Overmappe,
    "in_Undermappe": Undermappe,
    "in_Sagsnummer": Sagsnummer,
    "in_GeoSag": GeoSag,
    "in_NovaSag": NovaSag,
    "in_NovaToken": KMD_access_token,
    "in_KMDNovaURL": KMDNovaURL,
    "in_GoUsername": GoUsername,
    "in_GoPassword":  GoPassword,
    "tenant": tenant,
    "client_id": client_id,
    "thumbprint": thumbprint,
    "cert_path": cert_path
}

GetDocumentsForAktliste_Output_arguments = GetDocumentsForAktliste.invoke_GetDocumentsForAktliste(Arguments_GetDocumentsForAktliste)
dt_AktIndex = GetDocumentsForAktliste_Output_arguments.get("out_dt_AktIndex")



# ---- run DownloadFilesFromFilarkivAndUploadToSharePoint ----

Arguments_DownloadFilesFromFilarkivAndUploadToSharePoint = {
"in_dt_AktIndex": dt_AktIndex,
"in_FilarkivURL": FilarkivURL,
"in_Filarkiv_access_token": Filarkiv_access_token,
"in_FilarkivCaseID": FilarkivCaseID,
"in_SharePointAppID": SharePointAppID,
"in_SharePointTenant": SharePointTenant,
"in_SharePointURL": SharePointURL,
"in_Overmappe": Overmappe,
"in_Undermappe": Undermappe,
"in_RobotUserName": RobotUserName,
"in_RobotPassword": RobotPassword,
"in_MailModtager": MailModtager,
"in_Sagsnummer": Sagsnummer,
"tenant": tenant,
"client_id": client_id,
"thumbprint": thumbprint,
"cert_path": cert_path
}

DownloadFilesFromFilarkivAndUploadToSharePoint_Output_arguments = DownloadFilesFromFilarkivAndUploadToSharePoint.invoke_DownloadFilesFromFilarkivAndUploadToSharePoint(Arguments_DownloadFilesFromFilarkivAndUploadToSharePoint)
Test = DownloadFilesFromFilarkivAndUploadToSharePoint_Output_arguments.get("out_Text")
dt_AktIndex = DownloadFilesFromFilarkivAndUploadToSharePoint_Output_arguments.get("out_dt_AktIndex")
orchestrator_connection.log_trace(Test) 
# ---- Run "Generate&UploadAktlistPDF" ----

Arguments_GenerateAndUploadAktliste = {
"in_dt_AktIndex": dt_AktIndex,
"in_Sagsnummer": Sagsnummer,
"in_DokumentlisteDatoString":DokumentlisteDatoString, 
"in_RobotUserName": RobotUserName,
"in_RobotPassword": RobotPassword,
"in_SagsTitel": Sagstitel,
"in_SharePointAppID": SharePointAppID,
"in_SharePointTenant": SharePointTenant,
"in_Overmappe": Overmappe,
"in_Undermappe": Undermappe,
"in_SharePointURL": SharePointURL,
"in_GoUsername":GoUsername,
"in_GoPassword": GoPassword,
"tenant": tenant,
"client_id": client_id,
"thumbprint": thumbprint,
"cert_path": cert_path
}

GenerateAndUploadAktliste_Output_arguments = GenerateAndUploadAktliste.invoke_GenerateAndUploadAktliste(Arguments_GenerateAndUploadAktliste)
Test = GenerateAndUploadAktliste_Output_arguments.get("out_Text")
orchestrator_connection.log_trace(Test)

##---- run SendShareLinkToDeskpro ----##

Arguments_SendShareLinkToDeskpro = {

"in_SharePointAppID": SharePointAppID,
"in_SharePointTenant": SharePointTenant,
"in_SharePointURL": SharePointURL,
"in_Overmappe": Overmappe,
"in_Undermappe": Undermappe,
"in_RobotUserName": RobotUserName,
"in_RobotPassword": RobotPassword,
"in_PodioID": PodioID,
"in_AktbobAPIKey": AktbobAPIKey,
"in_DeskProID": DeskProID,
"in_DeskProTitel": DeskProTitel,
"in_MailModtager": MailModtager,
"in_Sagsnummer": Sagsnummer,
"tenant": tenant,
"client_id": client_id,
"thumbprint": thumbprint,
"cert_path": cert_path
}
    

SendShareLinkToDeskpro_Output_arguments = SendShareLinkToDeskpro.invoke_SendShareLinkToDeskpro(Arguments_SendShareLinkToDeskpro,orchestrator_connection)
Test = SendShareLinkToDeskpro_Output_arguments.get("out_Text")
orchestrator_connection.log_trace(Test)
