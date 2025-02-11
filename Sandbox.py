from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
#Packages
import os
import requests
from requests_ntlm import HttpNtlmAuth
from email.message import EmailMessage

import smtplib
from io import BytesIO

#Functions or scripts
from GetKMDAcessToken import GetKMDToken
from GetFilarkivToken import GetFilarkivToken
import GetDocumentList
import GetDocumentsForAktliste
import GenerateAndUploadAktliste
import DownloadFilesFromFilarkivAndUploadToSharePoint
import SendShareLinkToDeskpro

#   ---- Henter Assets ----
orchestrator_connection = OrchestratorConnection("Henter Assets", os.getenv('OpenOrchestratorSQL'),os.getenv('OpenOrchestratorKey'), None)
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


#---- Henter kø-elementer ----
#GO
Sagsnummer = "GEO-2024-043144"
MailModtager = "Gujc@aarhus.dk"
DeskProID = "2088"
DeskProTitel = "Aktindsigt i aktindsigter"
PodioID = "2988315804"
Overmappe = "2088 - Aktindsigt i aktindsigter"
Undermappe = "GEO-2024-043144 - GustavTestAktIndsigt2"
FilarkivCaseID = "dc35281b-4319-45b9-b32f-349d5d1834b7"
GeoSag = True
NovaSag = False

# # #Nova
# Sagsnummer = "S2021-456011"
# MailModtager = "Gujc@aarhus.dk"
# DeskProID = "2088"
# DeskProTitel = "Aktindsigt i aktindsigter"
# PodioID = "2988315804"
# Overmappe = "2088 - Aktindsigt i aktindsigter"
# Undermappe = "S2021-456011 - TEST - Ejendom uden ejendomsnr"
# FilarkivCaseID = "a6fd808f-c7fd-4149-aca4-c35113706b5e"
# GeoSag = False
# NovaSag = True


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
    "KMDNovaURL": KMDNovaURL
}


# ---- Run "GetDocumentList" ----
GetDocumentList_Output_arguments = GetDocumentList.invoke(Arguments, go_session)
Sagstitel = GetDocumentList_Output_arguments.get("sagstitel")
dt_DocumentList = GetDocumentList_Output_arguments.get("dt_DocumentList")
DokumentlisteDatoString = GetDocumentList_Output_arguments.get("out_DokumentlisteDatoString")



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
    "in_GoPassword":  GoPassword
}

GetDocumentsForAktliste_Output_arguments = GetDocumentsForAktliste.invoke_GetDocumentsForAktliste(Arguments_GetDocumentsForAktliste)
dt_AktIndex = GetDocumentsForAktliste_Output_arguments.get("out_dt_AktIndex")


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
"in_GoPassword": GoPassword
}

GenerateAndUploadAktliste_Output_arguments = GenerateAndUploadAktliste.invoke_GenerateAndUploadAktliste(Arguments_GenerateAndUploadAktliste)
Test = GenerateAndUploadAktliste_Output_arguments.get("out_Text")
print(Test)

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
"in_RobotPassword": RobotPassword
}

DownloadFilesFromFilarkivAndUploadToSharePoint_Output_arguments = DownloadFilesFromFilarkivAndUploadToSharePoint.invoke_DownloadFilesFromFilarkivAndUploadToSharePoint(Arguments_DownloadFilesFromFilarkivAndUploadToSharePoint)
Test = DownloadFilesFromFilarkivAndUploadToSharePoint_Output_arguments.get("out_Text")
print(Test)


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
"in_MailModtager": MailModtager,
"in_Sagsnummer": Sagsnummer
}


SendShareLinkToDeskpro_Output_arguments = SendShareLinkToDeskpro.invoke_SendShareLinkToDeskpro(Arguments_SendShareLinkToDeskpro)
Test = SendShareLinkToDeskpro_Output_arguments.get("out_Text")
print(Test)