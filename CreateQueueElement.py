import os
from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection


#   ---- Henter Assets ----
orchestrator_connection = OrchestratorConnection("Henter Assets", os.getenv('OpenOrchestratorSQL'),os.getenv('OpenOrchestratorKey'), None)

import json

data = {
    "Sagsnummer": "GEO-2024-043144",
    "MailModtager": "Gujc@aarhus.dk",
    "DeskProID": "2088",
    "DeskProTitel": "Aktindsigt i aktindsigter",
    "PodioID": "2931863091",
    "Overmappe": "2088 - Aktindsigt i aktindsigter",
    "Undermappe": "GEO-2024-043144 - GustavTestAktIndsigt2",
    "FilarkivCaseID": "dc35281b-4319-45b9-b32f-349d5d1834b7",
    "GeoSag": True,
    "NovaSag": False
}


json_string = json.dumps(data, indent=4)
print(json_string)

orchestrator_connection.create_queue_element("AktbobFromFilarkivToSharePoint","GEO-2024-043144",json_string)
