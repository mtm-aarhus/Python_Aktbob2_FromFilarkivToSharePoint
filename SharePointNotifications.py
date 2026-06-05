from SendSMTPMail import send_email


def send_missing_sharepoint_delivery_folder_email(
    receiver,
    Sagsnummer,
    Overmappe,
    Undermappe,
    folder_url,
):
    sender = "aktbob@aarhus.dk"
    subject = f"FEJL! {Sagsnummer} mangler udleveringsmappe i SharePoint"
    body = f"""Kære sagsbehandler,<br><br>
    Robotten kan ikke overføre dokumenter til sagen: {Sagsnummer}, fordi udleveringsmappen ikke kan findes i SharePoint.<br><br>
    Det sker typisk, hvis mappen er blevet omdøbt, flyttet eller slettet.<br><br>
    Forventet overmappe:<br>
    {Overmappe}<br><br>
    Forventet undermappe:<br>
    {Undermappe}<br><br>
    Forventet SharePoint sti:<br>
    {folder_url}<br><br>
    Ret mappenavnet tilbage eller opret mappen igen, inden du forsøger steppet 'Overfør dokumenter til udleveringsmappe (Sharepoint)' igen.<br><br>
    Det anbefales at følge <a href="https://aarhuskommune.atlassian.net/wiki/spaces/AB/pages/64979049/AKTBOB+--+Vejledning">vejledningen</a>,
    hvor du også finder svar på de fleste spørgsmål og fejltyper.
    """
    smtp_server = "smtp.adm.aarhuskommune.dk"
    smtp_port = 25

    send_email(
        receiver=receiver,
        sender=sender,
        subject=subject,
        body=body,
        smtp_server=smtp_server,
        smtp_port=smtp_port,
        html_body=True,
    )
