import smtplib
import ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from tkinter import messagebox


def send_email_alert(ip_address, printer_name, diff, alert_thresholds):
    # Vérifiez si l'adresse IP a un seuil d'alerte spécifique
    if ip_address in alert_thresholds:
        threshold = alert_thresholds[ip_address]

        # Vérifiez si l'écart atteint le seuil d'alerte
        if diff >= threshold:
            # Remplacez les valeurs par les informations de votre compte e-mail et les destinataires appropriés
            email_from = 'adresse_mail@mail.fr'
            email_to = ['destinataire1@mail.fr', 'destinataire2@mail.fr']
            smtp_server = 'smtp.xxx.com'
            smtp_port = 587
            smtp_username = 'mail_login_address'
            smtp_password = 'password_mail_login_address'

            subject = f"Alerte : Dépassement du seuil d'impression pour l'imprimante {printer_name}"
            body = f"Le compteur de pages de l'imprimante {printer_name} à l'adresse {ip_address} a atteint un écart de {diff}."

            # Créez un objet MIMEMultipart pour le message e-mail
            msg = MIMEMultipart()
            msg['From'] = email_from
            msg['To'] = ', '.join(email_to)
            msg['Subject'] = subject

            # Ajoutez le corps du message au format texte brut
            msg.attach(MIMEText(body, 'plain'))
            
             # Configuration du proxy SOCKS (retirer les " des deux lignes en dessous pour activer)
            #socks.setdefaultproxy(socks.PROXY_TYPE_SOCKS5, "http://127.0.0.1:9000/localproxy-67a1a4ae.pac",)
            #socket.socket = socks.socksocket

            try:
                # Établissez une connexion SMTP avec SSL et envoyez l'e-mail
                context = ssl.create_default_context()
                with smtplib.SMTP(smtp_server, smtp_port) as smtp:
                    smtp.starttls(context=context)
                    smtp.login(smtp_username, smtp_password)
                    smtp.send_message(msg)
                return True 

            except Exception as e:
                messagebox.showerror('Erreur', f'Erreur lors de l\'envoi de l\'e-mail : {str(e)}')
                return False 

def send_email_alert_serial_number_change(ip_address, printer_name, last_serial_number, current_serial_number):
    # Code pour envoyer un e-mail d'alerte
    # Remplacer les lignes suivantes par votre code d'envoi d'e-mail
    subject = f"Alerte : Changement de numéro de série pour l'imprimante {printer_name}"
    body = f"Le numéro de série de l'imprimante {printer_name} à l'adresse {ip_address} a changé.\nNuméro de série précédent : {last_serial_number}\nNuméro de série actuel : {current_serial_number}"
