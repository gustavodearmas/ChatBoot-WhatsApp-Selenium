from email.mime.multipart import MIMEMultipart
from jinja2 import Environment, FileSystemLoader
import os
import sys
import getopt
from smtplib import SMTP as SMTP
from email.mime.text import MIMEText
from excel_conexion import datos_correo

SMTP_SERVER = "smtp.office365.com"
SMTP_PORT = 587

env = Environment(
    loader=FileSystemLoader('%s/templates/' % os.path.dirname(__file__)))

def get_data(dis):
    data = dis
    return data


def main(argv, body, historial):
    """The entry point for the script"""
    sender = "gadea@compensar.com"
    password = "2shcrveW-"
    cc = ""

    try:
        opts, _args = getopt.getopt(argv, 'u:p:', ['user=', 'password='])
    except getopt.GetoptError:
        print('sendactionablemessage.py -u <username> -p <password>')
        sys.exit(2)

    for opt, arg in opts:
        if opt == '-u':
            sender = arg
        elif opt == '-p':
            password = arg

    if (not sender) or (not password):
        print('sendactionablemessage.py -u <username> -p <password>')
        sys.exit(2)


    send_message(sender, password, body, historial, cc)


def send_message(sender, password, body, historial, cc):
    msg = MIMEText(body, 'html')
    msg['Subject'] = "REDENCIÓN CURSO COMPENSAR PRESENCIAL _ " + list(historial.values())[0][2][1]
    msg['From'] = sender
    msg['Cc'] = cc
    msg['To'] = list(historial.values())[0][4]

    conn = SMTP(SMTP_SERVER, SMTP_PORT)
    try:
        conn.starttls()
        conn.set_debuglevel(False)
        conn.login(sender, password)
        conn.sendmail(sender, list(historial.values())[0][4], msg.as_string())
    finally:
        conn.quit()

    print('EMAIL ENVIADO: ', msg['To'])


def enviar_mail(g, HISTORIAL_):
    json_data = g
    template = env.get_template('child.html')
    output = template.render(data=json_data)
    main(sys.argv[1:], output, HISTORIAL_)
