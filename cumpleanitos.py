import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime, timedelta
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import os.path
import pickle
from dotenv import load_dotenv
import os


# Autenticación y conexión a Google Sheets
def autenticar_google_sheets():
    load_dotenv()
    CRED = os.getenv('G_CLIENT_SECRET')
    TOKEN = os.getenv('G_TOKEN')
    
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
    creds = None
    if os.path.exists(TOKEN):
        with open(TOKEN, 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CRED, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN, 'wb') as token:
            pickle.dump(creds, token)
    service = build('sheets', 'v4', credentials=creds)
    return service


def leer_google_sheets(sheet_id, rango):
    service = autenticar_google_sheets()
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=sheet_id, range=rango).execute()
    values = result.get('values', [])
    df = pd.DataFrame(values[1:], columns=values[0])  # Crea el DataFrame con los datos
    return df


# Función para enviar el correo electrónico
def enviar_correo(asunto, mensaje):
    load_dotenv()
    remitente = os.getenv('G_MAIL')  # Reemplaza con tu correo
    password = os.getenv('G_PASS')   # Reemplaza con tu contraseña
    # destinatario = 'tomas_koda@live.com.ar,tomas98salazar@gmail.com'
    destinatario = 'tomas98salazar@gmail.com,salazarmd146@gmail.com,marcelaviviana0603@gmail.com,joaco.jq@gmail.com'

    msg = MIMEMultipart()
    msg['From'] = remitente
    msg['To'] = destinatario
    msg['Subject'] = asunto

    msg.attach(MIMEText(mensaje, 'plain'))

    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(remitente, password)
    text = msg.as_string()
    server.sendmail(remitente, destinatario.split(','), text)
    server.quit()


# Función para revisar los cumpleaños y enviar avisos
def revisar_cumpleaños(df):
    # Extraer día y mes de la fecha actual
    hoy = datetime.now().date()
    dia_hoy = hoy.day
    mes_hoy = hoy.month
    
    for index, row in df.iterrows():
        fecha_cumple = datetime.strptime(row['Fecha de cumpleaños'], '%d/%m/%Y').date()
        fecha_cumple_avisar = fecha_cumple - timedelta(days=1)
        
        # Extraer día, mes y año de la fecha de cumpleaños
        dia_cumple = fecha_cumple_avisar.day
        mes_cumple = fecha_cumple_avisar.month
        fecha_mail = fecha_cumple.strftime('%d/%m')
        
        # dias_aviso = int(row['Días de aviso']) if 'Días de aviso' in df.columns else 1
        # if hoy == fecha_cumple - timedelta(days=dias_aviso):
        if dia_hoy == dia_cumple and mes_hoy == mes_cumple:  # Aviso un día antes
            
            # Construir el cuerpo del mensaje
            cuerpo_mensaje = (
                f"El cumpleaños de {row['Nombre']} {row['Apellido']} es el {fecha_mail}.\n"
                "¡No olvides desearle un lindo día!\n\n"
            )
            if row['Celular']:
                numero = row['Celular'].strip()  # Asegurarse de que no haya espacios
                numero_formateado = '+54'+numero
                # Enlace para WhatsApp
                enlace_whatsapp = f"https://wa.me/{numero_formateado}"
                cuerpo_mensaje += (
                    f"Podes llamar al número: {numero}\n"
                    f"También puedes enviarle un WhatsApp: {enlace_whatsapp}\n"
                )
            if row['Correo electrónico']:
                cuerpo_mensaje += f"Su email es: {row['Correo electrónico']}\n\n¡Saludos!"
                
            
            enviar_correo(
                f"Cumpleaños {row['Nombre']} {row['Apellido']}",
                cuerpo_mensaje
            )


def main():
    # Ejecución del script
    sheet_id = '1TACiVok4vE0U_F5e_6qouobZIB-5KMRN2bQFrWbX5y0'  # ID de la hoja
    rango = 'Cumple!A:E'     # Rango de tu hoja
    # rango = 'Cumpleanitos'        # Nombre de tabla (o rango de tabla)
    df_google = leer_google_sheets(sheet_id, rango)
    revisar_cumpleaños(df_google)


if __name__ == '__main__':
    main()