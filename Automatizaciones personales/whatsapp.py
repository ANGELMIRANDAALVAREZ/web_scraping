from twilio.rest import Client

# Tus credenciales de Twilio
account_sid = 'TU_ACCOUNT_SID'
auth_token = 'TU_AUTH_TOKEN'

# Inicializa el cliente de Twilio
client = Client(account_sid, auth_token)

# Detalles del mensaje
from_whatsapp_number = 'whatsapp:NUMERO_SANDBOX'  # El número de WhatsApp Sandbox de Twilio
to_whatsapp_number = 'whatsapp:NUMERO_DESTINO'  # Tu número de WhatsApp con el prefijo de país
message_body = 'Hola, este es un mensaje enviado desde Python!'

# Enviar el mensaje
message = client.messages.create(
                              from_=from_whatsapp_number,
                              body=message_body,
                              to=to_whatsapp_number
                          )

print(message.sid)
