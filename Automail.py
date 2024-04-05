import win32com.client as win32

def envia_email(desti_email, assunto, texto, attachment_path=None):
  """
  Manda um email pelo Outlook (Precisa instalar e configurar) e precisa definir os argumentos abaixo.

  Args:
      desti_email (str): The email address.
      assunto (str): The subject of the email.
      texto (str): The body of the email in HTML format.
      attachment_path (str, optional): The path to the file to attach. Defaults => None.

  Return:
      bool: True se email foi enviado, False se não foi enviado.
  """

  try:
    outlook = win32.Dispatch('outlook.application')
    email = outlook.CreateItem(0)

    email.To = desti_email
    email.Subject = assunto
    email.HTMLBody = texto

    if attachment_path:
      email.Attachments.Add(attachment_path)

    email.Send()
    print("Worked!")
    return True

  except Exception as e:
    print("Error:", e)
    return False

# Example usage
desti_email = "your_email@example.com"  #Quem irá receber
assunto = "Test Email from Python"      #Assunto do 
texto = """
<p>Olá</p>
<p>Test e-mail.</p>
"""
#attachment_path = "C:/path/to/your/file.txt"  # Optional, para enviar Anexo

envia_email(desti_email, assunto, texto, attachment_path)