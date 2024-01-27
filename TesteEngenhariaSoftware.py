import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = "1udL851r-Br1-wkvs_FFqffOOgFVoR0OmAmXNAs_YYd0"
SAMPLE_RANGE_NAME = "engenharia_de_software!A1:I28"


def main():
  """Shows basic usage of the Sheets API.
  Prints values from a sample spreadsheet.
  """
  creds = None
  # The file token.json stores the user's access and refresh tokens, and is
  # created automatically when the authorization flow completes for the first
  # time.
  if os.path.exists("token.json"):
    creds = Credentials.from_authorized_user_file("token.json", SCOPES)
  # If there are no (valid) credentials available, let the user log in.
  if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
      creds.refresh(Request())
    else:
      flow = InstalledAppFlow.from_client_secrets_file(
          "credentials.json", SCOPES
      )
      creds = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open("token.json", "w") as token:
      token.write(creds.to_json())

  try:
    service = build("sheets", "v4", credentials=creds)

    # Visualiza os itens da planilha Google Sheets

    sheet = service.spreadsheets()
    result = (
        sheet.values()
        .get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=SAMPLE_RANGE_NAME)
        .execute()
    )
    valores = result['values']
    #print(valores)


    # Insere/Edita os itens da planilha no Google Sheets


    resultadoIni = 0

    valores_adicionar = []

    media_alunos = []
    for x, lista in enumerate(valores):
      if x > 2:
        resultadoIni = 0
        for y, lista in enumerate(lista):
          if (y > 2 and y < 6 and valores[x][y].strip() != ''):
            resultadoIni += float(valores[x][y])
        resultadoIni = resultadoIni / 3
        valores_adicionar.append([int(resultadoIni)])
        media_alunos.append([float(resultadoIni)/10])
    # Remover o zero no início de cada valor em valores_adicionar
    valores_adicionar = [[str(valor[0]).lstrip('0')] for valor in valores_adicionar]

    result = (
      sheet.values()
      .update(spreadsheetId=SAMPLE_SPREADSHEET_ID, range='H4:H{}'.format(len(valores_adicionar) + 3),
              valueInputOption='RAW', body={'values': valores_adicionar})
      .execute()
    )

    valores_adicionar = [[]]

    #Código que verifica a situação do aluno e dita se passou ou não (por nota)
    for x, lista in enumerate(media_alunos):
      for i in lista:
        if float(i) >= 7.0:
          valores_adicionar.append(["Aprovado"])
        else:
          if float(i) < 5.0:
            valores_adicionar.append(["Reprovado por Nota"])
          else:
            if float(i) >= 5 and float(i) < 7:
              valores_adicionar.append(["Exame Final"])


    for  x in media_alunos:
      print(x)

    result = (
      sheet.values()
      .update(spreadsheetId=SAMPLE_SPREADSHEET_ID, range='G3:G{}'.format(len(valores_adicionar) +3),
              valueInputOption='RAW', body={'values': valores_adicionar})
      .execute()
    )

    #result = (
      #sheet.values()
      #.update(spreadsheetId=SAMPLE_SPREADSHEET_ID, range='H4', valueInputOption='USER_ENTERED', body={'values': valores_adicionar})
      #.execute()
   #)





  except HttpError as err:
    print(err)


if __name__ == "__main__":
  main()