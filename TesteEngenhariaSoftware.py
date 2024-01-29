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


    # Insere/Edita os itens da planilha no Google Sheets
    media_alunos = []
    for x, lista in enumerate(valores):
      if x > 2:
        resultadoIni = 0
        for y, lista in enumerate(lista):
          if (y > 2 and y < 5 and valores[x][y].strip() != ''):
            resultadoIni += float(valores[x][y])
        resultadoIni = resultadoIni / 2
        media_alunos.append([float(resultadoIni)])

    #Passando os valores de media_alunos para uma lista menos complexa
    media_alunos_final = [item[0] for item in media_alunos]
    print("Imprimindo a média inicial dos alunos...")
    print(media_alunos_final)

    valores_adicionar = [[]]

    #Retirando o total de aulas da planilha
    #isso permitirá um cálculo dinâmico do percentual de faltas
    for x, lista in enumerate(valores):
      if x == 1:
        for y, lista in enumerate(lista):
          totalAulas = valores[x][y]

    totalAulas = int(totalAulas.split(":")[1].strip())


    #Armazenando a quantidade de faltas de cada aluno
    quantidadeFaltas = []
    for x, lista in enumerate(valores):
      if x > 2:
        quantidadeFaltas.append(valores[x][2])

    #Armazenando o percentual de falta de cada aluno
    quantidadePercentualFaltas = []
    for x in quantidadeFaltas:
      percentualFaltas = (float(x) / totalAulas) * 100
      quantidadePercentualFaltas.append(int(percentualFaltas))
    print("Imprimindo o percentual de falta de cada aluno em uma lista...")
    print(quantidadePercentualFaltas)

    #Código que verifica a situação do aluno e dita se passou ou não (por nota)
    for x, lista in enumerate(media_alunos):
      for i in lista:
        if quantidadePercentualFaltas[x] > 25:
          valores_adicionar.append(["Reprovado por Falta"])
        else:
          nota = float(i)
          if nota >= 70:
            valores_adicionar.append(["Aprovado"])
          elif nota < 50:
            valores_adicionar.append(["Reprovado por Nota"])
          else:
            valores_adicionar.append(["Exame Final"])

    result = (
      sheet.values()
      .update(spreadsheetId=SAMPLE_SPREADSHEET_ID, range='G3:G{}'.format(len(valores_adicionar) +3),
              valueInputOption='RAW', body={'values': valores_adicionar})
      .execute()
    )

    #Pegando as notas da P3
    notasP3 = []
    for x, lista in enumerate(valores):
      if x > 2:
        notasP3.append(float(valores[x][5]))
    print("Imprimindo as notas da P3 de cada aluno...")
    print(notasP3)

    #Código que verifica as notas dos alunos em situação de "Exame Final"
    situacaoAlunos = []
    resultadoFinal = []
    for x, lista in enumerate(valores):
      if x > 2:
        situacaoAlunos.append(valores[x][6])

    for x, i in enumerate(situacaoAlunos):
      if i == "Exame Final":
        resultado = (float(media_alunos_final[x]) + float(notasP3[x])) / 2
        resultadoFinal.append(resultado)
      else:
        resultadoFinal.append(0)


    #Anexando os valores de resultadoFinal na planilha
    valores_adicionar = []
    for x in range(len(resultadoFinal)):
      valores_adicionar.append([int(resultadoFinal[x])])

    result = (
      sheet.values()
      .update(spreadsheetId=SAMPLE_SPREADSHEET_ID, range='H4:H{}'.format(len(valores_adicionar) + 3),
              valueInputOption='RAW', body={'values': valores_adicionar})
      .execute()
    )
    print("Notas finais dos alunos em situação 'Exame Final' inseridos na planilha")
    print(resultadoFinal) ##Visualização para verificar ciclo de execução do programa (apenas)

  except HttpError as err:
    print(err)


if __name__ == "__main__":
  main()

##Criado por: Gabriel André