# TesteEngenhariaSoftware
![logoTuntsRocksHeader c8146752](https://github.com/gabrielandre-math/TesteEngenhariaSoftwareTuntsRocks/assets/60861872/bdd7128b-0e62-4a71-a783-13c4c8f6e11e)

## Google Sheets API Integration

Este é um script Python que integra a API do Google Sheets para visualizar, inserir e editar dados em uma planilha Google Sheets. O script realiza as seguintes operações:

## Autenticação e Autorização:

- Verifica se há credenciais salvas no arquivo `token.json`. Se não houver credenciais válidas, solicita ao usuário que faça login.
- As credenciais são salvas para serem usadas em execuções futuras.

## Acesso à Planilha:

- Acesso à planilha com o ID `1udL851r-Br1-wkvs_FFqffOOgFVoR0OmAmXNAs_YYd0`.
- Leitura dos dados da faixa `engenharia_de_software!A1:I28`.

## Cálculo da Média Inicial:

- Calcula a média inicial dos alunos com base nas notas das colunas 3 e 4.

## Inserção/Edição de Dados:

- Insere ou edita os dados da média inicial na coluna G.

## Cálculo do Percentual de Faltas:

- Calcula o percentual de faltas para cada aluno com base nas faltas totais e armazena em uma lista.

## Verificação da Situação do Aluno:

- Determina se cada aluno está "Aprovado", "Reprovado por Falta", "Reprovado por Nota" ou em "Exame Final".

## Atualização da Situação na Planilha:

- Atualiza a situação dos alunos na coluna G da planilha.

## Obtenção das Notas da P3:

- Obtém as notas da P3 para cada aluno.

## Cálculo do Resultado Final para Alunos em Exame Final:

- Calcula o resultado final para alunos em "Exame Final" com base na média e nas notas da P3.

## Atualização do Resultado Final na Planilha:

- Atualiza os resultados finais na coluna H da planilha.

## Execução do Script
```bash
python TesteEngenhariaSoftware.py
```
Certifique-se de ter as credenciais de API no arquivo `credentials.json` e as bibliotecas necessárias instaladas. Execute o script para interagir com a planilha.

## Link da planilha preenchida
https://docs.google.com/spreadsheets/d/1udL851r-Br1-wkvs_FFqffOOgFVoR0OmAmXNAs_YYd0/edit?usp=sharing

**Criado por:** Gabriel André
