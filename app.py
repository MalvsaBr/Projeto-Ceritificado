import openpyxl as op
from PIL import Image, ImageDraw, ImageFont

workbook_alunos = op.load_workbook("caminho_arquivo_xlsx")
sheet_alunos = workbook_alunos['aba_escolhida']

for indice, linha in enumerate(sheet_alunos.iter_rows(min_row=2,max_row=20)):
    nome_curso = linha[0].value
    nome_participante = linha[1].value
    tipo_participacao = linha[2].value
    data_inicio = linha[3].value
    data_fim = linha[4].value
    carga_horaria = linha[5].value
    data_emissao = linha[6].value

    fonte_nome = ImageFont.truetype("caminho_fonte",90)
    fonte_geral = ImageFont.truetype("caminho_fonte",80)
    fonte_data = ImageFont.truetype("caminho_fonte",55)

    imagem = Image.open("caminho_fonte")
    desenhar = ImageDraw.Draw(imagem)

    desenhar.text((1020,833),nome_participante,fill='black',font=fonte_nome)
    desenhar.text((1060,957),nome_curso,fill='black',font=fonte_geral)
    desenhar.text((1435,1065),tipo_participacao,fill='black',font=fonte_geral)
    desenhar.text((1480,1182),str(carga_horaria),fill='black',font=fonte_geral)
    desenhar.text((750,1770),data_inicio,fill='black',font=fonte_data)
    desenhar.text((750,1930),data_fim,fill='black',font=fonte_data)
    desenhar.text((2220,1930),data_emissao,fill='black',font=fonte_data)

    imagem.save(f'"caminho_salvo_arquivo" {indice} {nome_participante} - Certificado.png')
