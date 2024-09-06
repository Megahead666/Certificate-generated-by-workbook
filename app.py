import openpyxl
from PIL import Image, ImageDraw, ImageFont

workbook_alunos = openpyxl.load_workbook('planilha.xlsx')
sheet_alunos = workbook_alunos['Sheet1']

for indice, linha in enumerate(sheet_alunos.iter_rows(min_row=2, max_row=10)):
    nome_pessoa = linha[0].value
    nome_skill = linha[1].value
    data_to = linha[2].value
    data_until = linha[3].value

    font_nome = ImageFont.truetype('./tahomabd.ttf',90)
    font_geral = ImageFont.truetype('./tahoma.ttf',40)

    
    certificado = Image.open('./certificate.png')
    
    desenhar = ImageDraw.Draw(certificado)
    
    desenhar.text((600, 610), nome_pessoa, fill='grey', font=font_nome)

    desenhar.text((1180, 760), nome_skill, fill='grey', font=font_geral)

    desenhar.text((250, 759), data_to, fill='grey', font=font_geral)
    
    desenhar.text((250, 820), data_until, fill='grey', font=font_geral)
    
    certificado.save(f'./{indice} {nome_pessoa}certificado.png')
