import openpyxl
from PIL import Image, ImageDraw, ImageFont

workbook_alunos = openpyxl.load_workbook('planilha_alunos.xlsx')
sheet_alunos = workbook_alunos['Sheet1']

for indice, linha in enumerate (sheet_alunos.iter_rows(min_row= 2)):
    nome_curso = linha[0].value
    nome_participacao = linha[1].value
    tipo_participacao = linha [2].value
    data_inicio = linha[3].value
    data_final = linha[4].value 
    carga_horaria = linha[5].value
    data_emissao = linha[6].value
    
    font_nome = ImageFont.truetype('./tahomabd.ttf',90)
    font_geral = ImageFont.truetype('./tahoma.ttf',80)
    
    image = Image.open('./certificado_padrao.jpg')
    desenhar = ImageDraw.Draw(image)
    
    desenhar.text((1020,827),nome_participacao,fill='black',font = font_nome)
    desenhar.text((1060,950),nome_participacao,fill='black',font = font_geral)
    desenhar.text((1435,1065),nome_participacao,fill='black',font = font_geral)
    desenhar.text((1480,1182), str(carga_horaria), fill='black', font = font_geral)
    
    image.save(f'./{indice} {nome_participacao} certificado.png')