import openpyxl
from PIL import ImageFont, Image, ImageDraw

workbook_formandos = openpyxl.load_workbook('arquivos\\formandos.xlsx')
sheet_formandos = workbook_formandos['formandos']

for indice, linha in enumerate(sheet_formandos.iter_rows(min_row=3)):
    nome = linha[1].value
    nome_curso = linha[2].value
    rg = linha[3].value
    data_conclusao = linha[4].value
    carga_horaria = linha[5].value
    

    font_geral = ImageFont.truetype('fontes\\CENTURY.TTF',90)
    font_nome = ImageFont.truetype('fontes\\CURLZ___.TTF',90)

    imagem = Image.open('arquivos\\modelo_diploma.png')

    desenhar = ImageDraw.Draw(imagem)

    desenhar.text((850,850),nome, fill='black', font=font_nome)
    desenhar.text((850,950),nome_curso, fill='black', font=font_nome)
    # desenhar.text((650,1200),data_conclusao, fill='black', font=font_nome)
    desenhar.text((1300,1150),carga_horaria, fill='black', font=font_nome)

    imagem.save(f'diplomas_emitidos\\{indice}{nome}.png')