from PIL import Image, ImageFont, ImageDraw
import xlrd as x


file_name = input('Digite o nome do arquivo: ')
certificate_model = input('Coloque o modelo do certificado: [1] [2] [3] ')


# ===================== \\ ==========================


def return_list_students(arquivo):
    table_stud = x.open_workbook(f'planilha_alunos/{arquivo}')
    sheet_students = table_stud.sheet_by_index(0)
    num_rows = sheet_students.nrows

    list_students = []

    i = 1
    while i < num_rows:
        cell_value = sheet_students.cell_value(i, 0)
        list_students.append(cell_value)
        
        i += 1

    return list_students

students_name = return_list_students(file_name)


# ======================= \\ ========================


def fazer_convite(lista):

    list_students = list(lista)
    img_font = ImageFont.truetype(r'BalsamiqSans-Regular.ttf', 90)
    img_directory = ''
    position_font = [0, 0]

    if int(certificate_model) == 1:
        img_directory = 'template_certificado/template_certificado1.png'
        position_font = [370, 280]
    elif int(certificate_model) == 2:
        img_directory = 'template_certificado/template_certificado2.png'
        position_font = [320, 320]
    elif int(certificate_model) == 3:
        img_directory = 'template_certificado/template_certificado3.png'
        position_font = [320, 320]

    for nome_aluno in list_students:
        img_tamplate = Image.open(img_directory)  # abrindo a imagem
        img_draw = ImageDraw.Draw(img_tamplate)  # imagem editavel para desenho

        img_draw.text((position_font[0], position_font[1]), nome_aluno, font=img_font, fill="#000")  # adicionando texto na imagem editvael pra deneho
        img_tamplate.show()
    
fazer_convite(students_name)
