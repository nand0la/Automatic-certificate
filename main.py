from PIL import Image, ImageFont, ImageDraw
import xlrd
import sys


class ListCertificate:
    def __init__(
        self,
        font_path,
        template_path,
        excel_file_path,
        index_sheet=0,
        index_colunm_name=0,
    ):
        self.font_path = font_path
        self.template_path = template_path
        self.excel_file_path = excel_file_path
        self.index_sheet = index_sheet
        self.index_colunm_name = index_colunm_name

    def return_list_students(self):
        table_student = xlrd.open_workbook(self.excel_file_path)
        sheet_students = table_student.sheet_by_index(self.index_sheet)

        list_students = []
        var_control = 1
        while var_control < sheet_students.nrows:
            cell_value = sheet_students.cell_value(var_control, self.index_colunm_name)
            list_students.append(cell_value)
            var_control += 1

        return list_students

    def listar_alunos(self, lista_alunos):
        for aluno in lista_alunos:
            print(aluno)


# ===================== \\ ==========================

font_path = (
    "/home/fernando/Documentos/projetos/Automatic-certificate/BalsamiqSans-Regular.ttf"
)
template_path = "/home/fernando/Documentos/projetos/Automatic-certificate/template_certificado/template_certificado1.png"
excel_file_path = "/home/fernando/Documentos/projetos/Automatic-certificate/planilha_alunos/alunos.xlsx"

certificados = ListCertificate(font_path, template_path, excel_file_path)
alunos_form_lista = certificados.return_list_students()
certificados.listar_alunos(alunos_form_lista)

# ======================= \\ ========================


# def fazer_convite(lista):
#
#     list_students = list(lista)
#     img_font = ImageFont.truetype(r"BalsamiqSans-Regular.ttf", 90)
#     img_directory = ""
#     position_font = [0, 0]
#
#     if int(certificate_model) == 1:
#         img_directory = "template_certificado/template_certificado1.png"
#         position_font = [370, 280]
#     elif int(certificate_model) == 2:
#         img_directory = "template_certificado/template_certificado2.png"
#         position_font = [320, 320]
#     elif int(certificate_model) == 3:
#         img_directory = "template_certificado/template_certificado3.png"
#         position_font = [320, 320]
#
#     for nome_aluno in list_students:
#         img_tamplate = Image.open(img_directory)  # abrindo a imagem
#         img_draw = ImageDraw.Draw(img_tamplate)  # imagem editavel para desenho
#
#         img_draw.text(
#             (position_font[0], position_font[1]), nome_aluno, font=img_font, fill="#000"
#         )  # adicionando texto na imagem editvael pra deneho
#         img_tamplate.show()
#
#
# fazer_convite(students_name)
