import xlrd
import sys

from PIL import Image, ImageFont, ImageDraw
from tkinter import Tk
from tkinter.filedialog import askopenfilename


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

    def do_certificate(
        self,
        lista,
        pos_font=[100, 100],
    ):
        img_font = ImageFont.truetype(self.font_path, 90)

        for student_name in lista:
            img_tamplate = Image.open(self.template_path)  # abrindo a imagem
            img_draw = ImageDraw.Draw(img_tamplate)  # imagem editavel para desenho

            img_draw.text(
                (pos_font[0], pos_font[1]),
                student_name,
                font=img_font,
                fill="#000",
            )  # adicionando texto na imagem editvael para desenho

            img_tamplate.show()


if __name__ == "__main__":
    Tk().withdraw()  # we don't want a full GUI, so keep the root window from appearing

    font_path = askopenfilename(title="Select the font")
    template_path = askopenfilename(title="Select the image template")
    excel_file_path = askopenfilename(title="select the excel file")

    certificados = ListCertificate(font_path, template_path, excel_file_path)
    alunos_form_lista = certificados.return_list_students()
    certificados.do_certificate(alunos_form_lista)
