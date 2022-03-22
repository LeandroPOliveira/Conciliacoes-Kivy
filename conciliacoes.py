from kivy.properties import StringProperty, ListProperty, BooleanProperty
from kivy.uix.label import Label
from kivy.uix.popup import Popup
from kivymd.uix.dialog import MDDialog
from reportlab.pdfgen import canvas
from PyPDF2 import PdfFileWriter, PdfFileReader
from win32com import client
from kivymd.app import MDApp
from kivymd.uix.datatables import MDDataTable
from kivy.lang.builder import Builder
from kivy.uix.screenmanager import ScreenManager, Screen
from kivy.metrics import dp
import os
import pandas as pd
import openpyxl
import numpy as np
from datetime import datetime
from kivy.utils import get_color_from_hex
from dateutil.relativedelta import relativedelta
import getpass


class LoginWindow(Screen):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.dicio = {}
        self.lista_usuarios = []
        with open('usuarios.txt') as user:
            self.usuarios = user.readlines()
            for user in self.usuarios:
                self.lista_usuarios.append(user.split(';')[0])
                self.dicio.update({user.split(';')[0]: user.split(';')[1]})
    def verifica_usuario(self):
        if self.dicio[self.ids.spinner_id.text].strip() == getpass.getuser():
            self.manager.current = 'validar'
        else:
            self.dialog = MDDialog(text="Usuário inválido!", radius=[20, 7, 20, 7], )
            self.dialog.open()

class DataWindow(Screen):
    meu_status = StringProperty('Não Verificado')
    meu_status1 = StringProperty('Não Verificado')
    meu_status2 = StringProperty('Não Verificado')
    status_btn = BooleanProperty(False)
    icone = StringProperty("alert")
    icone1 = StringProperty("alert")
    icone2 = StringProperty("alert")
    cor, cor1, cor2 = ListProperty([1, 1, 0, 1]), ListProperty([1, 1, 0, 1]), ListProperty([1, 1, 0, 1])

    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.lista = []
        self.validos = []


        for i in range(12):
            mes = datetime.today()
            data_limite = mes - relativedelta(months=i)
            self.lista.append(data_limite.strftime('%m.%Y'))

    def mes_selecionado(self, text):
        self.text = text
        return self.text

    def status(self):  # Verifica se a situação das conciliações, se está validada ou pendente
        with open('dados.txt', 'r') as f:
            lines = f.readlines()
            self.meu_status = ''
            self.meu_status1 = ''
            self.meu_status2 = ''
            for i in lines:
                i = i.split(';')
                if i[0] == self.text and i[1] == self.lista_usuarios[0] and i[2].strip() == 'OK':
                    self.meu_status, self.icone, self.cor = 'Validado', 'check-circle', [0, 1, 0, 1]
                    self.validos.append('OK')
                elif i[0] == self.text and i[1] == self.lista_usuarios[1] and i[2].strip() == 'OK':
                    self.meu_status1, self.icone1, self.cor1 = 'Validado', 'check-circle', [0, 1, 0, 1]
                    self.validos.append('OK')
                elif i[0] == self.text and i[1] == self.lista_usuarios[2] and i[2].strip() == 'OK':
                    self.meu_status2, self.icone2, self.cor2 = 'Validado', 'check-circle', [0, 1, 0, 1]
                    self.validos.append('OK')
                else:
                    if self.meu_status == '':
                        self.meu_status, self.icone, self.cor = 'Validação Pendente', 'alert-circle', [1, 0, 0, 1]
                    if self.meu_status1 == '':
                        self.meu_status1, self.icone1, self.cor1 = 'Validação Pendente', 'alert-circle', [1, 0, 0, 1]
                    if self.meu_status2 == '':
                        self.meu_status2, self.icone2, self.cor2 = 'Validação Pendente', 'alert-circle', [1, 0, 0, 1]

        pegar = self.manager.get_screen('main')
        usuario = pegar.ids.spinner_id.text
        if usuario == 'Paulo França' and len(self.validos) >= 3:
            self.status_btn = True
        else:
            self.status_btn = False

    def assina_gestor(self):
        self.caminho = 'G:\GECOT\CONCILIAÇÕES CONTÁBEIS\CONCILIAÇÕES_' + self.ids.spinner_id2.text[3:7] + \
                       '\\' + self.ids.spinner_id2.text
        c = canvas.Canvas('watermark.pdf')
        # Draw the image at x, y. I positioned the x,y to be where i like here
        c.drawImage(self.manager.get_screen('main').ids.spinner_id.text + '.png', 350, 40, 150, 90,
                    mask='auto')
        c.save()
        watermark = PdfFileReader(
            open("watermark.pdf", "rb"))

        lista = []
        for file in os.listdir(self.caminho):
            if file.endswith(".pdf"):
                output_file = PdfFileWriter()
                with open(self.caminho + '\\' + file, "rb") as f:
                    input_file = PdfFileReader(f)
                    # Number of pages in input document
                    page_count = input_file.getNumPages()

                    # Go through all the input file pages to add a watermark to them
                    for page_number in range(page_count):
                        input_page = input_file.getPage(page_number)
                        if page_number == page_count - 1:
                            input_page.mergePage(watermark.getPage(0))
                        output_file.addPage(input_page)

                    with open(self.caminho + '\\' + file[8:], "wb") as outputStream:
                        output_file.write(outputStream)

                os.remove(self.caminho + '\\' + file)


class BoxTeste(Screen):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.data_tables = None

        self.lista_usuarios = []
        with open('usuarios.txt') as user:
            usuarios = user.readlines()
            for u in usuarios:
                self.lista_usuarios.append(u.strip())

    def validacao(self):
        pegar = self.manager.get_screen('validar')
        competencia = pegar.ids.spinner_id2.text
        self.caminho = 'G:\GECOT\CONCILIAÇÕES CONTÁBEIS\CONCILIAÇÕES_' + competencia[3:7] + \
                       '\\' + competencia
        pasta1 = os.listdir(self.caminho)
        lista = [[], [], [], [], []]
        for i in pasta1[::-1]:
            if i.startswith('~'):
                pasta1.remove(i)

        for i in pasta1:
            if i.endswith('.xlsx'):
                wb = openpyxl.load_workbook(self.caminho + '\\' + i, read_only=True)
                sheets = wb.sheetnames
                ws = wb[sheets[0]]
                try:
                    conta = ws['A2'].value.split()
                except:
                    conta = ['', '']
                valor_deb = ws['C5'].value
                valor_cred = ws['D5'].value
                try:
                    data1 = ws['A5'].value.strftime('%m.%Y')
                except:
                    data1 = ws['A5'].value
                lista[0].append(conta[1])
                lista[1].append(data1)
                lista[2].append(valor_deb)
                lista[3].append(valor_cred)

                wb.close()

        self.data = pd.DataFrame(lista).T

        self.data.columns = ['Conta', 'Data', 'Debito', 'Credito', 'Balancete']

        # Listar planilhas dos balancetes
        self.pasta = 'G:\GECOT\CONCILIAÇÕES CONTÁBEIS\CONCILIAÇÕES_' + competencia[3:7] + \
                '\BALANCETES\SOCIETÁRIOS\\'

        lista = os.listdir(self.pasta)

        lista4 = {}
        for i in lista:
            if competencia in i or competencia.replace('.', '') in i:
                tempo = os.path.getmtime(self.pasta + i)
                tempo2 = datetime.fromtimestamp(tempo)
                lista4.update({i: tempo2})

        dados = list(lista4.keys())[list(lista4.values()).index(max(lista4.values()))]
        dados = pd.read_excel(self.pasta + dados, skiprows=12)
        dados = pd.DataFrame(dados)
        apoio = pd.read_excel('contas.xlsx')
        apoio = pd.DataFrame(apoio)
        for index1, row1 in self.data.iterrows():
            for index, row in dados.iterrows():
                if row1['Conta'] == row['Conta CSPE']:
                    self.data['Balancete'].loc[index1] = dados.loc[index, ' Saldo Acumulado']

        self.data[['Debito', 'Credito', 'Balancete']] = self.data[['Debito', 'Credito', 'Balancete']].apply(pd.to_numeric)
        self.data.fillna(0, inplace=True)
        self.data = self.data.round(2)
        self.data['Conciliação'] = self.data['Debito'] - self.data['Credito']
        self.data.drop(['Debito', 'Credito'], axis=1, inplace=True)
        self.data['Diferença'] = self.data['Conciliação'] - self.data['Balancete']
        self.data = pd.merge(self.data, apoio[['Conta', 'Usuario']], on=['Conta'], how='left')
        self.data['Status'] = np.where(self.data['Diferença'] != 0, 'Diferença de Valor',
                                  (np.where(self.data['Data'] != competencia, 'Data Incorreta', 'OK')))
        if self.manager.get_screen('main').ids.spinner_id.text != self.lista_usuarios[3]:
            self.data = self.data.loc[self.data['Usuario'] == self.manager.get_screen('main').ids.spinner_id.text]
        self.data = self.data.round(2)
        self.resultado = self.data.to_records(index=False)
        self.resultado = list(self.resultado)
        if len(self.resultado) == 1:
            self.resultado.append(('', '', '', '', '', '', ''))
        print(self.resultado)
        self.add_datatable()

    def add_datatable(self):
        self.data_tables = MDDataTable(pos_hint={'center_x': 0.5, 'y': 0.2},
                                       size_hint=(0.8, 0.7),
                                       use_pagination=True, rows_num=10,
                                       background_color_header=get_color_from_hex("#03a9e0"),
                                       check=True,
                                       column_data=[("[color=#ffffff]Conta[/color]", dp(30)),
                                                    ("[color=#ffffff]Data[/color]", dp(20)),
                                                    ("[color=#ffffff]Balancete[/color]", dp(30)),
                                                    ("[color=#ffffff]Conciliação[/color]", dp(30)),
                                                    ("[color=#ffffff]Diferença[/color]", dp(20)),
                                                    ("[color=#ffffff]Usuario[/color]", dp(35)),
                                                    ("[color=#ffffff]Status[/color]", dp(35)),
                                                    ],
                                       row_data=self.resultado, elevation=1)

        self.add_widget(self.data_tables)
        self.data_tables.bind(on_check_press=self.checked)

    def checked(self, instance_table, current_row):
        pegar = self.manager.get_screen('validar')
        competencia = pegar.ids.spinner_id2.text
        self.caminho = 'G:\GECOT\CONCILIAÇÕES CONTÁBEIS\CONCILIAÇÕES_' + competencia[3:7] + \
                       '\\' + competencia
        os.startfile(self.caminho + '\\' + 'Conta ' + current_row[0].replace('.', '') + '.xlsx')

    def assinar(self):
        valida = self.data['Status'].unique()
        if 'OK' in valida and len(valida) == 1:
            lista3 = []
            for i in self.data['Conta']:
                conta = 'Conta ' + i.replace('.', '') + '.xlsx'
                lista3.append(conta)

            # Create the watermark from an image
            c = canvas.Canvas('watermark.pdf')
            # Draw the image at x, y. I positioned the x,y to be where i like here
            c.drawImage(self.manager.get_screen('main').ids.spinner_id.text + '.png', 40, 50, 150, 100,
                        mask='auto')
            c.save()

            for i in lista3:
                # Open Microsoft Excel
                excel = client.Dispatch("Excel.Application")
                # Read Excel File
                sheets = excel.Workbooks.Open(self.caminho + '\\' + i)
                work_sheets = sheets.Worksheets[0]
                # Convert into PDF File
                path = self.caminho + '\\' + 'teste ' + i.replace('.xlsx', '.pdf')
                work_sheets.ExportAsFixedFormat(0, path)

                # Get the watermark file you just created
                watermark = PdfFileReader(open("watermark.pdf", "rb"))
                # Get our files ready
                output = PdfFileWriter()

                with open(path, "rb") as provisorio:
                    input = PdfFileReader(provisorio)
                    number_of_pages = input.getNumPages()

                    for current_page_number in range(number_of_pages):
                        page = input.getPage(current_page_number)
                        if page.extractText() != "":
                            output.addPage(page)

                    page_count = output.getNumPages()
                    output2 = PdfFileWriter()
                    # Go through all the input file pages to add a watermark to them
                    for page_number in range(page_count):
                        input_page = output.getPage(page_number)
                        if page_number == page_count - 1:
                            input_page.mergePage(watermark.getPage(0))
                        output2.addPage(input_page)

                    # finally, write "output" to document-output.pdf
                    with open(self.caminho + '\\' + 'pendente' + i.replace('.xlsx', '.pdf'), "wb") as outputStream:
                        output2.write(outputStream)

                os.remove(path)
                sheets.Close(True)
                # excel.Quit()

            adicionar = [self.manager.get_screen('validar').ids.spinner_id2.text,
                         self.manager.get_screen('main').ids.spinner_id.text, 'OK']
            adicionar = ';'.join(adicionar)

            with open('dados.txt', 'a') as f:
                f.write(f'\n{adicionar}')

        else:
            self.dialog = MDDialog(text="Erro! Verificar pendências!", radius=[20, 7, 20, 7],)
            self.dialog.open()


class WindowManager(ScreenManager):
    pass


class Example(MDApp):
    popupWindow = None

    def meu_popup(self):
        Example.popupWindow = Popup(title='Conciliações',
                               content=Label(text='Gerando relatório...', font_size=20),
                               size_hint=(None, None), size=(400, 300), auto_dismiss=False)
        Example.popupWindow.open()

    def build(self):
        return Builder.load_file('conciliacoes.kv')


Example().run()
