import openpyxl
import time
import win32com.client
import tkinter as tk
from tkinter import filedialog

# Oculta a janela do Tkinter
root = tk.Tk()
root.withdraw()

# Seleciona o arquivo Excel
planilha_path = filedialog.askopenfilename(
    title="Selecione o arquivo da planilha de remessas",
    filetypes=[("Arquivos Excel", "*.xlsx *.xls")]
)

if not planilha_path:
    print("Nenhum arquivo selecionado, encerrando.")
    input("Pressione Enter para sair...")
    exit()

wb = openpyxl.load_workbook(planilha_path)
ws = wb['Sheet1']

# Conecta ao SAP
try:
    SapGuiAuto = win32com.client.GetObject("SAPGUI")
    application = SapGuiAuto.GetScriptingEngine
    connection = application.Children(0)
    session = connection.Children(0)
except Exception as e:
    print("Erro ao conectar ao SAP:")
    print(e)
    input("Pressione Enter para sair...")
    exit()

# Loop nas remessas da planilha
for row in ws.iter_rows(min_row=2, values_only=True):
    remessa = str(row[0])
    data_entrega = row[1].strftime('%d.%m.%Y')

    print(f"Atualizando remessa {remessa} com data {data_entrega}")

    session.StartTransaction("VL02N")
    session.FindById("wnd[0]").Maximize()
    session.FindById("wnd[0]/usr/ctxtLIKP-VBELN").Text = remessa
    session.FindById("wnd[0]").SendVKey(0)
    time.sleep(1)

    # Go to > Header > Dates
    session.FindById("wnd[0]/mbar/menu[2]/menu[1]/menu[11]").Select()
    time.sleep(1)

    # Campo "End Actual"
    field_path = "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\11/ssubSUBSCREEN_BODY:SAPMV50A:2122/" + \
                 "subTSEG_STD:SAPLTSED:0100/tblSAPLTSEDTC_TSEG_STD/ctxtITSEGDIAE-TIME_TST04[10,0]"
    session.FindById(field_path).SetFocus()
    session.FindById(field_path).Text = data_entrega
    session.FindById(field_path).CaretPosition = 2

    # Salvar e voltar
    session.FindById("wnd[0]/tbar[0]/btn[11]").Press()  # Gravar
    time.sleep(1)
    session.FindById("wnd[0]/tbar[0]/btn[15]").Press()  # Voltar
    time.sleep(1)

print(" Todas as remessas foram atualizadas com sucesso!")
input("Pressione Enter para sair...")
