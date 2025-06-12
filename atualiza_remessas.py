import openpyxl
import time
import win32com.client
import os
import tkinter as tk
from tkinter import filedialog

root = tk.Tk()
root.withdraw()

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

try:
    SapGuiAuto = win32com.client.GetObject("SAPGUI")
    application = SapGuiAuto.GetScriptingEngine
    connection = application.Children(0)
    session = connection.Children(0)
except Exception as e:
    print(" SAP GUI não está aberto ou nenhuma sessão ativa!")
    print(e)
    input("Abra o SAP e faça login antes de rodar o script. Pressione Enter para sair...")
    exit()

for row in ws.iter_rows(min_row=2, max_row=4, values_only=True):
    remessa = str(row[0])
    data_entrega = row[1].strftime('%d.%m.%Y')

    print(f"Atualizando Remessa {remessa} para {data_entrega}")

    session.StartTransaction("VL02N")
    time.sleep(1)

    session.FindById("wnd[0]/usr/ctxtLIKP-VBELN").Text = remessa
    session.FindById("wnd[0]").SendVKey(0)
    time.sleep(1)

    session.FindById("wnd[0]/mbar/menu[2]/menu[1]").Select()
    session.FindById("wnd[1]/usr/tabsTABSTRIP_OVERVIEW/tabpT\\\\02").Select()
    session.FindById("wnd[1]/usr/subSUBSCREEN_HEADER:SAPLV50G:1105/ctxtLIKP-LFDAT").Text = data_entrega

    session.FindById("wnd[1]/tbar[0]/btn[11]").Press()
    session.FindById("wnd[0]/tbar[0]/btn[11]").Press()
    time.sleep(2)

print(" Atualização de remessas concluída com sucesso!")
input("Pressione Enter para sair...")
