import openpyxl
import time
import win32com.client
import tkinter as tk
from tkinter import filedialog
from datetime import datetime

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

log_path = "log_remessas.txt"
with open(log_path, "w", encoding="utf-8") as log:
    log.write(f"LOG DE EXECUÇÃO - {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n")
    log.write("-" * 50 + "\n")

    # Loop nas remessas
    for idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        remessa = str(row[0])
        data_raw = row[1]

        if not remessa or not data_raw:
            msg = f"[AVISO] Linha {idx} ignorada - remessa ou data vazia."
            print(msg)
            log.write(msg + "\n")
            continue

        data_entrega = data_raw.strftime('%d.%m.%Y')
        print(f"Processando linha {idx} | Remessa: {remessa} | Data: {data_entrega}")

        try:
            session.StartTransaction("VL02N")
            session.FindById("wnd[0]").Maximize()
            session.FindById("wnd[0]/usr/ctxtLIKP-VBELN").Text = remessa
            session.FindById("wnd[0]").SendVKey(0)
            time.sleep(1)

            # Acessa menu: Cabeçalho > Datas
            session.FindById("wnd[0]/mbar/menu[2]/menu[1]/menu[11]").Select()
            time.sleep(1)

            # Tenta localizar o campo
            field_path = "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\11/" + \
                         "ssubSUBSCREEN_BODY:SAPMV50A:2122/" + \
                         "subTSEG_STD:SAPLTSED:0100/tblSAPLTSEDTC_TSEG_STD/" + \
                         "ctxtITSEGDIAE-TIME_TST04[10,0]"

            try:
                campo = session.FindById(field_path)
            except:
                print(f"[INFO] Linha {idx}: campo não encontrado. Tentando adicionar linha com evento.")
                try:
                    # Clica no botão "+"
                    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\11/" +
                                     "ssubSUBSCREEN_BODY:SAPMV50A:2122/" +
                                     "subTSEG_STD:SAPLTSED:0100/btnAPPEND").Press()
                    time.sleep(1)

                    session.FindById("wnd[1]/usr/tblSAPLSTREVENTTC_EVENT_LIST/txtLSTREVENT-TYPE_ALIAS[0,0]").SetFocus()
                    session.FindById("wnd[1]/usr/tblSAPLSTREVENTTC_EVENT_LIST/txtLSTREVENT-TYPE_ALIAS[0,0]").caretPosition = 0
                    session.FindById("wnd[1]").SendVKey(0)
                    time.sleep(1)

                    campo = session.FindById(field_path)
                except Exception as add_err:
                    msg = f"[ERRO] Linha {idx} | Remessa {remessa} | Falha ao adicionar linha: {add_err}"
                    print(msg)
                    log.write(msg + "\n")
                    continue

            # Preenche o campo com a data da planilha
            campo.SetFocus()
            campo.Text = data_entrega
            campo.CaretPosition = 2

            # Salva e volta
            session.FindById("wnd[0]/tbar[0]/btn[11]").Press()  # Gravar
            time.sleep(1)
            session.FindById("wnd[0]/tbar[0]/btn[15]").Press()  # Voltar
            time.sleep(1)

            log.write(f"[OK] Linha {idx} | Remessa {remessa} atualizada com sucesso.\n")

        except Exception as e:
            msg = f"[ERRO] Linha {idx} | Remessa {remessa} | Erro geral: {e}"
            print(msg)
            log.write(msg + "\n")
            continue

    log.write("\nProcessamento finalizado.\n")

print("\nTodas as remessas foram processadas.")
print(f"Log salvo em: {log_path}")
input("Pressione Enter para sair...")
