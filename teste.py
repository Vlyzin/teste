import win32com.client
import time

try:
    SapGuiAuto = win32com.client.GetObject("SAPGUI")
    application = SapGuiAuto.GetScriptingEngine

    if application.Children.Count == 0:
        raise Exception("Nenhuma conexão SAP ativa foi encontrada.")

    connection = application.Children(0)

    if connection.Children.Count == 0:
        raise Exception("Nenhuma sessão SAP ativa foi encontrada.")

    session = connection.Children(0)

    # Manda o comando da transação
    session.findById("wnd[0]/tbar[0]/okcd").text = "j1b3n"
    session.findById("wnd[0]").sendVKey(0)

    print("✅ Comando J1B3N enviado com sucesso!")

except Exception as e:
    print("❌ Erro ao executar o script:")
    print(e)
