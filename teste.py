import win32com.client
import time

def encontrar_sessao_sap():
    try:
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        application = SapGuiAuto.GetScriptingEngine

        for i in range(application.Children.Count):
            connection = application.Children(i)
            for j in range(connection.Children.Count):
                session = connection.Children(j)
                # Verifica se a sessão está pronta
                if session.Info.IsLowSpeedConnection == False:
                    print(f" Sessão encontrada: Conexão {i}, Sessão {j}")
                    return session
        print(" Nenhuma sessão SAP válida foi encontrada.")
        return None

    except Exception as e:
        print(" Erro ao buscar sessão SAP:")
        print(e)
        return None


def abrir_transacao(session, transacao):
    try:
        session.findById("wnd[0]/tbar[0]/okcd").text = transacao
        session.findById("wnd[0]").sendVKey(0)
        print(f" Transação '{transacao}' executada com sucesso!")
    except Exception as e:
        print(f" Erro ao abrir a transação '{transacao}':")
        print(e)


# -------- EXECUÇÃO -------- #
if __name__ == "__main__":
    sessao = encontrar_sessao_sap()
    if sessao:
        abrir_transacao(sessao, "j1b3n")
