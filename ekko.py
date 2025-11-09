import tkinter as tk
from tkinter import messagebox
import win32com.client
import time
import datetime
import os
import pandas as pd
import traceback
import logging

# --- Configura o logger para depuração ---
logging.basicConfig(
    filename="sap_automation.log",
    level=logging.DEBUG,
    format="%(asctime)s %(levelname)s %(message)s",
)

# --- Funções auxiliares ---
def safe_get_username():
    try:
        return os.getlogin()
    except Exception:
        return os.environ.get("USERNAME", "unknown")

def wait_for_id(session, element_id, timeout=10, poll=0.5):
    """Tenta encontrar um controle pelo id até timeout (retorna o objeto ou lança erro)."""
    start = time.time()
    while True:
        try:
            ctl = session.findById(element_id)
            return ctl
        except Exception:
            if time.time() - start > timeout:
                raise
            time.sleep(poll)

# --- Função principal de automação SAP ---
def executar_automacao_sap():
    try:
        usuario = safe_get_username()
        pasta = fr"C:\Users\{usuario}\OneDrive - Accenture\Desktop\junior"
        os.makedirs(pasta, exist_ok=True)

        data_atual = datetime.datetime.now().strftime("%d_%m_%y")
        nome_xls = f"EKKO_{data_atual}.XLS"
        caminho_xls = os.path.join(pasta, nome_xls)
        caminho_txt = os.path.splitext(caminho_xls)[0] + ".txt"

        # --- Conecta ao SAP ---
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        application = SapGuiAuto.GetScriptingEngine
        connection = application.Children(0)
        session = connection.Children(0)

        # --- Executa SE16N ---
        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/tbar[0]/okcd").text = "se16n"
        session.findById("wnd[0]/tbar[0]/btn[0]").press()
        session.findById("wnd[0]/usr/ctxtGD-TAB").text = "EKKO"
        session.findById("wnd[0]/tbar[1]/btn[8]").press()

        # --- Exporta resultado ---
        time.sleep(2)
        session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
        session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").selectContextMenuItem("&PC")

        # --- Seleciona formato XLS ---
        session.findById(
            "wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/"
            "sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[2,0]"
        ).select()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()

        # --- Define caminho e nome ---
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = pasta
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = nome_xls
        session.findById("wnd[1]/tbar[0]/btn[0]").press()

        # --- Confirma substituição se aparecer ---
        try:
            session.findById("wnd[1]/tbar[0]/btn[20]").press()
            session.findById("wnd[1]/tbar[0]/btn[0]").press()
        except:
            pass

        time.sleep(2)

        # --- Converte XLS → TXT ---
        df = pd.read_excel(caminho_xls, header=None)
        df.to_csv(caminho_txt, sep="\t", index=False, header=False)
        os.remove(caminho_xls)

        messagebox.showinfo("Sucesso", f"Arquivo salvo e convertido:\n{caminho_txt}")

    except Exception as e:
        logging.error("Erro na automação: " + str(e) + "\n" + traceback.format_exc())
        messagebox.showerror("Erro", f"Ocorreu um erro na automação:\n\n{e}")

# --- Interface gráfica (Tkinter) ---
def criar_interface():
    janela = tk.Tk()
    janela.title("Junior Dev - Automação SAP")
    janela.geometry("400x300")
    janela.configure(bg="#030303")  # fundo pastel claro

    # --- Título ---
    label_titulo = tk.Label(
        janela,
        text="Junior Dev",
        font=("Helvetica", 24, "bold"),
        fg="#122ee4",
        bg="#060606"
    )
    label_titulo.pack(pady=50)

    # --- Botões ---
    estilo_botao = {
        "font": ("Helvetica", 12, "bold"),
        "bg": "#CEB2B2",
        "fg": "#22223b",
        "activebackground": "#0f0e0f",
        "activeforeground": "#f2e9e4",
        "width": 15,
        "height": 1,
        "relief": "ridge",
        "bd": 3,
    }

    btn_iniciar = tk.Button(
        janela,
        text="Iniciar",
        command=executar_automacao_sap,
        **estilo_botao
    )
    btn_iniciar.pack(pady=10)

    btn_voltar = tk.Button(
        janela,
        text="Voltar",
        command=janela.destroy,
        **estilo_botao
    )
    btn_voltar.pack(pady=5)

    janela.mainloop()

# --- Executa ---
if __name__ == "__main__":
    criar_interface()
