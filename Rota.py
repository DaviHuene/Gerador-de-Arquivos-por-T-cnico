import pandas as pd
import os
import customtkinter as ctk
from tkinter import filedialog, messagebox

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

def selecionar_aba(e=None):
    aba_escolhida = opcao_abas.get()
    if not aba_escolhida:
        messagebox.showerror("Erro", "Nenhuma aba selecionada.")
        return

    try:
        df = pd.read_excel(caminho_arquivo_global, sheet_name=aba_escolhida)
        df.columns = df.columns.str.strip()

        colunas_queridas = ['Novo TEC', 'Endereço', 'Novo End', 'Intervalo de Tempo', 'Contrato']
        colunas_df_map = {col.lower(): col for col in df.columns}

        colunas_final = []
        for c in colunas_queridas:
            c_lower = c.lower()
            if c_lower in colunas_df_map:
                colunas_final.append(colunas_df_map[c_lower])
            else:
                messagebox.showerror("Erro", f"Coluna não encontrada: '{c}'")
                return

        df = df[colunas_final]
        df = df[df[colunas_final[0]].notna()]

        pasta_saida = os.path.join(os.path.dirname(caminho_arquivo_global), 'Arquivos_por_Tecnico')
        os.makedirs(pasta_saida, exist_ok=True)

        for tecnico, grupo in df.groupby(colunas_final[0]):
            nome_arquivo = f"{str(tecnico).strip().replace(' ', '_')}.txt"
            caminho_txt = os.path.join(pasta_saida, nome_arquivo)
            grupo.to_csv(caminho_txt, sep='\t', index=False)

        messagebox.showinfo("Sucesso", f"Arquivos salvos em:\n{pasta_saida}")

    except Exception as e:
        messagebox.showerror("Erro", str(e))


def processar_planilha():
    global caminho_arquivo_global
    caminho_arquivo = filedialog.askopenfilename(
        title="Selecione a planilha Excel",
        filetypes=[("Arquivos Excel", "*.xlsx *.xls")]
    )

    if not caminho_arquivo:
        return

    caminho_arquivo_global = caminho_arquivo
    try:
        excel_file = pd.ExcelFile(caminho_arquivo)
        abas = excel_file.sheet_names

        # Exibe o menu de seleção de aba
        opcao_abas.configure(values=abas)
        opcao_abas.set(abas[0])  # Define a primeira aba como padrão
        opcao_abas.pack(pady=(10, 5))
        botao_processar_aba.pack(pady=(5, 20))

    except Exception as e:
        messagebox.showerror("Erro", str(e))


# Interface
app = ctk.CTk()
app.title("Gerador de arquivos por Técnico")
app.geometry("500x350")
app.resizable(False, False)

frame = ctk.CTkFrame(app, corner_radius=15)
frame.pack(padx=30, pady=30, fill="both", expand=True)

label = ctk.CTkLabel(frame, text="Selecionar a planilha Excel", font=("Arial", 16))
label.pack(pady=(40, 20))

botao = ctk.CTkButton(
    frame,
    text="Selecionar Planilha",
    command=processar_planilha,
    width=200,
    height=40,
    corner_radius=20,
    font=("Arial", 14)
)
botao.pack()

opcao_abas = ctk.CTkOptionMenu(frame, values=[], width=250)
botao_processar_aba = ctk.CTkButton(
    frame,
    text="Processar Aba Selecionada",
    command=selecionar_aba,
    width=220,
    height=35,
    corner_radius=15
)

app.mainloop()
#by ben and davi....