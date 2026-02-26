from tkinter import ttk
import tkinter as tk
from tkinter import messagebox, filedialog, simpledialog
import pandas as pd
import openpyxl
import os
import shutil
from datetime import datetime


# =========================
# CONFIGURAÇÕES
# =========================

SHEET_NAME = "GERAL NUMERADA"

SINONIMOS_COLUNAS = {
    'NOME': ['NOME', 'NOME COMPLETO', 'ATLETA'],
    'DOCUMENTO': ['DOCUMENTO', 'DOC', 'CPF', 'RG'],
    'Nº': ['NUMERO', 'Nº', 'PEITO'],
    'NASCIMENTO': ['DATA DE NASCIMENTO', 'NASCIMENTO', 'DATA NASC'],
    'SEXO': ['SEXO', 'MASC/FEM', 'GÊNERO'],
    'MODALIDADE': ['MODALIDADE', 'PROVA'],
    'KIT': ['TIPO DO KIT', 'KIT'],
    'TAMANHO': ['TAMANHO CAMISA', 'TAMANHO', 'CAMISA'],
    'EQUIPE': ['EQUIPE ', 'EQUIPE', 'CLUBE', 'TIME'],
    'MEIA': ['MEIA', 'MEIAS'],
    'PCD': ['PCD', 'DEFICIENTE'],
    'RETIRADO': ['RETIRADO', 'STATUS', 'ENTREGUE']
}


# =========================
# VARIÁVEIS GLOBAIS
# =========================

arquivo_atual = ""
df_dados = None
index_selecionado = None


# =========================
# UTIL
# =========================

def normalizar_colunas(df):
    df.columns = df.columns.str.strip().str.upper()
    for nome_oficial, variacoes in SINONIMOS_COLUNAS.items():
        for col in df.columns:
            if col in variacoes:
                df.rename(columns={col: nome_oficial}, inplace=True)
                break


# =========================
# IMPORTAÇÃO
# =========================

def selecionar_arquivo():
    global arquivo_atual, df_dados

    caminho = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if not caminho:
        return

    try:
        arquivo_atual = caminho
        df_temp = pd.read_excel(caminho, sheet_name=SHEET_NAME)

        normalizar_colunas(df_temp)

        for col in SINONIMOS_COLUNAS.keys():
            if col not in df_temp.columns:
                df_temp[col] = ""

        df_temp = df_temp.fillna("")
        df_temp["EXCEL_ROW"] = df_temp.index + 2

        if "Nº" in df_temp.columns:
            df_temp["Nº"] = (
                df_temp["Nº"]
                .astype(str)
                .str.replace(".0", "", regex=False)
                .str.strip()
            )

        df_dados = df_temp

        pasta_backup = os.path.join(os.path.dirname(caminho), "backups")
        os.makedirs(pasta_backup, exist_ok=True)
        shutil.copy(caminho,
                    os.path.join(pasta_backup,
                                 f"backup_{datetime.now().strftime('%H%M%S')}.xlsx"))

        lbl_arquivo.config(text=f"📁 {os.path.basename(caminho)}")
        entry_pesquisa.config(state="normal")
        atualizar_estatisticas()

        messagebox.showinfo("Success", "Spreadsheet loaded successfully!")

    except Exception as e:
        messagebox.showerror("Error", str(e))


# =========================
# EXIBIÇÃO
# =========================

def exibir_dados(atleta):
    nasc = atleta["NASCIMENTO"]

    if pd.notna(nasc):
        try:
            nasc = pd.to_datetime(nasc).strftime("%d/%m/%Y")
        except:
            nasc = str(nasc)

    retirado_valor = str(atleta.get("RETIRADO", "")).strip()

    if retirado_valor.startswith("ALTERADO"):
        status = f"Transferred → {retirado_valor}"
        entregue = True
    elif retirado_valor != "":
        status = f"Delivered on {retirado_valor}"
        entregue = True
    else:
        status = "Pending"
        entregue = False

    btn_confirmar.config(state="disabled" if entregue else "normal")

    texto = (
        f"BIB: {atleta.get('Nº','')}\n"
        f"NAME: {atleta.get('NOME','')}\n"
        f"BIRTHDATE: {nasc}\n"
        f"GENDER: {atleta.get('SEXO','')}\n\n"
        f"TEAM: {atleta.get('EQUIPE','')}\n"
        f"KIT: {atleta.get('KIT','')}\n"
        f"SIZE: {atleta.get('TAMANHO','')}\n"
        f"EVENT: {atleta.get('MODALIDADE','')}\n\n"
        f"STATUS: {status}"
    )

    lbl_dados.config(text=texto)


# =========================
# CONFIRMAR ENTREGA
# =========================

def confirmar_entrega():
    global index_selecionado

    if index_selecionado is None:
        return

    agora = datetime.now().strftime("%d/%m/%Y %H:%M")

    try:
        wb = openpyxl.load_workbook(arquivo_atual)
        ws = wb[SHEET_NAME]

        linha = int(df_dados.at[index_selecionado, "EXCEL_ROW"])

        col_retirado = None
        cabecalhos = {str(ws.cell(row=1, column=c).value).strip().upper(): c
                      for c in range(1, ws.max_column + 1)}

        for var in SINONIMOS_COLUNAS["RETIRADO"]:
            if var in cabecalhos:
                col_retirado = cabecalhos[var]
                break

        if not col_retirado:
            col_retirado = ws.max_column + 1
            ws.cell(row=1, column=col_retirado).value = "RETIRADO"

        ws.cell(row=linha, column=col_retirado).value = agora
        wb.save(arquivo_atual)

        df_dados.at[index_selecionado, "RETIRADO"] = agora
        atualizar_estatisticas()

        messagebox.showinfo("Success", "Delivery confirmed!")
        entry_pesquisa.delete(0, tk.END)
        lbl_dados.config(text="Waiting for search...")

    except Exception as e:
        messagebox.showerror("Error", str(e))


# =========================
# ESTATÍSTICAS
# =========================

def atualizar_estatisticas():
    if df_dados is None:
        return

    total = len(df_dados[df_dados["NOME"].astype(str).str.strip() != ""])
    entregues = df_dados[
        (df_dados["RETIRADO"].astype(str).str.strip() != "") &
        (df_dados["NOME"].astype(str).str.strip() != "")
    ].shape[0]

    lbl_estatisticas.config(
        text=f"Total: {total} | Delivered: {entregues} | Pending: {total - entregues}"
    )


# =========================
# INTERFACE
# =========================

root = tk.Tk()
root.title("Event Kit Manager")
root.geometry("900x650")

main_frame = tk.Frame(root)
main_frame.pack(expand=True, fill="both", padx=40, pady=20)

tk.Label(main_frame, text="Event Kit Manager",
         font=("Arial", 26, "bold")).pack(pady=15)

lbl_estatisticas = tk.Label(main_frame, text="Waiting for spreadsheet...")
lbl_estatisticas.pack()

tk.Button(main_frame,
          text="Import Spreadsheet",
          command=selecionar_arquivo).pack(pady=10)

lbl_arquivo = tk.Label(main_frame,
                       text="No file selected",
                       fg="gray")
lbl_arquivo.pack()

entry_pesquisa = tk.Entry(main_frame,
                          font=("Arial", 16),
                          state="disabled")
entry_pesquisa.pack(fill="x", pady=10)

lbl_dados = tk.Label(main_frame,
                     text="Waiting for search...",
                     justify="left",
                     font=("Arial", 12))
lbl_dados.pack(pady=10)

btn_confirmar = tk.Button(main_frame,
                          text="Confirm Delivery",
                          command=confirmar_entrega,
                          state="disabled")
btn_confirmar.pack(pady=20)

root.mainloop()
