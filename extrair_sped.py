import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os
import numpy as np

def selecionar_arquivos():
    global linhas_1100, linhas_1500
    linhas_1100 = []
    linhas_1500 = []

    arquivos_txt = filedialog.askopenfilenames(  # Permite múltipla seleção
        title="Selecione os arquivos .txt",
        filetypes=[("Arquivos de Texto", "*.txt"), ("Todos os arquivos", "*.*")]
    )

    if arquivos_txt:  # Se o usuário selecionou arquivos
        for arquivo_txt in arquivos_txt:
            linhas_1100_arquivo, linhas_1500_arquivo = processar_arquivo(arquivo_txt)
            linhas_1100.extend(linhas_1100_arquivo)
            linhas_1500.extend(linhas_1500_arquivo)
        
        messagebox.showinfo("Processamento Concluído", 
                            f"Linhas |1100| e |1500| extraídas de {len(arquivos_txt)} arquivos com sucesso!")

def processar_arquivo(arquivo_txt):
    linhas_1100 = []
    linhas_1500 = []

    # Adicionar o nome do arquivo como título
    nome_arquivo = os.path.basename(arquivo_txt)
    
    # Abrir o arquivo e buscar pelas linhas que começam com |1100| ou |1500|
    with open(arquivo_txt, 'r', encoding='latin1') as f:
        # Adicionar título de separação para identificar de qual arquivo os dados são
            
        linhas_1100.append([f"--- {nome_arquivo} ---"])
        linhas_1100.append("")  # Adicionar uma linha em branco
        linhas_1500.append([f"--- {nome_arquivo} ---"])
        linhas_1500.append("")  # Adicionar uma linha em branco
        
        for linha in f:
            if linha.startswith("|1100|"):
                linhas_1100.append(linha.strip().split("|")[1:-1])  # Remove o primeiro e último pipes
            elif linha.startswith("|1500|"):
                linhas_1500.append(linha.strip().split("|")[1:-1])  # Remove o primeiro e último pipes
                
        # Adicionar uma linha em branco após os dados de cada arquivo
        linhas_1100.append("")  # Espaçamento após o final do arquivo 1100
        linhas_1500.append("")  # Espaçamento após o final do arquivo 1500

    return linhas_1100, linhas_1500

def salvar_arquivo():
    if not linhas_1100 and not linhas_1500:
        messagebox.showwarning("Aviso", "Nenhum dado disponível para salvar. Selecione um arquivo primeiro.")
        return

    # Permitir salvar em Excel ou TXT
    arquivo_saida = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Arquivos Excel", "*.xlsx"), ("Arquivos de Texto", "*.txt")],
        title="Salvar Arquivo"
    )

    if arquivo_saida:
        # Criar DataFrames para |1100| e |1500|
        df_1100 = pd.DataFrame(linhas_1100)
        df_1500 = pd.DataFrame(linhas_1500)
        
        # Substituir valores vazios ou espaços por NaN
        df_1100.replace(r'^\s*$', np.nan, regex=True, inplace=True)
        df_1500.replace(r'^\s*$', np.nan, regex=True, inplace=True)
        
        # Remover colunas vazias
        if not df_1100.empty:
            df_1100 = df_1100.dropna(axis=1, how='all')
        if not df_1500.empty:
            df_1500 = df_1500.dropna(axis=1, how='all')

        # Verificar a extensão do arquivo e salvar no formato apropriado
        if arquivo_saida.endswith(".xlsx"):
            with pd.ExcelWriter(arquivo_saida, engine="openpyxl") as writer:
                if not df_1100.empty:
                    df_1100.to_excel(writer, sheet_name="1100", index=False, header=False)
                if not df_1500.empty:
                    df_1500.to_excel(writer, sheet_name="1500", index=False, header=False)
            messagebox.showinfo("Sucesso", f"Arquivo Excel salvo em:\n{os.path.abspath(arquivo_saida)}")
        
        elif arquivo_saida.endswith(".txt"):
            with open(arquivo_saida, "w", encoding="utf-8") as f:
                if not df_1100.empty:
                    # Escrever as linhas no arquivo TXT com separação
                    for linha in df_1100.values:
                        f.write("|" + "|".join(map(str, linha)) + "|\n")
                if not df_1500.empty:
                    f.write("\n")
                    for linha in df_1500.values:
                        f.write("|" + "|".join(map(str, linha)) + "|\n")
            messagebox.showinfo("Sucesso", f"Arquivo TXT salvo em:\n{os.path.abspath(arquivo_saida)}")

# Configuração inicial da interface gráfica
root = tk.Tk()
root.title("Extrator de Linhas |1100| e |1500| para Excel/TXT")
root.geometry("400x200")
root.resizable(False, False)

# Variáveis globais
linhas_1100 = []
linhas_1500 = []

btn_selecionar = tk.Button(root, text="Selecionar Arquivo", command=selecionar_arquivos, font=("Arial", 10))
btn_selecionar.pack(pady=20)

btn_salvar = tk.Button(root, text="Salvar Arquivo", command=salvar_arquivo, font=("Arial", 10))
btn_salvar.pack(pady=20)

btn_sair = tk.Button(root, text="Sair", command=root.quit, font=("Arial", 10))
btn_sair.pack(pady=20)

root.mainloop()
