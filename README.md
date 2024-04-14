import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from openpyxl import Workbook, load_workbook
import os
import random

def adicionar_funcionario():
    # Função para adicionar um novo funcionário
    # Obter os valores das caixas de entrada de texto
    nome = entry_nome.get()
    idade = entry_idade.get()
    cargo = entry_cargo.get()
    salario = entry_salario.get()

    # Verificar se todos os campos estão preenchidos
    if nome == "" or idade == "" or cargo == "" or salario == "":
        messagebox.showerror("Erro", "Por favor, preencha todos os campos.")
        return

    # Obter o caminho completo para o arquivo Excel
    diretorio_atual = os.path.dirname(os.path.abspath(_file_))
    caminho_arquivo = os.path.join(diretorio_atual, 'funcionarios.xlsx')

    # Adicionar o funcionário à base de dados
    wb = load_workbook(caminho_arquivo)  # Carregar o arquivo Excel existente
    ws = wb.active

    # Encontrar a próxima linha vazia na planilha
    linha_vazia = ws.max_row + 1

    # Adicionar os dados do funcionário à planilha
    ws.cell(row=linha_vazia, column=1, value=nome)
    ws.cell(row=linha_vazia, column=2, value=idade)
    ws.cell(row=linha_vazia, column=3, value=cargo)
    ws.cell(row=linha_vazia, column=4, value=salario)

    # Salvar as alterações no arquivo Excel
    wb.save(caminho_arquivo)

    print("Adicionando funcionário:", nome, idade, cargo, salario)

    # Limpar os campos de entrada de texto
    entry_nome.delete(0, tk.END)
    entry_idade.delete(0, tk.END)
    entry_cargo.delete(0, tk.END)
    entry_salario.delete(0, tk.END)

def limpar_campos():
    # Função para limpar os campos de entrada de texto
    entry_nome.delete(0, tk.END)
    entry_idade.delete(0, tk.END)
    entry_cargo.delete(0, tk.END)
    entry_salario.delete(0, tk.END)

root = tk.Tk()
root.title("Cadastro de Funcionários")
root.geometry("800x500")  # Definindo o tamanho da janela

# Título da janela
titulo = ttk.Label(root, text="Cadastro de Funcionários", font=('Verdana', 16, 'bold'))
titulo.pack(pady=10)

# Definindo cores para a interface
root.configure(bg='#dbe6fd')
frame_adicionar = ttk.Frame(root, padding=10, style='TFrame', relief='raised')
frame_adicionar.pack(expand=True, fill=tk.BOTH)  # Expande para ocupar todo o espaço da janela

# Labels e entradas de texto para os campos
campos = ["Nome", "Idade", "Cargo", "Salário", "Departamento", "Telefone"]
random.shuffle(campos)  # Embaralha a ordem dos campos aleatoriamente
for i, campo in enumerate(campos):
    label = ttk.Label(frame_adicionar, text=campo+":", style='TLabel', font=('Verdana', 14))  
    label.grid(row=i, column=0, sticky="e", pady=random.randint(5, 15), padx=10)  
    entry = ttk.Entry(frame_adicionar, font=('Verdana', 14))  
    entry.grid(row=i, column=1, sticky="we", pady=random.randint(5, 15), padx=10) 
    globals()["entry_" + campo.lower()] = entry

# Botão para adicionar funcionário
btn_adicionar = ttk.Button(frame_adicionar, text="Adicionar Funcionário", command=adicionar_funcionario, style='TButton')
btn_adicionar.grid(row=len(campos), columnspan=2, pady=10)

# Botão para limpar os campos
btn_limpar = ttk.Button(frame_adicionar, text="Limpar Campos", command=limpar_campos, style='TButton')
btn_limpar.grid(row=len(campos)+1, columnspan=2, pady=10)

# Estilo para a interface
style = ttk.Style()
style.configure('TFrame', background='#c6e2ff')
style.configure('TLabel', background='#dbe6fd', foreground='blue', font=('Verdana', 12))
style.configure('TEntry', fieldbackground='lightgreen', font=('Verdana', 12))
style.configure('TButton', foreground='blue', background='lightgreen', font=('Verdana', 12))

root.mainloop()
