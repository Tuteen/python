"""Codigo do sistema de estoque que armazena os registros em planilhas e tambem um historico"""
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import shutil
import os
import sys

# Definindo as colunas do DataFrame histórico
historico_df = pd.DataFrame(columns=['Item', 'Operação', 'Quantidade', 'Valor Pós Operação', 'Data/Hora', 'Usuario'])
caminho_pasta = r"D:\estoque"

# Definindo usuarios para login.
usuarios_permitidos = {
    "cassio": "ca2001",
    "giovanni": "krebi"}

usuario_atual = None

def carregar_dados():
    """Tentar carregar dados existentes da planilha"""
    global df
    caminho_completo = f"{caminho_pasta}/estoque.xlsx"
    try:
        df = pd.read_excel(caminho_completo)
        # Ordena o DataFrame em ordem alfabética baseado na coluna 'Item'
        return df.sort_values(by='Item')
    except FileNotFoundError:
        # Retorna um DataFrame vazio se a planilha não existir
        return pd.DataFrame(columns=['Item', 'Quantidade', 'Descrição', 'Estado'])
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao carregar dados: {e}")
        return pd.DataFrame(columns=['Item', 'Quantidade', 'Descrição', 'Estado'])
df = carregar_dados()

def carregar_historico():
    """carrega o historico existente"""
    caminho_historico = f"{caminho_pasta}/historico.xlsx"
    try:
        return pd.read_excel(caminho_historico)
    except FileNotFoundError:
        return pd.DataFrame(columns=['Item', 'Operação', 'Quantidade', 'Valor Pós Operação', 'Data/Hora', 'Usuario'])
historico_df = carregar_historico()

def realizar_backup(arquivo_dados):
    pasta_backup = f"{caminho_pasta}"
    if not os.path.exists(pasta_backup):
        os.makedirs(pasta_backup)
    timestamp = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
    arquivo_backup = f'{pasta_backup}/backup_{timestamp}.xlsx'
    shutil.copy(arquivo_dados, arquivo_backup)
    print(f'Backup realizado com sucesso: {arquivo_backup}')


def atualizar_tabela():
    """Função para atualizar a tabela"""
    global df
        # Adiciona uma coluna temporária com os nomes dos itens em minúsculas
    df['Item_temp'] = df['Item'].str.lower()
    # Ordena o DataFrame usando a coluna temporária
    df = df.sort_values(by='Item_temp')
    # Remove a coluna temporária após a ordenação
    df.drop('Item_temp', axis=1, inplace=True)
    for i in tree.get_children():
        tree.delete(i)
    for _, row in df.iterrows():
        tree.insert("", tk.END, values=list(row))

def pesquisar_item():
    '''Função que pesquisa um item dentro da planilha'''
    termo_pesquisa = pesquisa_entry.get().lower()
    for i in tree.get_children():
        tree.delete(i)
    for _, row in df.iterrows():
        if termo_pesquisa in row['Item'].lower():
            tree.insert("", tk.END, values=list(row))

# Função para adicionar um item
def adicionar_item():
    """
    Adiciona um novo item ao estoque.

    Esta função lê as entradas do usuário para o nome do item, quantidade,
    descrição e estado, adiciona essas informações ao DataFrame do estoque,
    e atualiza o arquivo do Excel correspondente. Também registra a ação no histórico.
    """
    global df, historico_df
    caminho_completo = f"{caminho_pasta}/estoque.xlsx"

    # Validação de campos vazios
    if not item_entry.get() or not quantidade_entry.get() or not descricao_entry.get():
        messagebox.showerror("Erro", "Todos os campos devem ser preenchidos.")
        return

    # Validação de quantidade (formato de número inteiro)
    try:
        quantidade = int(quantidade_entry.get())
        if quantidade < 0:
            raise ValueError("A quantidade não pode ser negativa.")
    except ValueError:
        messagebox.showerror("Erro", "Quantidade deve ser um número inteiro não negativo.")
        return

    novo_item = pd.DataFrame([{
        'Item': item_entry.get(), 
        'Quantidade': quantidade,
        'Descrição': descricao_entry.get(), 
        'Estado': estado_var.get()
        }])
    nome_item=item_entry.get() 
    df = pd.concat([df, novo_item], ignore_index=True)

    try:
        df.to_excel(caminho_completo, index=False)
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao salvar no Excel: {e}")
        return

    messagebox.showinfo("Informação", "Item adicionado com sucesso!")
    atualizar_tabela()
    limpar_campos()
    ##realizar_backup(caminho_completo)


    # Garantindo que os nomes das colunas do novo registro sejam consistentes com historico_df
    novo_registro_historico = pd.DataFrame([{
        'Item': nome_item,
        'Operação': 'Adicionar Item',
        'Quantidade': quantidade,
        'Valor Pós Operação': quantidade,
        'Data/Hora': datetime.now(),
        'Usuario': usuario_atual
    }])
    historico_df = pd.concat([historico_df, novo_registro_historico], ignore_index=True)

    # Salvar histórico
    historico_df.to_excel(f"{caminho_pasta}/historico.xlsx", index=False)

def limpar_campos():
    '''Função para limpar campos de entrada'''
    item_entry.delete(0, tk.END)
    quantidade_entry.delete(0, tk.END)
    descricao_entry.delete(0, tk.END)
    estado_var.set("Novo")

def editar_quantidade():
    """Função que edita a quantidade de um item do sistema
    faz a coleta do registro na planilha e registra a nova
    quantidade e tambem registra a alteração no historico 
    de alterações"""
    global df, historico_df
    caminho_completo = f"{caminho_pasta}/estoque.xlsx"
    selected_item = tree.focus()
    values = tree.item(selected_item, 'values')

    if values:
        def confirmar():
            global df, historico_df
            try:
                nova_quantidade = int(quantidade_entrada.get())
                operacao = operacao_var.get()
                item = values[0]
                quantidade_atual = df.loc[df['Item'] == item, 'Quantidade'].item()

                if operacao == 'Entrada':
                    quantidade_final = quantidade_atual + nova_quantidade
                else:  # Saída
                    quantidade_final = max(0, quantidade_atual - nova_quantidade)

                df.loc[df['Item'] == item, 'Quantidade'] = quantidade_final
                try:
                    df.to_excel(caminho_completo, index=False)
                    historico_df = pd.concat([historico_df, pd.DataFrame([{
                        'Item': item,
                        'Operação': operacao,
                        'Quantidade': nova_quantidade,
                        'Valor Pós Operação': quantidade_final,
                        'Data/Hora': datetime.now(),
                        'Usuario': usuario_atual
                    }])], ignore_index=True)
                    historico_df.to_excel(f"{caminho_pasta}/historico.xlsx", index=False)
                except Exception as e:
                    messagebox.showerror("Erro", f"Erro ao salvar no Excel: {e}")
                    return

                editar_janela.destroy()
                atualizar_tabela()
            except ValueError:
                messagebox.showerror("Erro", "Quantidade deve ser um número inteiro.")

        editar_janela = tk.Toplevel(root)
        editar_janela.title(f"Editar Quantidade - {values[0]}")

        tk.Label(editar_janela, text="Quantidade").grid(row=0, column=0)
        quantidade_entrada = tk.Entry(editar_janela)
        quantidade_entrada.grid(row=0, column=1)

        operacao_var = tk.StringVar(value="Entrada")
        tk.Radiobutton(editar_janela, text="Entrada", variable=operacao_var, value="Entrada").grid(row=1, column=0)
        tk.Radiobutton(editar_janela, text="Saída", variable=operacao_var, value="Saída").grid(row=1, column=1)

        tk.Button(editar_janela, text="Confirmar", command=confirmar).grid(row=2, column=0, columnspan=2)

 

def exibir_historico():
    '''Função para exibir o historico'''
    historico_janela = tk.Toplevel(root)
    historico_janela.title("Histórico de Ações")

    # Frame para Treeview e Scrollbar
    tree_frame = tk.Frame(historico_janela)
    tree_frame.pack(expand=True, fill='both')

    # Definindo as colunas do histórico
    colunas_historico = ['Item', 'Operação', 'Quantidade', 'Valor Pós Operação', 'Data/Hora', 'Usuario']
    historico_tree = ttk.Treeview(tree_frame, columns=colunas_historico, show='headings')
    for col in colunas_historico:
        historico_tree.heading(col, text=col)
        historico_tree.column(col, anchor='center')
    historico_tree.pack(side='left', expand=True, fill='both')

    # Adicionar scrollbar
    scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=historico_tree.yview)
    scrollbar.pack(side='right', fill='y')
    historico_tree.configure(yscrollcommand=scrollbar.set)

    # Inserindo os dados do histórico na árvore
    for _, row in historico_df.iterrows():
        # Formatando 'Valor Pós Operação' para remover o '.0' se for um número
        valor_pos_operacao = row['Valor Pós Operação']
        if pd.notna(valor_pos_operacao) and isinstance(valor_pos_operacao, float) and valor_pos_operacao.is_integer():
            valor_pos_operacao = int(valor_pos_operacao)
            
            quantidade = row['Quantidade']
        if pd.notna(quantidade) and isinstance(quantidade, float) and quantidade.is_integer():
            quantidade = int(quantidade)
        
        # Formatando a data/hora
        data_hora_formatada = row['Data/Hora'].strftime("%d/%m/%y %H:%M:%S") if pd.notna(row['Data/Hora']) else "N/A"

        historico_tree.insert("", tk.END, values=[
        row['Item'],
        row['Operação'],
        quantidade,
        valor_pos_operacao,
        data_hora_formatada,
        row['Usuario']
    ])
        
    historico_tree.pack(expand=True, fill='both')

    
    pesquisa_historico_entry = tk.Entry(historico_janela)
    pesquisa_historico_entry.pack()
    pesquisa_historico_entry.bind("<Return>", lambda event: pesquisar_historico(pesquisa_historico_entry, historico_tree))
    tk.Button(historico_janela, text="Pesquisar", command=lambda: pesquisar_historico(pesquisa_historico_entry, historico_tree)).pack()
    

def pesquisar_historico(entry, treeview):
    '''pesquisa um item no historico'''
    termo_pesquisa = entry.get().lower()
    for i in treeview.get_children():
        treeview.delete(i)
    for _, row in historico_df.iterrows():
        item = str(row['Item']).lower()
        if termo_pesquisa in item:
            # Formatando 'Valor Pós Operação' para remover '.0' se for um número inteiro
            valor_pos_operacao = row['Valor Pós Operação']
            if pd.notna(valor_pos_operacao) and valor_pos_operacao.is_integer():
                valor_pos_operacao = int(valor_pos_operacao)
                quantidade = row['Quantidade']
                if pd.notna(quantidade) and isinstance(quantidade, float) and quantidade.is_integer():
                    Quantidade = int(quantidade)

            treeview.insert("", tk.END, values=[
                row['Item'],
                row['Operação'],
                Quantidade,
                valor_pos_operacao,
                row['Data/Hora'].strftime("%d/%m/%Y %H:%M:%S") if pd.notna(row['Data/Hora']) else "N/A",
                row['Usuario']
            ])


# Função para deletar um item
def deletar_item():
    '''Deleta um item no historico'''
    global df, historico_df
    caminho_completo = f"{caminho_pasta}/estoque.xlsx"
    selected_item = tree.focus()
    values = tree.item(selected_item, 'values')

    if values and messagebox.askyesno("Confirmação", "Tem certeza que deseja remover este item?"):
        item_a_remover = values[0]

        try:
            # Removendo o item do DataFrame
            df.drop(df[df['Item'] == item_a_remover].index, inplace=True)
            df.to_excel(caminho_completo, index=False)

            # Registrando a ação no histórico
            historico_df = pd.concat([historico_df, pd.DataFrame([{
                'Item': item_a_remover,
                'Operação': 'Remover Item',
                'Quantidade': '',  # Não aplicável na remoção
                'Valor Pós Operação': '',  # Não aplicável na remoção
                'Data/Hora': datetime.now(),
                'Usuario': usuario_atual
            }])], ignore_index=True)
            historico_df.to_excel(f"{caminho_pasta}/historico.xlsx", index=False)

        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao salvar no Excel: {e}")
            return

        atualizar_tabela()

def verificar_login(*args):
    global usuario_atual
    usuario = usuario_entry.get()
    senha = senha_entry.get()
    if usuarios_permitidos.get(usuario) == senha:
        messagebox.showinfo("Login", "Login bem sucedido!")
        usuario_atual = usuario
        login_janela.destroy()
        # Aqui você pode chamar a função que inicia a tela principal do seu aplicativo
    else:
        messagebox.showerror("Login", "Usuário ou senha incorretos!")

def encerrar_programa():
    """Função para encerrar o programa de forma segura."""
    login_janela.destroy()  # Destruir a janela de login
    sys.exit()  # Sair do programa


login_janela = tk.Tk()
login_janela.title("Login")

tk.Label(login_janela, text="Usuário:").pack()
usuario_entry = tk.Entry(login_janela)
usuario_entry.pack()

tk.Label(login_janela, text="Senha:").pack()
senha_entry = tk.Entry(login_janela, show="*")
senha_entry.pack()

tk.Button(login_janela, text="Login", command=verificar_login).pack()
senha_entry.bind("<Return>", verificar_login)
login_janela.protocol("WM_DELETE_WINDOW", encerrar_programa)

login_janela.mainloop()

# Criar a janela principal
# Criar a janela principal
root = tk.Tk()
root.title("Sistema de Estoque")

# Organizando com grid
frame = tk.Frame(root)
frame.pack(padx=10, pady=10)

tk.Label(frame, text="Nome do Item").grid(row=0, column=0, sticky="w")
item_entry = tk.Entry(frame)
item_entry.grid(row=0, column=1)

tk.Label(frame, text="Quantidade").grid(row=1, column=0, sticky="w")
quantidade_entry = tk.Entry(frame)
quantidade_entry.grid(row=1, column=1)

tk.Label(frame, text="Descrição").grid(row=2, column=0, sticky="w")
descricao_entry = tk.Entry(frame)
descricao_entry.grid(row=2, column=1)

tk.Label(frame, text="Estado").grid(row=3, column=0, sticky="w")
estado_var = tk.StringVar(value="Novo")
tk.OptionMenu(frame, estado_var, "Novo", "Usado").grid(row=3, column=1)

tk.Button(frame, text="Adicionar Item", command=adicionar_item).grid(row=4, column=1, columnspan=1)
item_entry.bind("<Return>", lambda event: adicionar_item())
quantidade_entry.bind("<Return>", lambda event: adicionar_item())
descricao_entry.bind("<Return>", lambda event: adicionar_item())

# Campo de pesquisa
pesquisa_entry = tk.Entry(frame)
pesquisa_entry.grid(row=5, column=1)
pesquisa_entry.bind("<Return>", lambda event: pesquisar_item())
tk.Button(frame, text="Pesquisar", command=pesquisar_item).grid(row=5, column=2)


# Tabela de itens
cols = list(df.columns)
tree = ttk.Treeview(root, columns=cols, show='headings')
for col in cols:
    tree.heading(col, text=col)
tree.pack(expand=True, fill='both')

# Botões de editar, remover e exibir historico
# Criando um frame para os botões
button_frame = tk.Frame(root)
button_frame.pack(pady=10)

# Definindo a largura desejada para os botões
Largura_Botao = 20

# Botão de Editar Quantidade
botao_editar = tk.Button(button_frame, text="Editar Quantidade", command=editar_quantidade)
botao_editar.grid(row=0, column=0, padx=10, pady=10, sticky='ew')
botao_editar.config(width=Largura_Botao)

# Botão de Remover Item
botao_remover = tk.Button(button_frame, text="Remover Item", command=deletar_item)
botao_remover.grid(row=1, column=0, padx=10, pady=10, sticky='ew')
botao_remover.config(width=Largura_Botao)

# Botão de Exibir Histórico
botao_historico = tk.Button(button_frame, text="Exibir Histórico", command=exibir_historico)
botao_historico.grid(row=2, column=0, padx=10, pady=10, sticky='ew')
botao_historico.config(width=Largura_Botao)

# Iniciar a interface gráfica
atualizar_tabela()
root.mainloop()
