import tkinter as tk
from tkinter import filedialog, messagebox
from threading import Thread
from automacao import reenvio

def submeter():
    try:
        mensagem = entry_mensagem.get("1.0", tk.END).strip()
        caminho_arquivo = entry_arquivo.get()
        nome_planilha = entry_nome_planilha.get()
        nome_pagina = entry_nome_pagina.get()
        caminho_foto = entry_foto.get()

        # Mostrar os dados em uma caixa de mensagem (popup)
        messagebox.showinfo("Dados Submetidos", f"Mensagens: {mensagem}\nNome da Planilha: {nome_planilha}\nNome Página: {nome_pagina}\nCaminho do Arquivo: {caminho_arquivo}\nCaminho da foto: {caminho_foto}")

        # Chamar a automação em um thread separada
        thread = Thread(target=reenvio, args=(mensagem, caminho_arquivo, nome_planilha, nome_pagina, caminho_foto))
        thread.start()

    except Exception as e:
        print(f"Erro ao submeter: {e}")
        messagebox.showerror("Erro", f"Ocorreu um erro ao submeter: {e}")

# Função para selecionar o arquivo de planilha
def selecionar_arquivo():
    caminho_arquivo = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx")])
    entry_arquivo.delete(0, tk.END)
    entry_arquivo.insert(0, caminho_arquivo)

# Função para selecionar a foto
def selecionar_foto():
    caminho_foto = filedialog.askopenfilename(filetypes=[("Foto", "*.png")])
    entry_foto.delete(0, tk.END)
    entry_foto.insert(0, caminho_foto)

# Configuração da janela
root = tk.Tk()
root.title("Tela de Automacao")
root.geometry("600x400")
root.configure(bg="white")

# Layout
tk.Label(root, text="Mensagem de reenvio", bg="white").pack(pady=(10, 0))
entry_mensagem = tk.Text(root, width=50, height=10, wrap='word')
entry_mensagem.pack(pady=(0, 10))

# Label para exibir a mensagem capturada (salvo na variável "mensagem")
mensagem = tk.Label(root, text="", bg="white")
mensagem.pack(pady=(10, 10))


tk.Label(root, text="Nome da Planilha", bg="white").pack(pady=(10, 0))
entry_nome_planilha = tk.Entry(root, width=50)
entry_nome_planilha.pack(pady=(0, 10))

tk.Label(root, text="Nome Página", bg="white").pack(pady=(10, 0))
entry_nome_pagina = tk.Entry(root, width=50)
entry_nome_pagina.pack(pady=(0, 10))

tk.Label(root, text="Submeta o arquivo", bg="white").pack(pady=(10, 0))
entry_arquivo = tk.Entry(root, width=50)
entry_arquivo.pack(pady=(0, 5))
btn_browse = tk.Button(root, text="Selecionar Arquivo", command=selecionar_arquivo)
btn_browse.pack(pady=(0, 10))

tk.Label(root, text="Imagem", bg="white").pack(pady=(10, 0))
entry_foto = tk.Entry(root, width=50)
entry_foto.pack(pady=(0, 10))
btn_browse = tk.Button(root, text="Selecionar Foto", command=selecionar_foto)
btn_browse.pack(pady=(0, 10))

tk.Button(root, text="Submeter", command=submeter).pack(pady=(10, 10))


# Iniciar o loop principal
root.mainloop()
