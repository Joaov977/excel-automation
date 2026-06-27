import tkinter as tk

janela = tk.Tk()

janela.title("Sistema de Relatório de Vendas")
largura = 500
altura = 350

largura_tela = janela.winfo_screenwidth()
altura_tela = janela.winfo_screenheight()

x = (largura_tela // 2) - (largura // 2)
y = (altura_tela // 2) - (altura // 2)

janela.geometry(f"{largura}x{altura}+{x}+{y}")

titulo = tk.Label(
    janela,
    text="Sistema de Relatório de Vendas",
    font=("Arial", 18, "bold")
)

titulo.pack(pady=30)

botao_relatorio = tk.Button(
    janela,
    text="📊 Gerar Relatório",
    font=("Arial", 12),
    width=25,
    height=2
)

botao_relatorio.pack(pady=10)

botao_sair = tk.Button(
    janela,
    text="❌ Sair",
    font=("Arial", 12),
    width=25,
    height=2,
    command=janela.destroy
)

botao_sair.pack(pady=10)

janela.mainloop()