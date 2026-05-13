from openpyxl import Workbook

planilha = Workbook()

aba = planilha.active

aba["A1"] = "Produto"
aba["B1"] = "Preço"

produtos = [
    ["Mouse", 100],
    ["Teclado", 150],
    ["Monitor", 900]
]

linha = 2

for produto in produtos:
    aba[f"A{linha}"] = produto[0]
    aba[f"B{linha}"] = produto[1]

    linha += 1

planilha.save("vendas.xlsx")

print("Planilha criada")