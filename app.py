from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from datetime import datetime
from openpyxl.chart import BarChart, Reference

planilha = Workbook()

aba = planilha.active

aba.title = "Relatório de Vendas"

aba["A7"] = f"Relatório gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M')}"

cabecalhos = ["Produto", "Preço", "Quantidade", "Total"]

for coluna, titulo in enumerate(cabecalhos, start=1):
    celula = aba.cell(row=1, column=coluna)

    celula.value = titulo

    celula.font = Font(
        bold=True,
        color="FFFFFF"
    )

    celula.fill = PatternFill(
        start_color="1F4E78",
        end_color="1F4E78",
        fill_type="solid"
    )

    celula.alignment = Alignment(
        horizontal="center"
    )

aba.column_dimensions["A"].width = 25
aba.column_dimensions["B"].width = 15
aba.column_dimensions["C"].width = 15
aba.column_dimensions["D"].width = 15

produtos = [
    ["Mouse", 100, 2],
    ["Teclado", 150, 1],
    ["Monitor", 900, 3],
    ["Notebook", 3500, 1]
]

borda = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin")
)

linha = 2

for produto in produtos:
    nome = produto[0]
    preco = produto[1]
    quantidade = produto[2]

    total = preco * quantidade

    aba[f"A{linha}"] = nome
    aba[f"B{linha}"] = preco
    aba[f"C{linha}"] = quantidade
    aba[f"D{linha}"] = total

    if total >= 1000:
        aba [f"D{linha}"].fill = PatternFill(
            start_color="C6EFCE",
            end_color="C6EFCE",
            fill_type="solid"
        )

    aba[f"B{linha}"].number_format = 'R$ #,##0.00'
    aba[f"D{linha}"].number_format = 'R$ #,##0.00'

    for coluna in ["A", "B", "C", "D"]:
        aba[f"{coluna}{linha}"].alignment = Alignment(horizontal="center")
        aba[f"{coluna}{linha}"].border = borda

    linha += 1

aba[f"C{linha}"] = "Total Geral"

aba[f"C{linha}"].font = Font(bold=True)
aba[f"D{linha}"].font = Font(bold=True)


aba[f"D{linha}"] = f"=SUM(D2:D{linha-1})"


aba[f"D{linha}"].number_format = 'R$ #,##0.00'

aba[f"C{linha}"].fill = PatternFill(
    start_color="D9EAD3",
    end_color="D9EAD3",
    fill_type="solid"
)

aba[f"D{linha}"].fill = PatternFill(
    start_color="D9EAD3",
    end_color="D9EAD3",
    fill_type="solid"
)


aba[f"C{linha}"].alignment = Alignment(horizontal="center")
aba[f"D{linha}"].alignment = Alignment(horizontal="center")

aba[f"C{linha}"].border = borda
aba[f"D{linha}"].border = borda

aba.freeze_panes = "A2"

aba.auto_filter.ref = f"A1:D{linha-1}"

aba.sheet_view.showGridLines = False

aba.sheet_view.zoomScale = 120

grafico = BarChart()

grafico.title = "Vendas por Produto"
grafico.y_axis.title = "Valor Total"
grafico.x_axis.title = "Produtos"

dados = Reference(
    aba,
    min_col=4,
    min_row=1,
    max_row=linha-1
)

categorias = Reference(
    aba,
    min_col=1,
    min_row=2,
    max_row=linha-1
)

grafico.add_data(dados,titles_from_data=True)
grafico.set_categories(categorias)

aba.add_chart(grafico,"F2")

planilha.save("vendas.xlsx")

print("Planilha criada com sucesso!")