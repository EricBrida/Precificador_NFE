# Imports
from openpyxl import Workbook

# Função para salvar os dados em uma planilha Excel
def salvar_em_excel(codigo, descricao, qtd, preco, preco_total, icms, ali_icms, ipi, ali_ipi, custo, num_nfe):

    wb = Workbook()
    ws = wb.active
    ws.title = num_nfe

    ws.append([
        "Código",
        "Descrição",
        "Quantidade",
        "Preço",
        "Preço Total",
        "ICMS",
        "Aliquota ICMS",
        "IPI",
        "Aliquota IPI",
        "Custo Final"
    ])

    ws.append([
        codigo,
        descricao,
        qtd,
        round(preco, 2),
        round(preco_total, 2),
        round(icms, 2),
        round(ali_icms, 2),
        round(ipi, 2),
        round(ali_ipi, 2),
        round(custo, 2),
    ])

    wb.save(f"Excel/{num_nfe}.xlsx")