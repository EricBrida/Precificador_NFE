import os
import json
from openpyxl import Workbook

arquivo = 'valores_nfe.json'

def carregar_arquivos(arquivo):
    with open(arquivo, "r", encoding="utf-8") as f:
        dados = json.load(f)
        return dados
    
def custoaquisicao (precounit, ipi, icms):
    custoaqui = precounit + ipi + icms

    return custoaqui

def icmscatt (precounit, ipi, ali):
    icmsc = (precounit + ipi) * ali

    return icmsc 

def custoliquido (precounit, ipi, icms, ali):
    custoaqui = custoaquisicao (precounit, ipi, icms)
    icmscat = icmscatt(precounit, ipi, ali)

    return custoaqui - (icmscat + icms)

def passar_txt(cont1, cont2):
    with open(f"NFE/{num_nfe}.txt", "a", encoding="utf-8") as f:
        f.write(cont1)
        f.write(cont2)

def salvar_em_excel(codigo, precounitario_nfe, ipi_nfe, icms_nfe, ali_nfe, items_precificados):
    ws.append([
        codigo,
        precounitario_nfe,
        round(ipi_nfe, 2),
        round(icms_nfe, 2),
        ali_nfe,
        round(items_precificados, 2)
    ])

items = carregar_arquivos(arquivo)

num_nfe = input("Digite o número da nota / arquivo: ")

wb = Workbook()
ws = wb.active
ws.title = num_nfe

ws.append([
    "Código",
    "Preço Unitário NF-E",
    "IPI",
    "ICMS",
    "Alíquota",
    "Preço Precificado"
])

codigo = items['codigo']
qtditems = items['quantidade']
precounitario_nfe = items['valor_unitario_nfe']
icms_nfe = items['icms']
ipi_nfe = items['ipi']
ali_nfe = items['aliquota']

items_precificados = []

for i in range(len(codigo)):

    precounit = float(precounitario_nfe[i])
    ipi = float(ipi_nfe[i] / qtditems[i])
    icms = float(icms_nfe[i] / qtditems[i])
    ali = float(ali_nfe[i])
    # ali = ali if ali != 0 else 0.04

    precounitario_nfe[i] = precounit
    ipi_nfe[i] = ipi
    icms_nfe[i] = icms
    ali_nfe[i] = ali

    items_precificados.append(custoliquido(precounit, ipi, icms, ali))

os.system("cls")

for i in range(len(codigo)):

    print(f"Código - {codigo[i]} || Preco Unitário NF-E - {precounitario_nfe[i]} || IPI - {ipi_nfe[i]:.2f} || ICMS - {icms_nfe[i]:.2f} || Aliquota - {ali_nfe[i]}")
    print(f"Preço para calcular margem mínima -> R$ {items_precificados[i]:.2f}\n\n\n")

    cont1 = f"Código - {codigo[i]} || Preco Unitário NF-E - {precounitario_nfe[i]} || IPI - {ipi_nfe[i]:.2f} || ICMS - {icms_nfe[i]:.2f} || Aliquota - {ali_nfe[i]}\n"
    cont2 = f"Preço para calcular margem mínima -> R$ {items_precificados[i]:.2f}\n\n"

    passar_txt(cont1, cont2)
    salvar_em_excel(codigo[i], precounitario_nfe[i], ipi_nfe[i], icms_nfe[i], ali_nfe[i], items_precificados[i])

wb.save(f"Excel/{num_nfe}.xlsx")