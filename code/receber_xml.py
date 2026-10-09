# Imports
import os
import sys
import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from decimal import Decimal, ROUND_HALF_UP

from precificar import precificar

# Setando estrutura da XML
NS = {"nfe": "http://www.portalfiscal.inf.br/nfe"}


# Função para ler o XML
def ler_xml(caminho):
    arvore = ET.parse(caminho)
    raiz = arvore.getroot()

    # Definindo os caminhos necessários para navegar na XML
    numero = raiz.findtext(
        ".//nfe:ide/nfe:nNF", namespaces=NS
    )

    fornecedor = raiz.findtext(
        ".//nfe:emit/nfe:xNome", namespaces=NS
    )

    produtos = []

    for item in raiz.findall(".//nfe:infNFe/nfe:det", NS):
        prod = item.find("nfe:prod", NS)
        imposto = item.find("nfe:imposto", NS)

        if prod is None:
            continue

        # Definição do caminho de impostos
        icms = imposto.find("nfe:ICMS", NS) if imposto is not None else None
        grupo_icms = next(iter(icms), None) if icms is not None else None
        ipi = imposto.find("nfe:IPI/nfe:IPITrib", NS) if imposto is not None else None

        quantidade = Decimal(prod.findtext("nfe:qCom", default="0", namespaces=NS))
        preco = Decimal(prod.findtext("nfe:vUnCom", default="0", namespaces=NS))
        preco_total = preco * quantidade

        ali_icms =Decimal(
            grupo_icms.findtext("nfe:pICMS", default=0, namespaces=NS)
        ) if grupo_icms is not None else Decimal("0")
        
        ali_ipi = Decimal(
            ipi.findtext("nfe:pIPI", default=0, namespaces=NS)
        ) if ipi is not None else Decimal("0")

        icms = preco_total * (ali_icms / 100) if ali_icms > 0 else Decimal("0")
        ipi = preco_total * (ali_ipi / 100) if ali_ipi > 0 else Decimal("0")

        custo = (preco + (ipi / quantidade) - (icms / quantidade))

        # Arredondando as casas
        preco = preco.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
        preco_total = preco_total.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
        icms = icms.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
        ali_icms = ali_icms.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
        ipi = ipi.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
        ali_ipi = ali_ipi.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
        custo = custo.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)

        produtos.append({
            "codigo": prod.findtext("nfe:cProd", namespaces=NS),
            "descricao": prod.findtext("nfe:xProd", namespaces=NS),
            "quantidade": quantidade,
            "preco": preco,
            "preco_total": preco_total,
            "icms": icms,
            "ali_icms": ali_icms,
            "ipi": ipi,
            "ali_ipi": ali_ipi,
            "custo": custo
        })

    return numero, fornecedor, produtos


# Função para selecionar e carregar o XML
def selecionar_xml(label_arquivo, label_nota, label_fornecedor, tabela):
    caminho = filedialog.askopenfilename(
        title="Selecione o XML da NF-e",
        filetypes=[("Arquivos XML", "*.xml")]
    )

    if not caminho:
        return

    try:
        numero, fornecedor, produtos = ler_xml(caminho)

        # Passa os respectivos valores para calcular e salvar em txt e excel
        precificar(numero, fornecedor, produtos)

        print(f"{numero} \n\n {fornecedor} \n\n {produtos}")

        label_arquivo.config(
            text=os.path.basename(caminho)
        )

        label_nota.config(
            text=f"NF-e: {numero or '-'}"
        )

        label_fornecedor.config(
            text=f"Fornecedor: {fornecedor or '-'}"
        )

        # Limpa a tabela anterior
        for item in tabela.get_children():
            tabela.delete(item)

        # Insere os novos produtos
        for p in produtos:
            tabela.insert("", "end", values=(
                p["codigo"],
                p["descricao"],
                p["quantidade"],
                p["preco"],
                p["preco_total"],
                p["icms"],
                p["ali_icms"],
                p["ipi"],
                p["ali_ipi"],
                p["custo"]
            ))

    except Exception as erro:
        messagebox.showerror(
            "Erro ao ler XML", str(erro)
        )


# Função que cria a interface
def carregarUi():
    janela = tk.Tk()
    janela.title("Leitor de NF-e")
    janela.geometry("1600x900")

    if getattr(sys, "frozen", False):
        pasta_base = sys._MEIPASS
    else:
        pasta_base = os.path.dirname(os.path.abspath(__file__))

    caminho_icone = os.path.join(pasta_base, "icone_nfe.ico")

    janela.iconbitmap(caminho_icone)

    titulo = tk.Label(
        janela,
        text="Leitor de Notas Fiscais Eletrônicas",
        font=("Arial", 16, "bold")
    )
    titulo.pack(pady=15)

    label_arquivo = tk.Label(
        janela,
        text="Nenhum arquivo selecionado"
    )
    label_arquivo.pack(pady=5)

    label_nota = tk.Label(
        janela,
        text="NF-e: -"
    )
    label_nota.pack(pady=5)

    label_fornecedor = tk.Label(
        janela,
        text="Fornecedor: -"
    )
    label_fornecedor.pack(pady=5)

    # Configuração da tabela
    colunas = (
        "Codigo",
        "Descricao",
        "Quantidade",
        "Preco",
        "Preco Total",
        "ICMS Total",
        "Aliquota ICMS",
        "IPI Total",
        "Aliquota IPI",
        "Custo"
    )

    tabela = ttk.Treeview(
        janela,
        columns=colunas,
        show="headings"
    )

    for coluna in colunas:
        tabela.heading(coluna, text=coluna)
        tabela.column(coluna, width=130)

    tabela.column("Descricao", width=350)

    tabela.pack(
        fill="both",
        expand=True,
        padx=15,
        pady=15
    )

    # Botão para selecionar XML
    botao = tk.Button(
        janela,
        text="Selecionar XML da NF-e",
        command=lambda: selecionar_xml(
            label_arquivo,
            label_nota,
            label_fornecedor,
            tabela
        ),
        font=("Arial", 12),
        padx=15,
        pady=8
    )

    botao.pack(pady=10)

    janela.mainloop()