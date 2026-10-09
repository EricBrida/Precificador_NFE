# Imports
from salvar_excel import salvar_em_excel

# Função para calculo do custo
def custoaquisicao (precounit, qtd, icms, ipi):
    custoaqui = (precounit + (ipi / qtd)) - (icms / qtd)

    return custoaqui

# Função para salvar em txt
def passar_txt(cont1, cont2, num_nfe):
    with open(f"NFE/{num_nfe}.txt", "a", encoding="utf-8") as f:
        f.write(cont1)
        f.write(cont2)

# Função "mãe" da precificação
def precificar(num_nfe, fornecedor, produtos):

    print(f"Numero NFE: {num_nfe}")

    for produto in produtos:
        codigo = produto['codigo']
        descricao = produto['descricao']
        qtd = produto['quantidade']
        preco = produto['preco']
        preco_total = produto['preco_total']
        icms = produto['icms']
        ali_icms = produto['ali_icms']
        ipi = produto['ipi']
        ali_ipi = produto['ali_ipi']

        custo = custoaquisicao(preco, qtd, icms, ipi)

        print(f"\n\nCódigo - {codigo} || Preco Unitário NF-E - {preco:.2f} || ICMS - {icms:.2f} Aliquota - {ali_icms:.2f} || IPI - {ipi:.2f} Aliquota - {ali_ipi:.2f} ")
        print(f"Preço para calcular margem mínima -> R$ {custo:.2f}\n\n")

        cont1 = f"Código - {codigo} || Preco Unitário NF-E - {preco:.2f} || ICMS - {icms:.2f} Aliquota - {ali_icms:.2f} || IPI - {ipi:.2f} Aliquota - {ali_ipi:.2f} "
        cont2 = f"Preço para calcular margem mínima -> R$ {custo:.2f}\n\n\n"

        passar_txt(cont1, cont2, num_nfe)
        # Necessário arrumar armazenamento dos dados
        # salvar_em_excel(codigo, descricao, qtd, preco, preco_total, icms, ali_icms, ipi, ali_ipi, custo, num_nfe)