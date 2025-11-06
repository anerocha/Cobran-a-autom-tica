import win32com.client as client
import pandas as pd
import datetime as dt

# ================= Lendo o arquivo Excel =================
tabela = pd.read_excel('teste final.xlsx', sheet_name='Planilha1')  # ajuste se mudar a aba
print(tabela.head())
print(tabela.info())

# ================= Data atual =================
hoje = dt.datetime.now()
print("Execução em:", hoje)

# ================= Colunas esperadas (conforme sua planilha) =================
COL_ASIN   = 'ASIN'
COL_ITEM   = 'item_name'
COL_AGING  = 'aging'
COL_QTD    = 'on_hand_quantity_cw'
COL_CUSTO  = 'on_hand_cost_cw'
COL_VENDOR = 'vendor_code_name'
COL_EMAIL  = 'Email'
COL_NOME   = 'MOME do BS ou ISM'   # exatamente como está no arquivo enviado

# Validação simples de colunas
colunas_necessarias = [COL_ASIN, COL_ITEM, COL_AGING, COL_QTD, COL_CUSTO, COL_VENDOR, COL_EMAIL, COL_NOME]
faltando = [c for c in colunas_necessarias if c not in tabela.columns]
if faltando:
    raise KeyError(f"Faltam colunas na planilha: {faltando}")

# Filtrar linhas com e-mail válido
tabela = tabela[tabela[COL_EMAIL].astype(str).str.contains('@', na=False)]
print("Linhas após validar e-mail:", len(tabela))

# Casts básicos
tabela[COL_AGING] = pd.to_numeric(tabela[COL_AGING], errors='coerce').fillna(0).astype(int)
tabela[COL_QTD]   = pd.to_numeric(tabela[COL_QTD], errors='coerce').fillna(0).astype(int)
tabela[COL_CUSTO] = pd.to_numeric(tabela[COL_CUSTO], errors='coerce').fillna(0.0)

# ================= Agrupar por vendor e e-mail =================
grupos = tabela.groupby([COL_VENDOR, COL_EMAIL], dropna=False)

# Montar estrutura "dados" no mesmo espírito do seu código antigo
# Cada item de "dados" é um dicionário com tudo que precisamos para um e-mail
dados = []
for (vendor, email), df_vendor in grupos:
    nome_resp_series = df_vendor[COL_NOME].dropna().astype(str).str.strip()
    nome_resp = nome_resp_series.mode().iloc[0] if not nome_resp_series.empty else "Responsável"
    total_vendor = float(df_vendor[COL_CUSTO].sum())
    dados.append({
        "destinatario": email,
        "vendor": str(vendor) if pd.notna(vendor) else "Vendor",
        "nome": nome_resp,
        "total_vendor": total_vendor,
        "itens": df_vendor[[COL_ASIN, COL_ITEM, COL_AGING, COL_QTD, COL_CUSTO]].copy()
    })

print(f"Total de e-mails a preparar: {len(dados)}")

# ================= Inicializando o Outlook =================
outlook = client.Dispatch('Outlook.Application')
emissor = outlook.Session.Accounts['ageanerocha853@gmail.com']  # ajuste sua conta

# ================= Função para formatar BRL =================
def brl(v):
    try:
        return f"R$ {float(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except:
        return "R$ 0,00"

# ================= Envio de e-mails =================
for d in dados:
    destinatario = d["destinatario"]
    vendor       = d["vendor"]
    nome         = d["nome"]
    total_boss   = brl(d["total_vendor"])
    itens        = d["itens"]

    # Assunto
    assunto = f"[BOSS] {vendor} — Itens em Diving Deep — Total {total_boss}"

    # Monta lista de itens em texto, estilo rápido
    linhas = []
    for _, row in itens.iterrows():
        asin  = str(row[COL_ASIN]).strip()
        item  = str(row[COL_ITEM]).strip()
        aging = int(row[COL_AGING])
        qtd   = int(row[COL_QTD])
        custo = brl(row[COL_CUSTO])
        linhas.append(f"- ASIN {asin} ({item}), {aging} dias em BOSS, {qtd} unidade(s), custo {custo}")

    lista_itens_txt = "\n".join(linhas)

    # Corpo no modelo que você pediu, adaptado para múltiplos itens
    corpo_mensagem = f"""
Olá {nome},

Estou entrando em contato referente aos itens do vendor "{vendor}" que estão atualmente com múltiplas reclamações de clientes.
O valor atual do BOSS para este vendor é de {total_boss}, representando um impacto financeiro significativo que requer ação imediata.

Identificamos um padrão preocupante onde os clientes reportam receber apenas uma unidade do produto, quando a página de vendas indica claramente que o item deveria conter duas unidades. Esta divergência entre a descrição do produto e o que está sendo entregue está gerando insatisfação dos clientes e aumentando a taxa de devolução.

Itens impactados:
{lista_itens_txt}

Dado este cenário e considerando o alto valor do BOSS impactado, solicitamos ação imediata para:

1. Verificação junto ao fornecedor sobre:
• Confirmação da quantidade correta que deve ser enviada
• Verificação se houve alteração no bundle do produto
• Análise de possíveis problemas no processo de separação

2. Avaliação das seguintes medidas:
• Atualização imediata da página do produto caso a quantidade correta seja uma unidade
• Correção do processo de envio caso a quantidade correta seja duas unidades
• Revisão do catálogo para garantir que todas as informações estejam precisas

Para que os itens sejam removidos do BOSS e voltem a ser comercializados, evitando assim maiores impactos financeiros, é fundamental que estas verificações sejam realizadas e documentadas, garantindo a correção da informação ou do processo de envio.

Aguardamos seu retorno com urgência.

Atenciosamente,
Geane
Retail Vendor Management | Amazon
""".strip()

    # Criando a mensagem
    mensagem = outlook.CreateItem(0)
    mensagem.To = destinatario
    mensagem.Subject = assunto
    mensagem.Body = corpo_mensagem

    # Enviando a mensagem
    mensagem._oleobj_.Invoke(*(64209, 0, 8, 0, emissor))  # usa a conta emissora
    mensagem.Save()
    mensagem.Send()
    print(f"E-mail enviado para {destinatario} | vendor: {vendor} | itens: {len(itens)}")

print("E-mails enviados com sucesso.")
