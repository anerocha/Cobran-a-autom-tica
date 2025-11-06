import win32com.client as client
import pandas as pd
import datetime as dt

# ======================================
# CONFIGURAÇÃO
# ======================================
ARQUIVO_EXCEL = r'/mnt/data/teste final.xlsx'  # Caminho da planilha
EMAIL_CONTA_EMISSORA = 'ageanerocha853@gmail.com'  # E-mail emissor no Outlook

# ======================================
# LENDO A PLANILHA
# ======================================
tabela = pd.read_excel(ARQUIVO_EXCEL)

# Normaliza os nomes das colunas
tabela.columns = tabela.columns.str.strip().str.lower()

# Renomeie se necessário para ajustar aos nomes exatos
# Exemplo: se o arquivo tiver 'ItemName' em vez de 'item_name', ajuste abaixo
tabela = tabela.rename(columns={
    'itemname': 'item_name',
    'itename': 'item_name',
    'on_hand_cost_cw': 'valor',
    'aging': 'dias',
    'email': 'email',
    'owner_bs': 'nome'
})

# Mostra uma prévia dos dados
print(tabela[['asin', 'item_name', 'dias', 'valor', 'email', 'nome']].head())

# ======================================
# CONFIGURANDO O OUTLOOK
# ======================================
outlook = client.Dispatch('Outlook.Application')
emissor = outlook.Session.Accounts[EMAIL_CONTA_EMISSORA]

# ======================================
# ENVIANDO OS E-MAILS
# ======================================
for i, linha in tabela.iterrows():
    asin = str(linha['asin'])
    produto = str(linha['item_name'])
    dias = int(linha['dias'])
    valor = float(linha['valor'])
    destinatario = str(linha['email']).strip()
    nome = str(linha['nome']).strip()

    # Assunto
    assunto = f"Ação necessária – ASIN {asin} com reclamações de clientes"

    # Corpo da mensagem (TEXTO EXATAMENTE COMO VOCÊ PEDIU)
    corpo = f"""
Olá {nome},

Estou entrando em contato referente ao ASIN {asin} ({produto}) que está atualmente com múltiplas reclamações de clientes. 
O ASIN já está em BOSS há {dias} dias. O valor atual do BOSS é de $ {valor:,.2f}, representando um impacto financeiro significativo que requer ação imediata.

Identificamos um padrão preocupante onde os clientes reportam receber apenas uma unidade do produto, quando a página de vendas indica claramente que o item deveria conter duas unidades. 
Esta divergência entre a descrição do produto e o que está sendo entregue está gerando insatisfação dos clientes e aumentando a taxa de devolução.

Dado este cenário e considerando o alto valor do BOSS impactado, solicitamos ação imediata para:

1. Verificação junto ao fornecedor sobre:
• Confirmação da quantidade correta que deve ser enviada
• Verificação se houve alteração no bundle do produto
• Análise de possíveis problemas no processo de separação

2. Avaliação das seguintes medidas:
• Atualização imediata da página do produto caso a quantidade correta seja uma unidade
• Correção do processo de envio caso a quantidade correta seja duas unidades
• Revisão do catálogo para garantir que todas as informações estejam precisas

Para que o item seja removido do BOSS e volte a ser comercializado, evitando assim maiores impactos financeiros, é fundamental que estas verificações sejam realizadas e documentadas, garantindo a correção da informação ou do processo de envio.

Aguardamos seu retorno com urgência.

Atenciosamente,
[SEU NOME]
"""

    # Criando e enviando o e-mail
    mensagem = outlook.CreateItem(0)
    mensagem.To = destinatario
    mensagem.Subject = assunto
    mensagem.Body = corpo

    mensagem._oleobj_.Invoke(*(64209, 0, 8, 0, emissor))
    mensagem.Save()
    mensagem.Send()

print("E-mails enviados com sucesso.")
