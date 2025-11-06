# -*- coding: utf-8 -*-
"""
Envio de e-mails BOSS, agrupado por vendor, com corpo no modelo solicitado.
Requisitos:
  pip install pandas openpyxl pywin32
Rodar no Windows com Outlook logado.
"""

import os
import datetime as dt
import pandas as pd
from pathlib import Path

# ================== CONFIG ==================
ARQUIVO = r"C:\caminho\teste final.xlsx"   # ajuste o caminho do seu arquivo
ABA = "Planilha1"                           # sua planilha tem essa aba

# Colunas da planilha (conforme arquivo enviado)
COL_ASIN   = "ASIN"
COL_ITEM   = "item_name"
COL_AGING  = "aging"
COL_QTD    = "on_hand_quantity_cw"
COL_CUSTO  = "on_hand_cost_cw"
COL_VENDOR = "vendor_code_name"
COL_EMAIL  = "Email"
COL_NOME   = "MOME do BS ou ISM"  # (mantive o nome exatamente como está na planilha)

# Opcional: se no futuro existirem essas colunas, os filtros serão aplicados
COL_STATUS = "Status"       # opcional
COL_ACTION = "Action"       # opcional
STATUS_ALVO = "Diving Deep"
ACAO_ENVIO_POA_ISM = "ENVIAR POA PARA ISM"

# Outlook
CONTA_OUTLOOK = "ageanerocha853@gmail.com"   # sua conta
CC_OPCIONAL = ""   # "financeiro@empresa.com; cobranca@empresa.com"
BCC_OPCIONAL = ""  # "gestor@empresa.com"

ASSINATURA_TXT = (
    "\n\nAtenciosamente,\n"
    "Geane\n"
    "Retail Vendor Management | Amazon\n"
)

ASSUNTO_FMT = "[BOSS] {vendor} — Itens em Diving Deep — Total R$ {total:,.2f}"

# Modo teste: True salva .msg, False envia
MODO_TESTE = True
PASTA_SAIDA_MSG = "eml_out"

# ================== LEITURA & PREP ==================
def carregar_base(caminho, aba):
    df = pd.read_excel(caminho, sheet_name=aba)
    # Normalizar espaços
    df.columns = [c.strip() for c in df.columns]
    return df

def aplicar_filtros_opcionais(df):
    cols = df.columns.str.lower().tolist()
    tem_status = COL_STATUS.lower() in cols
    tem_action = COL_ACTION.lower() in cols

    if tem_status:
        df = df[df[COL_STATUS].astype(str).str.strip().str.upper() == STATUS_ALVO.upper()]

    # Se a ação for ENVIAR POA PARA ISM, seguiria roteamento específico.
    # Como já temos o e-mail direto na base, usaremos COL_EMAIL na prática.
    # Esse bloco fica aqui caso você queira lógica adicional no futuro.
    # if tem_action:
    #     mask_ism = df[COL_ACTION].astype(str).str.strip().str.upper() == ACAO_ENVIO_POA_ISM.upper()
    #     # Exemplo: tratar algo específico para esses casos se necessário.

    return df

def sanitizar_tipos(df):
    for c in [COL_QTD, COL_CUSTO, COL_AGING]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")
    return df

# ================== CORPO DO E-MAIL ==================
def format_brl(v):
    try:
        return f"R$ {float(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "R$ 0,00"

def montar_lista_itens_texto(df_vendor):
    linhas = []
    for _, row in df_vendor.iterrows():
        asin = str(row.get(COL_ASIN, "")).strip()
        nome = str(row.get(COL_ITEM, "")).strip()
        aging = int(row.get(COL_AGING, 0) or 0)
        qtd = int(row.get(COL_QTD, 0) or 0)
        custo = format_brl(row.get(COL_CUSTO, 0) or 0)
        linhas.append(f"- ASIN {asin} ({nome}), {aging} dias em BOSS, {qtd} unidade(s), custo {custo}.")
    return "\n".join(linhas)

def montar_tabela_html(df_vendor):
    col_map = {
        COL_ASIN: "ASIN",
        COL_ITEM: "Item",
        COL_AGING: "Aging (dias)",
        COL_QTD: "Qtde",
        COL_CUSTO: "Custo (R$)"
    }
    cols = [c for c in col_map if c in df_vendor.columns]
    view = df_vendor[cols].rename(columns=col_map).copy()
    if "Custo (R$)" in view.columns:
        view["Custo (R$)"] = view["Custo (R$)"].fillna(0).map(lambda x: f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
    if "Aging (dias)" in view.columns:
        view["Aging (dias)"] = view["Aging (dias)"].fillna(0).astype(int)
    return view.to_html(index=False, border=0, justify="left")

def corpo_email_modelo(nome_resp, vendor, total_vendor_brl, df_vendor):
    """
    Adaptação do seu modelo para múltiplos ASINs do mesmo vendor.
    """
    lista_texto = montar_lista_itens_texto(df_vendor)
    tabela_html = montar_tabela_html(df_vendor)

    corpo_txt = f"""Olá {nome_resp},

Estou entrando em contato referente aos itens do vendor "{vendor}" que estão atualmente com múltiplas reclamações de clientes. 
O valor atual do BOSS para este vendor é de {total_vendor_brl}, representando um impacto financeiro significativo que requer ação imediata.

Identificamos um padrão onde clientes reportam receber apenas uma unidade do produto quando a página indica duas unidades. 
Esta divergência entre a descrição e o entregue está gerando insatisfação e aumento de devoluções.

Itens impactados:
{lista_texto}

Dado este cenário e considerando o alto valor do BOSS, solicitamos ação imediata para:

1. Verificação junto ao fornecedor sobre:
• Confirmação da quantidade correta que deve ser enviada
• Verificação se houve alteração no bundle do produto
• Análise de possíveis problemas no processo de separação

2. Avaliação das seguintes medidas:
• Atualização imediata da página, caso a quantidade correta seja uma unidade
• Correção do processo de envio, caso a quantidade correta seja duas unidades
• Revisão do catálogo para garantir que todas as informações estejam precisas

Para que os itens sejam removidos do BOSS e voltem a ser comercializados, evitando maiores impactos financeiros, 
é fundamental que estas verificações sejam realizadas e documentadas, garantindo a correção da informação ou do processo de envio.
"""
    corpo_txt += ASSINATURA_TXT

    corpo_html = f"""
    <p>Olá {nome_resp},</p>
    <p>Estou entrando em contato referente aos itens do vendor <b>{vendor}</b> que estão atualmente com múltiplas reclamações de clientes.<br>
    O valor atual do BOSS para este vendor é de <b>{total_vendor_brl}</b>, representando um impacto financeiro significativo que requer ação imediata.</p>

    <p>Identificamos um padrão onde clientes reportam receber apenas uma unidade do produto, quando a página indica duas unidades. 
    Esta divergência entre a descrição e o que é entregue está gerando insatisfação e aumentando a taxa de devolução.</p>

    <p><b>Itens impactados:</b></p>
    <pre style="font-family: Consolas, monospace; white-space: pre-wrap; margin: 8px 0 16px 0;">{lista_texto}</pre>

    <p><b>Tabela de apoio:</b></p>
    {tabela_html}

    <p><b>Solicitações:</b></p>
    <ol>
      <li><b>Verificação junto ao fornecedor</b>:
        <ul>
          <li>Confirmação da quantidade correta que deve ser enviada</li>
          <li>Verificação se houve alteração no bundle do produto</li>
          <li>Análise de possíveis problemas no processo de separação</li>
        </ul>
      </li>
      <li><b>Avaliação de medidas</b>:
        <ul>
          <li>Atualização imediata da página do produto caso a quantidade correta seja uma unidade</li>
          <li>Correção do processo de envio caso a quantidade correta seja duas unidades</li>
          <li>Revisão do catálogo para garantir que todas as informações estejam precisas</li>
        </ul>
      </li>
    </ol>

    <p>Para que os itens sejam removidos do BOSS e voltem a ser comercializados, evitando maiores impactos financeiros,
    é fundamental que estas verificações sejam realizadas e documentadas, garantindo a correção da informação ou do processo de envio.</p>

    <p>Aguardamos seu retorno com urgência.</p>
    <br>
    <p>Atenciosamente,<br>
    Geane<br>
    Retail Vendor Management | Amazon</p>
    """
    return corpo_txt, corpo_html

# ================== OUTLOOK HELPERS ==================
def get_outlook_account(account_email):
    import win32com.client as win32
    outlook = win32.Dispatch("Outlook.Application")
    # Seleciona conta
    conta = None
    for a in outlook.Session.Accounts:
        if str(a).strip().lower() == str(account_email).strip().lower():
            conta = a
            break
    return outlook, conta

def enviar_email(to_addr, assunto, corpo_html, cc="", bcc="", conta=None):
    import win32com.client as win32
    outlook = win32.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.To = to_addr
    if cc:
        mail.CC = cc
    if bcc:
        mail.BCC = bcc
    mail.Subject = assunto
    mail.HTMLBody = corpo_html
    if conta:
        # SendUsingAccount
        mail._oleobj_.Invoke(*(64209, 0, 8, 0, conta))
    mail.Send()

def salvar_msg(to_addr, assunto, corpo_html, cc="", bcc="", pasta=PASTA_SAIDA_MSG):
    import win32com.client as win32
    Path(pasta).mkdir(parents=True, exist_ok=True)
    outlook = win32.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.To = to_addr
    mail.CC = cc
    mail.BCC = bcc
    mail.Subject = assunto
    mail.HTMLBody = corpo_html
    nome = f"{dt.datetime.now():%Y%m%d_%H%M%S}__{to_addr.replace(';','_')}__{assunto[:60].replace(' ','_')}.msg"
    caminho = os.path.join(pasta, nome)
    mail.SaveAs(caminho)
    return caminho

# ================== PIPELINE ==================
def main():
    df = carregar_base(ARQUIVO, ABA)
    df = aplicar_filtros_opcionais(df)
    df = sanitizar_tipos(df)

    # Validar colunas críticas
    cols = set(df.columns)
    obrig = {COL_ASIN, COL_ITEM, COL_AGING, COL_QTD, COL_CUSTO, COL_VENDOR, COL_EMAIL, COL_NOME}
    faltando = obrig - cols
    if faltando:
        raise KeyError(f"Colunas faltantes na planilha: {faltando}")

    # Remover linhas sem e-mail
    df = df[df[COL_EMAIL].astype(str).str.contains("@", na=False)]

    if df.empty:
        print("Sem linhas válidas para envio.")
        return

    # Agrupamento por vendor e e-mail
    grupos = df.groupby([COL_VENDOR, COL_EMAIL], dropna=False)

    outlook = None
    conta = None
    if not MODO_TESTE:
        outlook, conta = get_outlook_account(CONTA_OUTLOOK)

    enviados = 0
    salvos = 0

    for (vendor, email), df_vendor in grupos:
        vendor = str(vendor or "").strip() or "Vendor"
        email  = str(email or "").strip()

        if not email:
            continue

        total_vendor = float(df_vendor[COL_CUSTO].fillna(0).sum())
        total_vendor_brl = format_brl(total_vendor)

        # Nome responsável: pega o mais frequente do grupo
        if COL_NOME in df_vendor.columns:
            try:
                nome_resp = df_vendor[COL_NOME].dropna().astype(str).str.strip()
                nome_resp = nome_resp.mode().iloc[0] if not nome_resp.empty else "Responsável"
            except Exception:
                nome_resp = "Responsável"
        else:
            nome_resp = "Responsável"

        assunto = ASSUNTO_FMT.format(vendor=vendor, total=total_vendor)
        _, corpo_html = corpo_email_modelo(nome_resp, vendor, total_vendor_brl, df_vendor)

        if MODO_TESTE:
            caminho = salvar_msg(email, assunto, corpo_html, CC_OPCIONAL, BCC_OPCIONAL)
            print(f"[DRY RUN] .msg salvo: {caminho}  |  {vendor}  |  {email}  | itens: {len(df_vendor)}")
            salvos += 1
        else:
            enviar_email(email, assunto, corpo_html, CC_OPCIONAL, BCC_OPCIONAL, conta)
            print(f"Enviado: {vendor}  |  {email}  | itens: {len(df_vendor)}")
            enviados += 1

    print(f"Concluído. Enviados: {enviados}, Salvos (.msg): {salvos}")

if __name__ == "__main__":
    main()
