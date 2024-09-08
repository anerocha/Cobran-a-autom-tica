import win32com.client as client
import pandas as pd
import datetime as dt

# Lendo o arquivo Excel
tabela = pd.read_excel('Contas a Receber.xlsx')
print(tabela)
print(tabela.info())

# Obtendo a data atual
hoje = dt.datetime.now()
print(hoje)

# Coletando apenas os dados de clientes que estão devendo
tabela_devedores = tabela.loc[tabela['Status'] == 'Em aberto']
print(tabela_devedores)
tabela_devedores = tabela_devedores.loc[tabela_devedores['Data Prevista para pagamento'] < hoje]
print(tabela_devedores)

# Inicializando o Outlook
outlook = client.Dispatch('Outlook.Application')
emissor = outlook.Session.Accounts['ageanerocha853@gmail.com']  # Substitua pelo seu e-mail

# Preparando os dados para envio
dados = tabela_devedores[['Valor em aberto', 'Data Prevista para pagamento', 'E-mail', 'NF']].values.tolist()

# Enviando o e-mail para todos os destinatários
for dado in dados:
    destinatario = dado[2]
    nf = dado[3]
    prazo = dado[1]
    prazo = prazo.strftime("%d/%m/%Y")
    valor = dado[0]
    assunto = 'Atraso de pagamento'
    
    # Criando a mensagem
    mensagem = outlook.CreateItem(0)
    mensagem.To = destinatario
    mensagem.Subject = assunto
    corpo_mensagem = f'''
    Prezado Cliente,

    Verificamos um atraso no pagamento referente à NF {nf} com vencimento em {prazo} e valor total de R${valor:.2f}.
    Gostaríamos de verificar se há algum problema que necessite de auxílio de nossa equipe. 

    Em caso de dúvidas, é só entrar em contato com nosso time através do e-mail tecobreiautomaticamentecompython@gmail.com

    Atenciosamente,
    Candiotto da Hashtag
    '''
    mensagem.Body = corpo_mensagem
    
    # Enviando a mensagem
    mensagem._oleobj_.Invoke(*(64209, 0, 8, 0, emissor))
    mensagem.Save()
    mensagem.Send()

print("E-mails enviados com sucesso.")
