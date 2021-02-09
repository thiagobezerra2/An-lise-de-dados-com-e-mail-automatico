import pandas as pd

tabela_vendas = pd.read_excel("/content/drive/MyDrive/Colab Notebooks/ Vendas.xlsx")
display(tabela_vendas)

tabela_faturamento = tabela_vendas[['ID Loja','Valor Final']].groupby('ID Loja').sum()
tabela_faturamento = tabela_faturamento.sort_values(by="Valor Final", ascending=False)
display(tabela_faturamento)

tabela_quantidade = tabela_vendas [['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
display(tabela_quantidade)

ticket_medio = (tabela_faturamento["Valor Final"] / tabela_quantidade ["Quantidade"]).to_frame()
ticket_medio = ticket_medio.rename(columns= {0: "Ticket Medio"})
display(ticket_medio)

def enviar_email(nome_da_loja, tabela):
  import smtplib
  import email.message

  server = smtplib.SMTP('smtp.gmail.com:587')  
  corpo_email = f"""
  <p> Prezados,</p>
  <p> Segue o relatorio de vendas </p>
  {tabela.to_html()}
  <p>Anteciosamente</p>
  """
    
  msg = email.message.Message()
  msg['Subject'] = f"Relatorio de vendas  - {nome_da_loja}"
    
  # Fazer antes (apenas na 1ª vez): Ativar Aplicativos não Seguros.
    # Gerenciar Conta Google -> Segurança -> Aplicativos não Seguros -> Habilitar
  # Caso mesmo assim dê o erro: smtplib.SMTPAuthenticationError: (534,
  # Você faz o login no seu e-mail e depois entra em: https://accounts.google.com/DisplayUnlockCaptcha
  msg['From'] = 'email@gmail.com'
  msg['To'] = 'email@gmail.com'
  password = "senha"
  msg.add_header('Content-Type', 'text/html')
  msg.set_payload(corpo_email )
    
  s = smtplib.SMTP('smtp.gmail.com: 587')
  s.starttls()
  # Login Credentials for sending the mail
  s.login(msg['From'], password)
  s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))
  print('Email enviado')

tabela_completa = tabela_faturamento.join(tabela_quantidade).join(ticket_medio)
enviar_email("Diretoria", tabela_completa)

lojas = tabela_vendas["ID Loja"].unique()
print(lojas)