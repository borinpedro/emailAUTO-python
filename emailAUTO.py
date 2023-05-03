import pandas as pd
import win32com.client as win32
from time import sleep







sleep(3)

print('\33[32m-=-'*10)
print('INICIANDO BANCO DE DADOS...')
print('-=-'*10)

sleep(2)

print('Abrindo DATA FRAME...')
print('\33[32m-=-'*10)

sleep(2)

#importando a planilha do excel para base de dados
df = pd.read_excel('nomeDoArquivo.xlsx')

 #visualizando a tabela
pd.set_option('display.max_columns',None)

#dados da tabela que quer importar para visualização, agrupados por ...
valor1 = df[['frutas','preço','total']]
print(valor1)



# enviar um email com o relatório
#neste exemplo, o codigo ira iniciar a aplicação do outlook com o perfil existente no windows e enviara atraves dele o email
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'destinatario@gmail.com'
mail.Subject = 'Assunto'
mail.HTMLBody = f'''

<p>prezado(a),</p>

<p>Segue o exemplo de relatorio a ser implantado no sistema.</p>

<p>segue os valores:</p>
{valor1.to_html(formatters={'preço':'R${:,.2f}'.format ,'total': 'R${:,.2f}'.format})}


<p>Qualquer alteração entre em contato comigo.</p>

<p>Atenciosamente.,</p>
<p>Borin</p>
'''

mail.Send()

print('Email Enviado')
