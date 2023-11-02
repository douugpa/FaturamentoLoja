import pandas as pd
import win32com.client as win32

# Importar a base de dados
tabela_vendas = pd.read_excel('Vendas.xlsx')
pd.set_option('display.max_columns', None)

# Calcular o Faturamento por Loja
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
# Calcular quantidade de produtos vendidos por loja
quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
# Ticket médio por produto em cada loja
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})
# Enviar e-mail com o relatório gerado
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'example@mail.com' # Informar qualquer email para receber
mail.Subject = 'Faturamento das Lojas'
mail.HTMLBody = f'''
<p>Prezados,</p>

<p>Segue o relatório de Faturamento das Lojas.</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade Vendida:</p>
{quantidade.to_html()}

<p>Ticket Médio dos Produtos por Loja: </p>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

<p>Qualquer dúvida, estamos à disposição.</p>
<br>
<p>Att,</p>

'''

mail.Send()