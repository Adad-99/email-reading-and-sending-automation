import pandas as pd
import win32com.client as win32
from tkinter import filedialog

#Buscar arquivo a ser lido em pastas
arquivo = filedialog.askopenfilename(filetypes=[("", "*.xls;*.xlsx")])

# Importar a base de dados
tabela_vendas = pd.read_excel(arquivo)

# Visualizar a base de dados
pd.set_option('display.max_columns', None)

# Faturamento por loja
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()

# Quantidade de produtos vendidos por loja
qtd_vend = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()

# Ticket medio por produto em cada loja
ticket_medio = (faturamento['Valor Final'] / qtd_vend['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Medio'})

# Enviar um email com o relatorio 
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.to = 'Email@a_receber.com'
mail.subject = 'Relatorio de Vendas por Loja'
mail.HTMLBody = f'''

<p>Prezados,</p>

<p>segue o Relarorio de vendas por cada loja.</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidaede Vendida:</p>
{qtd_vend.to_html()}

<p>Ticket Médio dos produtos em cada Loja:</p>
{ticket_medio.to_html(formatters={'Ticket Medio': 'R${:,.2f}'.format})}

<p>Att.,</p>
<p>Usuario</p>
'''

mail.Send()
print('Email enviado')