import pandas as pd
import win32com.client as win32


df = pd.read_excel(r"#seu caminho")

pd.set_option('display.max_columns', None)

faturamento = df[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
#faturamento = faturamento.rename(columns={'Valor Final': 'Valor final'})

print(faturamento)
print('-' * 50)

qt = df[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()

print(qt)
print('-' * 50)

tm = (faturamento['Valor Final'] / qt ['Quantidade']).to_frame()
tm = tm.rename(columns={0: 'Ticket médio'})
print(tm)

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = '#SEU EMAIL'
mail.Subject = 'Relatório de Vendas por Loja'
mail.HTMLBody = f'''

<p>Prezados, segue o relatório de vendas por cada loja.</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final' : 'R${:,.2f}'.format })}

<p>Quantidade vendida:</p>
{qt.to_html()}

<p>Ticket médio dos produtos em cada loja:</p>
{tm.to_html(formatters={'Ticket médio' : 'R${:,.2f}'.format })}


<p>Qualquer dúvida estou a disposição.</p>

<p>ATT..</p>
<p>Guedes</p>


'''



mail.Send()

print('Email enviado')










