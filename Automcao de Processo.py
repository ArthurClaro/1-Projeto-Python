### Passo 1 - Importar Arquivos e Bibliotecas
import pandas as pd
import win32com.client as win32 
from pathlib import Path



emails= pd.read_excel (r"C:\Users\01234\Downloads\Projeto AutomacaoIndicadores\Projeto AutomacaoIndicadores\Bases de Dados\Emails.xlsx" ) 
# display(emails)
lojas= pd.read_csv (r"C:\Users\01234\Downloads\Projeto AutomacaoIndicadores\Projeto AutomacaoIndicadores\Bases de Dados\Lojas.csv" ,sep=';',encoding='ISO-8859-1' ) 
# display(lojas)
vendas= pd.read_excel (r"C:\Users\01234\Downloads\Projeto AutomacaoIndicadores\Projeto AutomacaoIndicadores\Bases de Dados\Vendas.xlsx" ) 
# display(vendas)

### Passo 2 - Definir Criar uma Tabela para cada Loja e Definir o dia do Indicador
#tranzendo a tabela Lojas para dentro de Vendas :
vendas = vendas.merge(lojas, on='ID Loja')
# display(vendas)

# Criar 1 arquivo pra cada Loja : criar 1 dataframe pra cada Loja 
# criar uma tabela pra cada Loja
dicionario_lojas = {}
for loja in lojas["Loja"]:
    # print(loja)
    dicionario_lojas[loja] = vendas.loc[vendas["Loja"]==loja, :]  #loc [linha,coluna]

# display(dicionario_lojas["Iguatemi Esplanada"])
# display(dicionario_lojas["Rio Mar Recife"])
# dicionario_lojas

#pegando o Dia do Indicador :
# vendas['Data'][-1:]
dia_indicador = vendas["Data"].max()
print(dia_indicador)
print('{}/{}'.format(dia_indicador.day, dia_indicador.month))

### Passo 3 - Salvar a planilha na pasta de backup
#indentificar se ja existe a loja: Porque se eu receber uma nova lista, Ele não Add uma nova Pasta 

caminho_backup = Path(r"C:\Users\01234\Downloads\Projeto AutomacaoIndicadores\Projeto AutomacaoIndicadores\Backup Arquivos Lojas")

arquivos_pasta_backup = caminho_backup.iterdir()
# lista_nome_backup = []
# for arquivo in arquivos_pasta_backup:
#     lista_nome_backup.append(arquivo.name)
lista_nome_backup = [arquivo.name for arquivo in arquivos_pasta_backup]
# print(lista_nome_backup)

for loja in dicionario_lojas:
    if loja not in lista_nome_backup: #ou seja se não estiver no meu array
            nova_pasta = caminho_backup / loja 
            nova_pasta.mkdir()


    #salvar dentro da pasta
    nome_arquivo = '{}_{}_{}.xlsx'.format(dia_indicador.month, dia_indicador.day, loja)
    print(nome_arquivo)
    local_arquivo =  caminho_backup /loja / nome_arquivo
    # print(local_arquivo)

    dicionario_lojas[loja].to_excel(local_arquivo)

    
### Passo 4 - Calcular o indicador para 1 loja
# definição de metas

meta_faturamento_dia = 1000
meta_faturamento_ano= 1650000
meta_qtdeprodutos_dia =4
meta_qtdeprodutos_ano = 120
meta_ticketmedio_dia=500
meta_ticketmedio_ano=500

for loja in dicionario_lojas:

# display(dicionario_lojas[loja]['Valor Final'].sum())
   vendas_loja = dicionario_lojas[loja]
   vendas_loja_dia = vendas_loja.loc[vendas_loja['Data']==dia_indicador , :] 
   # display(vendas_loja_dia)

   # faturamento
   faturamento_ano = vendas_loja['Valor Final'].sum()
   # print(faturamento_ano)
   faturamento_dia= vendas_loja_dia['Valor Final'].sum()
   # print(faturamento_dia)

   # diversidade de produtos
   qtde_produtos_ano= len(vendas_loja['Produto'].unique())
   # print(qtde_produtos_ano)

   qtde_produtos_dia = len(vendas_loja_dia['Produto'].unique())
   # print(qtde_produtos_dia)

   # ticket medio
   valor_venda= vendas_loja.groupby('Código Venda').sum()
      # display(valor_venda)
   ticket_medio_ano = valor_venda['Valor Final'].mean()
   # print(ticket_medio_ano)

   valor_venda_dia = vendas_loja_dia.groupby('Código Venda').sum()
      # display(valor_venda_dia)
   ticket_medio_dia = valor_venda_dia['Valor Final'].mean()
   # print(ticket_medio_dia)

# enviar email:

   outlook = win32.Dispatch('outlook.application')
   # display(emails)
   # adad= emails.loc[emails['Loja']==loja, 'E-mail'].values[0]
   # print(adad)


   nome= emails.loc[emails['Loja']==loja, 'Gerente'].values[0]
   mail = outlook.CreateItem(0) # para criar emaol 
   mail.To = emails.loc[emails['Loja']==loja, 'E-mail'].values[0]
   mail.Subject = 'OnePage Dia {}/{} - Loja {}'.format(dia_indicador.day, dia_indicador.month,loja)
   # mail.Body = 'Texto do E-mail'




   if faturamento_dia >= meta_faturamento_dia:
         cor_fat_dia = 'green'
   else:
         cor_fat_dia = 'red'
   if faturamento_ano >= meta_faturamento_ano:
         cor_fat_ano = 'green'
   else:
         cor_fat_ano = 'red'
   if qtde_produtos_dia >= meta_qtdeprodutos_dia:
         cor_qtde_dia = 'green'
   else:
         cor_qtde_dia = 'red'
   if qtde_produtos_ano >= meta_qtdeprodutos_ano:
         cor_qtde_ano = 'green'
   else:
         cor_qtde_ano = 'red'
   if ticket_medio_dia >= meta_ticketmedio_dia:
         cor_ticket_dia = 'green'
   else:
         cor_ticket_dia = 'red'
   if ticket_medio_ano >= meta_ticketmedio_ano:
         cor_ticket_ano = 'green'
   else:
         cor_ticket_ano = 'red'

   mail.HTMLBody = f'''
      <p>Bom dia, {nome}</p>

      <p>O resultado de ontem <strong>({dia_indicador.day}/{dia_indicador.month})</strong> da <strong>Loja {loja}</strong> foi:</p>

      <table>
         <tr>
         <th>Indicador</th>
         <th>Valor Dia</th>
         <th>Meta Dia</th>
         <th>Cenário Dia</th>
         </tr>
         <tr>
         <td>Faturamento</td>
         <td style="text-align: center">R${faturamento_dia:.2f}</td>
         <td style="text-align: center">R${meta_faturamento_dia:.2f}</td>
         <td style="text-align: center"><font color="{cor_fat_dia}">◙</font></td>
         </tr>
         <tr>
         <td>Diversidade de Produtos</td>
         <td style="text-align: center">{qtde_produtos_dia}</td>
         <td style="text-align: center">{meta_qtdeprodutos_dia}</td>
         <td style="text-align: center"><font color="{cor_qtde_dia}">◙</font></td>
         </tr>
         <tr>
         <td>Ticket Médio</td>
         <td style="text-align: center">R${ticket_medio_dia:.2f}</td>
         <td style="text-align: center">R${meta_ticketmedio_dia:.2f}</td>
         <td style="text-align: center"><font color="{cor_ticket_dia}">◙</font></td>
         </tr>
      </table>
      <br>
      <table>
         <tr>
         <th>Indicador</th>
         <th>Valor Ano</th>
         <th>Meta Ano</th>
         <th>Cenário Ano</th>
         </tr>
         <tr>
         <td>Faturamento</td>
         <td style="text-align: center">R${faturamento_ano:.2f}</td>
         <td style="text-align: center">R${meta_faturamento_ano:.2f}</td>
         <td style="text-align: center"><font color="{cor_fat_ano}">◙</font></td>
         </tr>
         <tr>
         <td>Diversidade de Produtos</td>
         <td style="text-align: center">{qtde_produtos_ano}</td>
         <td style="text-align: center">{meta_qtdeprodutos_ano}</td>
         <td style="text-align: center"><font color="{cor_qtde_ano}">◙</font></td>
         </tr>
         <tr>
         <td>Ticket Médio</td>
         <td style="text-align: center">R${ticket_medio_ano:.2f}</td>
         <td style="text-align: center">R${meta_ticketmedio_ano:.2f}</td>
         <td style="text-align: center"><font color="{cor_ticket_ano}">◙</font></td>
         </tr>
      </table>

      <p>Segue em anexo a planilha com todos os dados para mais detalhes.</p>

      <p>Qualquer dúvida estou à disposição.</p>
      <p>Att., Arthur</p>
      '''
   # colocar anexos:
   attachment = Path.cwd() / caminho_backup / loja / f'{dia_indicador.month}_{dia_indicador.day}_{loja}.xlsx'
   print(attachment)
   mail.Attachments.Add(str(attachment))

   mail.Send()

   print('E-mail da Loja {} enviado'.format(loja))

   ### Passo 5 - Enviar por e-mail para o gerente
   ### Passo 6 - Automatizar todas as lojas
   ### Passo 7 - Criar ranking para diretoria
#### - Ao final, sua rotina deve enviar ainda um e-mail para a diretoria (informações também estão no arquivo Emails.xlsx) com 2 rankings das lojas em anexo, 1 ranking do dia e outro ranking anual. 
#### Além disso, no corpo do e-mail, deve ressaltar qual foi a melhor e a pior loja do dia e também a melhor e pior loja do ano. O ranking de uma loja é dado pelo faturamento da loja.


faturamento_lojas = vendas.groupby('Loja')[["Loja",'Valor Final']].sum()
faturamento_lojas_ano = faturamento_lojas.sort_values(by='Valor Final',ascending=False)
# display(faturamento_lojas_ano)

# salvando em excel O Ranking :
nome_arquivo = '{}_{}_Ranking Anual.xlsx'.format(dia_indicador.month, dia_indicador.day)
faturamento_lojas_ano.to_excel(r"C:\Users\01234\Downloads\Projeto AutomacaoIndicadores\Projeto AutomacaoIndicadores\Backup Arquivos Lojas\{}".format(nome_arquivo))

vendas_dia =  vendas.loc[vendas['Data']==dia_indicador , :] 
faturamento_lojas_dia= vendas_dia.groupby('Loja')[["Loja",'Valor Final']].sum()
faturamento_lojas_dia=faturamento_lojas_dia.sort_values(by='Valor Final',ascending=False)
# display(faturamento_lojas_dia)

nome_arquivo = '{}_{}_Ranking Dia.xlsx'.format(dia_indicador.month, dia_indicador.day)
faturamento_lojas_dia.to_excel(r"C:\Users\01234\Downloads\Projeto AutomacaoIndicadores\Projeto AutomacaoIndicadores\Backup Arquivos Lojas\{}".format(nome_arquivo))

### Passo 8 - Enviar e-mail para diretoria

outlook = win32.Dispatch('outlook.application')

mail = outlook.CreateItem(0) 
mail.To = emails.loc[emails['Loja']=="Diretoria", 'E-mail'].values[0]
mail.Subject = f'Ranking Dia {dia_indicador.day}/{dia_indicador.month} '
mail.Body = f'''
Prezados, bom dia

Melhor loja do Dia em Faturamento: Loja {faturamento_lojas_dia.index[0]} com Faturamento R${faturamento_lojas_dia.iloc[0, 0]:.2f}
Pior loja do Dia em Faturamento: Loja {faturamento_lojas_dia.index[-1]} com Faturamento R${faturamento_lojas_dia.iloc[-1, 0]:.2f}

Melhor loja do Ano em Faturamento: Loja {faturamento_lojas_ano.index[0]} com Faturamento R${faturamento_lojas_ano.iloc[0, 0]:.2f}
Pior loja do Ano em Faturamento: Loja {faturamento_lojas_ano.index[-1]} com Faturamento R${faturamento_lojas_ano.iloc[-1, 0]:.2f}

Segue em anexo os rankings do ano e do dia de todas as lojas.

Qualquer dúvida estou à disposição.

Att.,
Arthur
'''
# colocar anexos:
attachment = Path.cwd() / caminho_backup / f'{dia_indicador.month}_{dia_indicador.day}_Ranking Anual.xlsx'
mail.Attachments.Add(str(attachment))
print(str(attachment))
attachment = Path.cwd() / caminho_backup / f'{dia_indicador.month}_{dia_indicador.day}_Ranking Dia.xlsx'
mail.Attachments.Add(str(attachment))
print(str(attachment))

mail.Send()
print('E-mail da Diretoria enviado')