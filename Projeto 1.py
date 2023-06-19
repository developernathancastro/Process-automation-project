##Passo1 - Importar Arquivos e Bibliotecas

#importar bilbiotecas
import pandas as pd
import win32com.client as win32
import pathlib
from pathlib import Path

##Importar bases de dados
pd.set_option('display.max_columns', None)

emails = pd.read_excel(r'C:\Users\natha\Projeto 1 - Automações de Processo - Aplicação de Mercado de Trbalho\Bases de Dados\Emails.xlsx')

vendas = pd.read_excel(r'C:\Users\natha\Projeto 1 - Automações de Processo - Aplicação de Mercado de Trbalho\Bases de Dados\vendas.xlsx')

lojas = pd.read_csv(r'C:\Users\natha\Projeto 1 - Automações de Processo - Aplicação de Mercado de Trbalho\Bases de Dados\Lojas.csv', sep=';', encoding= 'iso-8859-1' )


df_consolidado = vendas.merge(lojas, on = 'ID Loja')
#print(df_consolidado)

##passo 2- definir, criar uma tabela para cada Loja e Definir o dia do Indicador

dicionario_lojas = {}
for loja in lojas['Loja']:
    dicionario_lojas[loja] = df_consolidado.loc[df_consolidado['Loja'] == loja, :]


dia_indicador = df_consolidado['Data'].max()
print('{}/{}'.format(dia_indicador.day, dia_indicador.month))


##passo 3 - Salvar a Planilha na pasta de backup

##identificar se a pasta já existe
caminho_backup = caminho = Path(r'C:/Users/natha/Projeto 1 - Automações de Processo - Aplicação de Mercado de Trbalho/Backup Arquivos Lojas')
arquivos_pasta = caminho.iterdir()

lista_nomes_backup = [arquivo.name for arquivo in arquivos_pasta]

for loja in dicionario_lojas:
    if loja  not in lista_nomes_backup:
        nova_pasta = caminho_backup / loja
        nova_pasta.mkdir()

###salvar dentro da pasta
    nome_arquivo = '{}_{}_{}.xlsx'.format(dia_indicador.day, dia_indicador.month, loja)
    local_arquivo = caminho_backup / loja / nome_arquivo
    dicionario_lojas[loja].to_excel(local_arquivo)

##definicao de meta
meta_faturamento_dia = 1000
meta_faturamento_ano = 1650000
meta_qtdeprodutos_dia = 4
meta_qtdeprodutos_ano = 120
meta_ticket_medio_dia = 500
meta_ticket_medio_ano = 500

##passo 4- calcula o indicador para 1 loja

for loja in dicionario_lojas:
    vendas_loja = dicionario_lojas[loja]
    vendas_loja_dia = vendas_loja.loc[vendas_loja['Data'] == dia_indicador, :]

    ##faturamento

    faturamento_ano = vendas_loja['Valor Final'].sum()
    faturamento_dia = vendas_loja_dia['Valor Final'].sum()



    ##diversidade de produtos
    qtde_produtos_ano = len(vendas_loja['Produto'].unique())
    qtde_produtos_dia = len(vendas_loja_dia['Produto'].unique())


    ##ticket médio
    valor_venda = vendas_loja.groupby('Código Venda').sum(numeric_only=True)
    ticket_medio_ano = valor_venda['Valor Final'].mean()

    #ticket médio dia
    valor_venda_dia = vendas_loja_dia.groupby('Código Venda').sum(numeric_only=True)
    ticket_medio_dia = valor_venda_dia['Valor Final'].mean()


    ##passo 5 - enviar por e-mail para o gerente

    outlook = win32.Dispatch('outlook.application')

    nome = emails.loc[emails['Loja'] == loja, 'Gerente'].values[0]
    mail = outlook.CreateItem(0)
    mail.To = emails.loc[emails['Loja'] == loja, 'E-mail'].values[0]

    mail.Subject = f'One Page dia {dia_indicador.day}/{dia_indicador.month} - Loja {loja}'

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

    if ticket_medio_dia >= meta_ticket_medio_dia:
        cor_ticket_dia = 'green'
    else:
        cor_ticket_dia = 'red'

    if ticket_medio_ano >= meta_ticket_medio_ano:
        cor_ticket_ano = 'green'
    else:
        cor_ticket_ano = 'red'

    mail.HTMLBody = f'''
    <p>Bom dia, {nome}</p>
    
    <p> O resultado de ontem <strong>({dia_indicador.day}/{dia_indicador.month})</strong> da <strong> Loja <{loja}</strong> Campinas foi: </p>
    
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
            <td style="text-align: center">R${meta_ticket_medio_dia:.2f}</td>
            <td style="text-align: center"><font color="{cor_ticket_dia}">◙</font></td>
          </tr>
    </table>
    <br>
    <table>
          <tr>
            <th>Indicador</th>
            <th>Valor Ano </th>
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
            <td style="text-align: center">R${meta_ticket_medio_ano:.2f}</td>
            <td style="text-align: center"><font color="{cor_ticket_ano}">◙</font></td>
          </tr>
        </table>
        
    <p> Segue em anexo a planilha com todos os dados para mais detalhes. </p>
    
    <p>Qualquer dúvida estou à disposição.</p> 
    <p> Atenciosamente, Nathan</p>
     '''

    attachment = caminho_backup / loja / f'{dia_indicador.day}_{dia_indicador.month}_{loja}.xlsx'
    mail.Attachments.Add(str(attachment))

    mail.Send()
    print(f'Email da Loja {loja} enviado.')

#Passo 7 - criar ranking para diretoria

faturamento_lojas = df_consolidado.groupby('Loja')[['Loja', 'Valor Final']].sum(numeric_only=True)
faturamento_lojas_ano = faturamento_lojas.sort_values(by='Valor Final', ascending=False)

nome_arquivo = '{}_{}_Ranking_Anual.xlsx'.format(dia_indicador.day, dia_indicador.month)
local_arquivo = caminho_backup / nome_arquivo
faturamento_lojas_ano.to_excel(local_arquivo)


vendas_dia = df_consolidado.loc[df_consolidado['Data'] == dia_indicador, :]
faturamento_lojas_dia = df_consolidado.groupby('Loja')[['Loja', 'Valor Final']].sum(numeric_only=True)
faturamento_lojas_dia = faturamento_lojas_dia.sort_values(by='Valor Final', ascending=False)

nome_arquivo = '{}_{}_Ranking_Dia.xlsx'.format(dia_indicador.day, dia_indicador.month)
local_arquivo = caminho_backup / nome_arquivo
faturamento_lojas_dia.to_excel(local_arquivo)

##Passo 8 - Enviar E-mail para diretoria


outlook = win32.Dispatch('outlook.application')

mail = outlook.CreateItem(0)
mail.To = emails.loc[emails['Loja'] == 'Diretoria', 'E-mail'].values[0]

mail.Subject = f'Ranking dia {dia_indicador.day}/{dia_indicador.month}'
mail.body = f'''
Prezados, bom dia

Melhor loja do Dia em Faturamento: Loja {faturamento_lojas_dia.index[0]} com Faturamento R${faturamento_lojas_dia.iloc[0, 0]:.2f}.
Pior loja do Dia em Faturamento: Loja {faturamento_lojas_dia.index[-1]} com Faturamento R${faturamento_lojas_dia.iloc[-1, 0]:.2f}.

Melhor loja do Ano em Faturamento: Loja{faturamento_lojas_ano.index[0]} com Faturamento R${faturamento_lojas_ano.iloc[0, 0]:.2f}.
Pior loja do Ano em Faturamento: Loja {faturamento_lojas_ano.index[-1]} com Faturamento R${faturamento_lojas_ano.iloc[-1, 0]:.2f}. 

Segue em anexo os rankings do ano e do dia de todas as lojas.

Qualquer dúvida estou à disposição. 
Atenciosamente, Nathan Castro.
'''

attachment = caminho_backup / f'{dia_indicador.day}_{dia_indicador.month}_Ranking_Anual.xlsx'
mail.Attachments.Add(str(attachment))
attachment = caminho_backup / f'{dia_indicador.day}_{dia_indicador.month}_Ranking_Dia.xlsx'
mail.Attachments.Add(str(attachment))

mail.Send()
print('Email da Diretoria enviado.')

























