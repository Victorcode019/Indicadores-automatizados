# importando bibliotecas 
import pandas as pd
import pathlib
import win32com.client as win32
import os

# localizando o diretório atual
diretorio_atual = os.path.dirname(__file__)
print(diretorio_atual)

# definindo caminho das pastas 
vendas = pd.read_excel(os.path.join(diretorio_atual, "Bases de Dados", "Vendas.xlsx"))
lojas = pd.read_csv(os.path.join(diretorio_atual, "Bases de Dados", "Lojas.csv"), sep=";", encoding="latin1")
emails = pd.read_excel(os.path.join(diretorio_atual, "Bases de Dados", "Emails.xlsx"))
vendas = vendas.merge(lojas, on="ID Loja")

# Criar uma Tabela para cada Loja e Definir o dia do Indicador
dic_lojas = {}
for loja in lojas['Loja']:
    dic_lojas[loja] = vendas.loc[vendas['Loja']==loja, :]

# Definindo o dia do indicador
dia_indicador = vendas["Data"].max()
# print(f"Dia do indicador: {dia_indicador}")

# Salvar a planilha na pasta de backup
    # identificar se a pasta já existe 
caminho_backup = pathlib.Path(r'Backup Arquivos Lojas')
caminho_backup.mkdir(parents=True, exist_ok=True)
arquivos_pasta_backup = caminho_backup.iterdir()
lista_nomes_backup = [arquivo.name for arquivo in arquivos_pasta_backup]

# Criar uma nova pasta para a loja, se não existir
for loja in dic_lojas:
    if loja not in lista_nomes_backup:
        nova_pasta = caminho_backup / loja
        nova_pasta.mkdir(parents=True, exist_ok=True)

    # salvar dentro da pasta
    nome_arquivo = f'{dia_indicador.month}_{dia_indicador.day}_{loja}.xlsx'
    local_arquivo = nova_pasta / nome_arquivo
    # print(local_arquivo)
    # print(nova_pasta.exists())
    dic_lojas[loja].to_excel(local_arquivo)
    # print(os.listdir(nova_pasta))

# definiçao de meta  
meta_faturamento_dia = 1000
meta_faturamento_ano = 1650000
meta_qtde_produtos_dia = 4
meta_qtde_produtos_ano = 120
meta_ticket_medio_dia = 500
meta_ticket_medio_ano = 500

# Análise de resultados e envio de e-mails
for loja in dic_lojas:
    vendas_loja = dic_lojas[loja]
    vendas_loja_dia = vendas_loja.loc[vendas_loja['Data']==dia_indicador, :]

    #faturamento
    faturamento_ano = vendas_loja['Valor Final'].sum()
    #print(f'Faturamento do ano: {faturamento_ano}')
    faturamento_dia = vendas_loja_dia['Valor Final'].sum()
    #print(f'Faturamento do dia: {faturamento_dia}')

    #diversidade de produtos
    qtde_produtos_ano = vendas_loja['Produto'].unique()
    #print(f'Quantidade de produtos vendidos no ano: {len(qtde_produtos_ano)}')
    qtde_produtos_dia = vendas_loja_dia['Produto'].unique()
    #print(f'Quantidade de produtos vendidos no dia: {len(qtde_produtos_dia)}')

    # ticket médio 
    valor_venda = vendas_loja.groupby('Código Venda').sum(numeric_only=True)
    ticket_medio_ano = valor_venda['Valor Final'].mean()
    #print(ticket_medio_ano)
    # ticket médio dia
    valor_medio_dia = vendas_loja_dia.groupby('Código Venda').sum(numeric_only=True)
    ticket_medio_dia = valor_medio_dia['Valor Final'].mean()
    #print(ticket_medio_dia)

    # enviar o email
    outlook = win32.Dispatch("outlook.application")

    nome = emails.loc[emails['Loja']==loja, 'Gerente'].values[0]
    mail = outlook.CreateItem(0)
    mail.To = emails.loc[emails['Loja']==loja, 'E-mail'].values[0]
    mail.Subject = f"OnePage Dia {dia_indicador.day}/{dia_indicador.month} - Loja {loja}"

    if faturamento_dia >= meta_faturamento_dia:
        cor_fat_dia = 'green'
    else:
        cor_fat_dia = 'red'

    if faturamento_ano >= meta_faturamento_ano:
        cor_fat_ano = 'green'
    else:
        cor_fat_ano = 'red'

    if qtde_produtos_dia >= meta_qtde_produtos_dia:
        cor_qtde_dia = 'green'
    else:
        cor_qtde_dia = 'red'

    if qtde_produtos_ano >= meta_qtde_produtos_ano:
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

    # mail.Body = "Texto do email"
    mail.HTMLBody = f"""
    <p>Bom Dia, {nome}</p> 

    <p>O resultado de ontem <strong>({dia_indicador.day}/{dia_indicador.month})</strong> da Loja <strong>{loja}</strong> foi </p>

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
        <td style="text-align: center">{meta_qtde_produtos_dia}</td>
        <td style="text-align: center"><font color="{cor_qtde_dia}">◙</font>/td>
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
        <th>Valor Dia</th>
        <th>Meta Dia</th>
        <th>Cenário Dia</th>
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
        <td style="text-align: center">{meta_qtde_produtos_ano}</td>
        <td style="text-align: center"><font color="{cor_qtde_ano}">◙</font></td>
    </tr>
    <tr>
        <td>Ticket Médio</td>
        <td style="text-align: center">R${ticket_medio_ano:.2f}</td>
        <td style="text-align: center">R${meta_ticket_medio_ano:.2f}</td>
        <td style="text-align: center"><font color="{cor_ticket_ano}">◙</font></td>
    </tr>
    </table>
    <p>Segue em anexo a planilha com todos os dados para mais detalhes.</p>

    <p>Qualquer dúvida estou à disposição.</p>
    <p>Att., Victor</p>
    """

    # anexos
    attachment = pathlib.Path.cwd() / caminho_backup / loja / f'{dia_indicador.month}_{dia_indicador.day}_{loja}.xlsx'
    mail.Attachments.Add(str(attachment))

    mail.Send()

# criar ranking para a diretoria
faturamento_lojas = vendas.groupby('Loja')[['Valor Final', 'Loja']].sum(numeric_only=True)
faturamento_lojas_ano = faturamento_lojas.sort_values(by='Valor Final', ascending=False)
# display(faturamento_lojas_ano)
nome_arquivo = '{}_{}_Ranking Anual.xlsx'.format(dia_indicador.month, dia_indicador.day)
faturamento_lojas_ano.to_excel(r"Backup Arquivos Lojas\{}".format(nome_arquivo))
vendas_dia = vendas.loc[vendas['Data'] == dia_indicador, :]
faturamento_lojas_dia = vendas_dia.groupby('Loja')[['Valor Final', 'Loja']].sum(numeric_only=True)
faturamento_lojas_dia = faturamento_lojas_dia.sort_values(by='Valor Final', ascending=False)
# display(faturamento_lojas_dia)
nome_arquivo = '{}_{}_Ranking Diário.xlsx'.format(dia_indicador.month, dia_indicador.day)
faturamento_lojas_dia.to_excel(r"Backup Arquivos Lojas\{}".format(nome_arquivo))

# enviar o email
outlook = win32.Dispatch("outlook.application")

mail = outlook.CreateItem(0)
mail.To = emails.loc[emails['Loja']=='Diretoria', 'E-mail'].values[0]
mail.Subject = f"Ranking Dia {dia_indicador.day}/{dia_indicador.month}"
mail.Body = f'''
Prezados, Bom dia

Melhor loja do dia em Faturamento: Loja {faturamento_lojas_dia.index[0]} com Faturamento R${faturamento_lojas_dia.iloc[0,0]:.2f}
Pior loja do dia em Faturamento: Loja {faturamento_lojas_dia.index[-1]} com Faturamento R${faturamento_lojas_dia.iloc[-1,0]:.2f}

Melhor loja do ano em Faturamento: Loja {faturamento_lojas_ano.index[0]} com Faturamento R${faturamento_lojas_ano.iloc[0,0]:.2f}
Pior loja do ano em Faturamento: Loja {faturamento_lojas_ano.index[-1]} com Faturamento R${faturamento_lojas_ano.iloc[-1,0]:.2f}

Segue em anexo os rankings do ano e do dia de todas as lojas.

Qualquer dúvida, estou à disposição.

Att.,
Victor
'''

# anexos
attachment = pathlib.Path.cwd() / caminho_backup / f'{dia_indicador.month}_{dia_indicador.day}_Ranking Anual.xlsx'
mail.Attachments.Add(str(attachment))
attachment = pathlib.Path.cwd() / caminho_backup / f'{dia_indicador.month}_{dia_indicador.day}_Ranking Dia.xlsx'
mail.Attachments.Add(str(attachment))

mail.Send()
# print("Email da diretora enviado")

# fim do script