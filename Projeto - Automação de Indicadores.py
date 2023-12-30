#!/usr/bin/env python
# coding: utf-8

# # Automação de Indicadores
# 
# ### Objetivo: Treinar e criar um Projeto Completo que envolva a automatização de um processo feito no computador
# 
# ### Descrição:
# 
# Imagine que você trabalha em uma grande rede de lojas de roupa com 25 lojas espalhadas por todo o Brasil.
# 
# Todo dia, pela manhã, a equipe de análise de dados calcula os chamados One Pages e envia para o gerente de cada loja o OnePage da sua loja, bem como todas as informações usadas no cálculo dos indicadores.
# 
# Um One Page é um resumo muito simples e direto ao ponto, usado pela equipe de gerência de loja para saber os principais indicadores de cada loja e permitir em 1 página (daí o nome OnePage) tanto a comparação entre diferentes lojas, quanto quais indicadores aquela loja conseguiu cumprir naquele dia ou não.

# O seu papel, como Analista de Dados, é conseguir criar um processo da forma mais automática possível para calcular o OnePage de cada loja e enviar um email para o gerente de cada loja com o seu OnePage no corpo do e-mail e também o arquivo completo com os dados da sua respectiva loja em anexo.
# 
# Ex: O e-mail a ser enviado para o Gerente da Loja A deve ser como exemplo

# ### Arquivos e Informações Importantes
# 
# - Arquivo Emails.xlsx com o nome, a loja e o e-mail de cada gerente. Obs: Sugiro substituir a coluna de e-mail de cada gerente por um e-mail seu, para você poder testar o resultado
# 
# - Arquivo Vendas.xlsx com as vendas de todas as lojas. Obs: Cada gerente só deve receber o OnePage e um arquivo em excel em anexo com as vendas da sua loja. As informações de outra loja não devem ser enviados ao gerente que não é daquela loja.
# 
# - Arquivo Lojas.csv com o nome de cada Loja
# 
# - Ao final, sua rotina deve enviar ainda um e-mail para a diretoria (informações também estão no arquivo Emails.xlsx) com 2 rankings das lojas em anexo, 1 ranking do dia e outro ranking anual. Além disso, no corpo do e-mail, deve ressaltar qual foi a melhor e a pior loja do dia e também a melhor e pior loja do ano. O ranking de uma loja é dado pelo faturamento da loja.
# 
# - As planilhas de cada loja devem ser salvas dentro da pasta da loja com a data da planilha, a fim de criar um histórico de backup
# 
# ### Indicadores do OnePage
# 
# - Faturamento -> Meta Ano: 1.650.000 / Meta Dia: 1000
# - Diversidade de Produtos (quantos produtos diferentes foram vendidos naquele período) -> Meta Ano: 120 / Meta Dia: 4
# - Ticket Médio por Venda -> Meta Ano: 500 / Meta Dia: 500
# 
# Obs: Cada indicador deve ser calculado no dia e no ano. O indicador do dia deve ser o do último dia disponível na planilha de Vendas (a data mais recente)
# 
# Obs2: Dica para o caracter do sinal verde e vermelho: pegue o caracter desse site (https://fsymbols.com/keyboard/windows/alt-codes/list/) e formate com html

# In[59]:


#Importação das bibliotecas
import pandas as pd
import pathlib
import win32com.client as win32


# In[60]:


#Importação de arquivos 
emails = pd.read_excel(r'Bases de Dados/Emails.xlsx')
lojas = pd.read_csv(r'Bases de Dados/Lojas.csv', sep = ';', encoding='latin1')
vendas = pd.read_excel(r'Bases de Dados/Vendas.xlsx')

display(emails)
display(lojas)
display(vendas)


# In[61]:


#Incluir nome da loja no dataframe vendas

vendas = vendas.merge(lojas, on='ID Loja')
display(vendas)


# In[62]:


dicionario_lojas = {}
for loja in lojas['Loja']:
    dicionario_lojas[loja] = vendas.loc[vendas['Loja']==loja, :]
display(dicionario_lojas['Iguatemi Esplanada'])


# In[63]:


#Identificando a data mais recente
dia_indicador = vendas['Data'].max()


# In[64]:


# Identificar existência de pasta de backup
caminho_backup = pathlib.Path(r'Backup Arquivos Lojas/')

arquivos_pasta_backup = caminho_backup.iterdir()

lista_nomes_backup = []
for arquivo in arquivos_pasta_backup:
    lista_nomes_backup.append(arquivo.name)
# Usando list comprehension: lista_nomes_backup = [arquivo.name for arquivo in arquivo_pasta_backup]

#Criação de pastas ainda não existentes
for loja in dicionario_lojas:
    if loja not in lista_nomes_backup:
        nova_pasta = caminho_backup / loja
        nova_pasta.mkdir()
        
    #Salvando cada arquivo dentro de sua respectiva pasta
    nome_arquivo = '{}_{}_{}.xlsx'.format(dia_indicador.month, dia_indicador.day, loja)
    local_arquivo = caminho_backup / loja / nome_arquivo
    dicionario_lojas[loja].to_excel(local_arquivo)


# In[65]:


#definição de metas

meta_faturamento_dia = 1000
meta_faturamento_ano = 1650000
meta_qtde_produtos_dia = 4
meta_qtde_produtos_ano = 120
meta_ticket_medio_dia = 500
meta_ticket_medio_ano = 500


# In[66]:


#calculando indicadores de Faturamento, Diversidade de Produtos, Ticket Médio
lista_indicadores = []
for loja in dicionario_lojas:
    vendas_loja_ult_dia = dicionario_lojas[loja].loc[dicionario_lojas[loja]["Data"]==dia_indicador,:]
                                                                           
    faturamento_ano = dicionario_lojas[loja]["Valor Final"].sum()
    faturamento_ult_dia= vendas_loja_ult_dia["Valor Final"].sum()
    diversidade_produtos_ano = len(dicionario_lojas[loja]['Produto'].unique())
    diversidade_produtos_ult_dia = len(vendas_loja_ult_dia['Produto'].unique())
    ticket_medio_ano = faturamento_ano / len(dicionario_lojas[loja]['Código Venda'].unique())
    if len(vendas_loja_ult_dia['Código Venda'].unique())>0:
        ticket_medio_dia = faturamento_ult_dia / len(vendas_loja_ult_dia['Código Venda'].unique())
    else:
        ticket_medio_dia = 0
        
    indicadores = [loja, faturamento_ano, faturamento_ult_dia, diversidade_produtos_ano, diversidade_produtos_ult_dia, ticket_medio_ano, ticket_medio_dia]
    lista_indicadores.append(indicadores)

    #Enviar e-mail com o apurado dos indicadores para o gerente responsável por cada shopping
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = emails.loc[emails['Loja'] == loja,'E-mail'].values[0]
    nome_gerente = emails.loc[emails['Loja'] == loja, 'Gerente'].values[0]
    mail.Subject = f'Indicadores de {dia_indicador.day}/{dia_indicador.month} - {loja}'
    mail.HTMLBody = f'''
    <p>Olá, {nome_gerente}</p>
    <p>O resultado de ontem <strong>({dia_indicador.day}/{dia_indicador.month})</strong> do <strong>{loja}</strong> foi:</p>
    <table>
        <tr>
            <th style="text-align:center;padding-left: 30px;padding-rightt: 30px">Indicador</th>
            <th style="text-align:center;padding-left: 20px;padding-rightt: 20px">Valor Dia</th>
            <th style="text-align:center;padding-left: 20px;padding-rightt: 20px">Meta Dia</th>
            <th style="text-align:center;padding-left: 20px;padding-rightt: 20px">Cenário Dia</th>
        </tr>
        <tr>
            <td style="text-align:center;padding-left: 30px;padding-rightt: 30px">Faturamento</td>
            <td style="text-align:center">R$ {faturamento_ult_dia}</td>
            <td style="text-align:center">R$ {meta_faturamento_dia:.2f}</td>
            <td style="text-align:center">{'<font color="green">▲</font>' if faturamento_ult_dia > meta_faturamento_dia else '<font color="red">▼</font>'}</td>
        </tr>
        <tr>
            <td style="text-align:center;padding-left: 30px;padding-rightt: 30px">Diversidade de Produtos</td>
            <td style="text-align:center">{diversidade_produtos_ult_dia}</td>
            <td style="text-align:center">{meta_qtde_produtos_dia}</td>
            <td style="text-align:center">{'<font color="green">▲</font>' if diversidade_produtos_ult_dia > meta_qtde_produtos_dia else '<font color="red">▼</font>'}</td>
        </tr>
        <tr>
            <td style="text-align:center;padding-left: 30px;padding-rightt: 30px">Ticket Médio</td>
            <td style="text-align:center">R$ {ticket_medio_dia:.2f}</td>
            <td style="text-align:center">R$ {meta_ticket_medio_dia:.2f}</td>
            <td style="text-align:center">{'<font color="green">▲</font>' if ticket_medio_dia > meta_ticket_medio_dia else '<font color="red">▼</font>'}</td>
        </tr>
    </table>
    <br>
    <p>Indicadores Anualizados: </p>
    <table>
        <tr>
            <th style="text-align:center;padding-left: 30px;padding-rightt: 30px">Indicador</th>
            <th style="text-align:center;padding-left: 20px;padding-rightt: 20px">Valor Dia</th>
            <th style="text-align:center;padding-left: 20px;padding-rightt: 20px">Meta Dia</th>
            <th style="text-align:center;padding-left: 20px;padding-rightt: 20px">Cenário Dia</th>
        </tr>
        <tr>
            <td style="text-align:center;padding-left: 30px;padding-rightt: 30px">Faturamento</td>
            <td style="text-align:center;padding-left: 20px;padding-rightt: 20px">R$ {faturamento_ano:.2f}</td>
            <td style="text-align:center;padding-left: 20px;padding-rightt: 20px">R$ {meta_faturamento_ano:.2f}</td>
            <td style="text-align:center;padding-left: 0px;padding-rightt: 20px">{'<font color="green">▲</font>' if faturamento_ano > meta_faturamento_ano else '<font color="red">▼</font>'}</td>
        </tr>
        <tr>
            <td style="text-align:center;padding-left: 30px;padding-rightt: 30px">Diversidade de Produtos</td>
            <td style="text-align:center">{diversidade_produtos_ano}</td>
            <td style="text-align:center">{meta_qtde_produtos_ano}</td>
            <td style="text-align:center">{'<font color="green">▲</font>' if diversidade_produtos_ano > meta_qtde_produtos_ano else '<font color="red">▼</font>'}</td>
        </tr>
        <tr>
            <td style="text-align:center;padding-left: 30px;padding-rightt: 30px">Ticket Médio</td>
            <td style="text-align:center">R$ {ticket_medio_ano:.2f}</td>
            <td style="text-align:center">R$ {meta_ticket_medio_ano:.2f}</td>
            <td style="text-align:center">{'<font color="green">▲</font>' if ticket_medio_ano > meta_ticket_medio_ano else '<font color="red">▼</font>'}</td>
        </tr>
    </table>
    <p>Segue  em anexo a planilha de com todos os dados para mais detalhes.</p>
    <p>Att. Célio Gomes</p>
    '''
    attachment = pathlib.Path.cwd()/local_arquivo
    mail.Attachments.Add(str(attachment))
    mail.Send()
df_indicadores = pd.DataFrame(lista_indicadores, columns=['Shopping', 'Fatur. Ano', 'Fatur. Últ. Dia', 'Diversidade Prod. Ano', 'Diversidade Prod. Últ Dia', 'Ticket Médio Ano', 'Ticket Médio Últ Dia'])
display(df_indicadores)


# In[67]:


#Criação dos Rankings de cada período (Anual e Diário)
df_faturamento_lojas_ano = vendas.groupby('Loja')["Valor Final"].sum()
df_faturamento_lojas_ano = df_faturamento_lojas_ano.sort_values(ascending=False)
display(df_faturamento_lojas_ano)

nome_arquivo = '{}_{}_Ranking Anual.xlsx'.format(dia_indicador.month, dia_indicador.day)
df_faturamento_lojas_ano.to_excel(r'Backup Arquivos Lojas\{}'.format(nome_arquivo))

df_faturamento_lojas_ult_dia = vendas.loc[vendas["Data"]==dia_indicador,:].groupby('Loja')['Valor Final'].sum()
df_faturamento_lojas_ult_dia = df_faturamento_lojas_ult_dia.sort_values(ascending = False)
display(df_faturamento_lojas_ult_dia)

nome_arquivo = '{}_{}_Ranking Diário.xlsx'.format(dia_indicador.month, dia_indicador.day)
df_faturamento_lojas_ult_dia.to_excel(r'Backup Arquivos Lojas\{}'.format(nome_arquivo))


# In[68]:


#Enviar e-mail para a diretoria
outlook =win32.Dispatch('outlook.application')

mail = outlook.CreateItem(0)
mail.To = emails.loc[emails['Loja']=="Diretoria", 'E-mail'].values[0]
mail.Subject = f'Ranking dia {dia_indicador.day}/{dia_indicador.month}'
mail.Body = f'''
Prezados,

Melhor loja do dia em Faturmaneto: {df_faturamento_lojas_ult_dia.index[0]} com Faturammento: R$ {df_faturamento_lojas_ult_dia.iloc[0]:.2f}
Pior loja do dia em Faturmaneto: {df_faturamento_lojas_ult_dia.index[-1]} com Faturammento: R$ {df_faturamento_lojas_ult_dia.iloc[-1]:.2f}

Melhor loja do dia em Faturmaneto: {df_faturamento_lojas_ano.index[0]} com Faturammento: R$ {df_faturamento_lojas_ano.iloc[0]:.2f}
Pior loja do dia em Faturmaneto: {df_faturamento_lojas_ano.index[-1]} com Faturammento: R$ {df_faturamento_lojas_ano.iloc[-1]:.2f}

Para mais detalhes, verificar os Rankings Anual e Diário da rede de lojas em anexo.

Atte,
Célio Gomes
'''
attachment = pathlib.Path.cwd() / caminho_backup / '{}_{}_Ranking Diário.xlsx'.format(dia_indicador.month, dia_indicador.day)
mail.Attachments.Add(str(attachment))
attachment = pathlib.Path.cwd() / caminho_backup / '{}_{}_Ranking Anual.xlsx'.format(dia_indicador.month, dia_indicador.day)
mail.Attachments.Add(str(attachment))
mail.Send()
print('E-mail da Diretoria Enviado')

