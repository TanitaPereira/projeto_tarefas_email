import pandas as pd
tabela = pd.read_excel('vendas.xlsx') # caso contrario tem que especificar /cola o caminho do windows (inverte a barra)
print(tabela['Vendas'].sum())
print(tabela[tabela['Produto']=='Camisa']['Vendas'].sum())

nova_linha = {
    'Nome' : 'Juca',
    'Vendas' : 3900,
    'Produto' : 'Camisa'
}

tabela =tabela._append(nova_linha, ignore_index=True)

tabela.to_excel('vendas_atualizadas.xlsx', index=False)

print(tabela[tabela['Produto']=='Camisa']['Vendas'].sum())