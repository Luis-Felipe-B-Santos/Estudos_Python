import pandas as pd
import os

simpsons = pd.read_html('https://en.wikipedia.org/wiki/List_of_The_Simpsons_episodes')
print(len(simpsons))
print(simpsons[1].head())


caminho_do_arquivo = os.path.join(os.path.dirname(__file__), '..', 'data', 'Episodios_Simpsons.xlsx')
caminho_do_arquivo = os.path.abspath(caminho_do_arquivo)

with pd.ExcelWriter(caminho_do_arquivo, engine = 'openpyxl') as writer:
    for i, tabela in enumerate(simpsons):
        nome_da_aba = f'Temporada_{i+1}'
        tabela.to_excel(writer, sheet_name=nome_da_aba)

print(f'Arquivo salvo em: {caminho_do_arquivo}')