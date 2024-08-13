import os
import pandas as pd

# Defina o diretório onde estão os arquivos .xlsx
diretorio_entrada = r'caminho\para\seu\diretorio\xlsx'
diretorio_saida = r'caminho\para\seu\diretorio\resultado\arquivo_combinado.xlsx'

# Liste todos os arquivos .xlsx no diretório
arquivos_xlsx = [f for f in os.listdir(diretorio_entrada) if f.endswith('.xlsx')]

# Crie uma lista para armazenar os DataFrames
dataframes = []

for xlsx_file in arquivos_xlsx:
    # Defina o caminho completo para o arquivo
    caminho_arquivos = os.path.join(diretorio_entrada, xlsx_file)

    # Leia o arquivo .xlsx para um DataFrame
    df = pd.read_excel(caminho_arquivos)

    # Adicione o DataFrame à lista
    dataframes.append(df)

# Combine todos os DataFrames em um só
uniao_arquivos = pd.concat(dataframes, ignore_index=True)

# Salve o DataFrame combinado em um novo arquivo .xlsx
uniao_arquivos.to_excel(diretorio_saida, index=False)

print(f"Arquivos combinados e salvos em: {diretorio_saida}")
