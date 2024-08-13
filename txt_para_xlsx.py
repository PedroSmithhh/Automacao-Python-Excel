import os
import pandas as pd

# Defina o diretório onde estão os arquivos .txt
diretorio_txt = r'caminho/para/seu/diretorio/txts'
diretorio_xlsx = r'caminho/para/seu/diretorio/xlsx'

# Garanta que o diretório de saída existe
os.makedirs(diretorio_xlsx, exist_ok=True)

# Liste todos os arquivos .txt no diretório
arquivos_txt = [f for f in os.listdir(diretorio_txt) if f.endswith('.txt')]

for txt_file in arquivos_txt:

    # Defina o caminho completo para o arquivo de entrada e saída
    caminho_arquivos_txt = os.path.join(diretorio_txt, txt_file)
    caminho_arquivos_xlsx = os.path.join(diretorio_xlsx, txt_file.replace('.txt', '.xlsx'))

    # Leia o arquivo .txt para um DataFrame
    df = pd.read_csv(caminho_arquivos_txt, delimiter='\t')  # Altere o delimitador conforme necessário

    # Salve o DataFrame como um arquivo .xlsx
    df.to_excel(caminho_arquivos_xlsx, index=False)

    print(f"Convertido: {txt_file} para {caminho_arquivos_xlsx}")

print("Conversão concluída!")