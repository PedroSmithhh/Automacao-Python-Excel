import openpyxl as open
import re
import os

# Carrega o excel
excel = open.load_workbook(r"Caminho_Arquivo") # Coloque o caminho em que seu documento excel se encontra no seu computador dentro dos parenteses

# Nome da planilha
planilha = excel['Nome_Planilha'] # Escreva aqui o nome da planilha que voce quer manipular

def limpar_celula(valor):
    if isinstance(valor, str):
        return re.sub(r'\D', '', valor)
    return str(valor)


def tira_55(valor):
    if valor.startswith('55'):
        return valor[2:]
    return valor

def coloca_hifen(valor):
    if valor[:-4] != '-':
        valor_list = list(valor)
        valor_list.insert(-4, '-')
        valor_string = ''.join(valor_list)
        return valor_string
    return valor

def espaço_DDD(valor):
    valor = valor.lstrip('0')
    if valor[2:] != ' ':
        valor_list = list(valor)
        valor_list.insert(2, ' ')
        valor_string = ''.join(valor_list)
        return valor_string
    return valor

def padrao_numeros(valor):
    padrao = re.compile(r'^\d{2} \d{4,5}-\d{4}$') # Expressão regular para os formatos aceitos (** ****-**** ou ** *****-****)
    return padrao.match(valor)

def caracteres_iguais(valor):
    for caracter in valor:
        if valor[0] != caracter:
            return False
    return True

print('Bem vindo ao formnatador de planilhas versão 1.0.0')
      
while True:
    opcao_criar_alterar = input('Voce deseja [C]riar uma nova planilha com seus dados ou [A]lterar sua antiga com eles? ')

    if opcao_criar_alterar.lower() == 'c' or opcao_criar_alterar == 'a':
        break
    else:
        os.system('cls')
        print('Por favor digite um valor valido (A -> Alterar | C -> Criar)')

# Encontrar a coluna com o cabeçalho 'Telefone'
coluna_telefone = None
for celula in planilha[1]: # Como se trata de cabeçalho, assume-se que este se encontra na primeira linha
    if celula.value == 'Telefone' or celula.value == 'telefone':
        coluna_telefone = celula.column
        break

# Caso nenhuma coluna com o cabeçalho 'Telefone' seja encontrada
if coluna_telefone is None:
    raise ValueError("Coluna 'Telefone' não encontrada.")

os.system('cls')
print('Processando suas informações, isso pode levar alguns segundos...')

linha_errada = []
numeros_mostrados = set()
for linha in planilha.iter_rows(min_col= coluna_telefone, max_col= coluna_telefone, min_row=2):
    for celula in linha:
        celula.value = limpar_celula(celula.value)
        celula.value = tira_55(celula.value)

        # Verifica a ocorrencia de numeros que possuem todos seus caracteres iguais, o que caracteriza um numero inexistente, e se houver é adicionado à lista linha errada
        if caracteres_iguais(celula.value):
            linha_errada.append(celula.row)

        celula.value = coloca_hifen(celula.value)
        celula.value = espaço_DDD(celula.value)
        
        # Verifica se o valor, após todas as alterações esta conforme o padrao esperado, se nao estiver coloca na lista de linhas erradas
        if not padrao_numeros(celula.value):
            linha_errada.append(celula.row)
        
        # Apaga valores duplicados a partir da segunda ocorrencia, se houver, os coloca na lista de linhas erradas
        if celula.value in numeros_mostrados:
            linha_errada.append(celula.row)
        else:
            numeros_mostrados.add(celula.value)

# Excluir linhas erradas, começando da última para evitar problemas de deslocamento
for linha in reversed(linha_errada):
    planilha.delete_rows(linha)

if opcao_criar_alterar.lower() == 'c':
    # Cria uma nova planilha com as alterações
    excel.save(r"Nome_Arquivo/Novo_nome") # Coloque o caminho onde voce deseja salvar seu novo arquivo junto com seu novo nome

if opcao_criar_alterar.lower() == 'a':
    # Altera a planilha manipulada com os novos dados
    excel.save(r"Caminho_Arquivo") # Coloque o caminho do arquivo que voce esta usando

os.system('cls')
print('Suas informações foram processadas com sucesso')