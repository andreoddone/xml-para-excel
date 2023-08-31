import xmltodict
import os
import pandas as pd
""" import json """


def pegar_informacoes(nome_arquivo_xml):
    """ 
    Essa função define uma pasta onde está contido o arquivo XML que será convertido 
    para um dicionário do python. Depois são definidas as chaves relevantes que são 
    usadas para capturar valores presentes no dicionário. São usadas condições lógicas 
    if para dar conta da eventual ausência de determinadas chaves nos dicionários, e 
    também para dar conta de diferentes nomenclaturas que possam existir para uma mesma chave.
    O parâmetro "rb" da função open significa 'read byte', pois o parâmetro 'r' dessa 
    função retorna um objeto do tipo string, só que a função parse só funciona em objetos 
    do tipo bytes, por isso é usado o parâmetro 'rb' para gerar um objeto do tipo byte.
    
    :param nome_arquivo_xml: esse parâmetro indica o caminho para o arquivo XML que será convertido em dicionário
    """
    # print(f"\nAs informações serão coletadas do arquivo {nome_arquivo_xml}:\n") 
    with open(f"notas-fiscais/{nome_arquivo_xml}", "rb") as arquivo_xml:
        dicionario_arquivo_xml = xmltodict.parse(arquivo_xml)
        """ 
        Abaixo estão comentados os comandos try e except, que serviram em um primeiro momento
        para printar, durante a execução da função pegar_informacoes, qual chave o método não
        conseguiria encontrar. Usando essa estratégia foi possível criar condições para adequar
        o nome da chave que deveria ser buscada dentro do dicionário.
        """
        """ try: """
        if "NFe" in dicionario_arquivo_xml:
            informacoes_do_dicionario = dicionario_arquivo_xml["NFe"]["infNFe"]
        else:
            informacoes_do_dicionario = dicionario_arquivo_xml["nfeProc"]["NFe"]["infNFe"]        
        info_numero_nota = informacoes_do_dicionario["@Id"]
        info_nome_empresa_emissora = informacoes_do_dicionario["emit"]["xNome"]
        info_nome_empresa_destinataria = informacoes_do_dicionario["dest"]["xNome"]
        info_endereco_entrega = informacoes_do_dicionario["dest"]["enderDest"]
        if "vol" in informacoes_do_dicionario["transp"]:
            info_peso_produto = informacoes_do_dicionario["transp"]["vol"]["pesoB"]
        else:
            info_peso_produto = "O peso bruto não foi informado no arquivo."
        """
        A cada execução da função pegar_informacoes a linha abaixo faz um append à lista 
        linhas_tabela com uma lista que contém os valores coletados do dicionário. Posteriormente,
        cada lista inserida por esse append à lista linhas_tabela irá gerar uma linha em uma 
        tabela do excel.
        """
        linhas_tabela.append([info_numero_nota, info_nome_empresa_emissora, info_nome_empresa_destinataria, info_endereco_entrega, info_peso_produto])
        """ 
        except Exception as e:
            print(e)
            print(json.dumps(dicionario_arquivo_xml, indent=4))
        """
        """
            Essa linha acima usa uma função do módulo json para facilitar a visualização 
            do dicionário printado por meio da definição de um valor de identação.
        """       
            

""" 
A linha abaixo cria uma lista, usando a função listdir, com todos os arquivos que 
estão presentes na pasta que a função exige como parâmetro.
"""
lista_de_arquivos = os.listdir("notas-fiscais")

""" 
A variável coluna_tabela contém uma lista que armazena o título de colunas que comporão
uma tabela do excel. A ordem do elementos nessa lista irá influenciar na ordem das colunas
que serão criadas na tabela, e essa ordem deve condizer com a ordem dos valores capturados
pela função pegar_informacoes para serem inseridos na lista linhas_tabela.
"""
colunas_tabela = ["Número da nota fiscal", "Nome da empresa prestadora de serviço", "Nome da empresa tomadora de serviço", "Dados do endereço de entrega", "Peso bruto do produto transportado"]
linhas_tabela = []

"""
Essa estrutura de repetição serve para executar o método pegar_informacoes sobre os arquivos
que compõem a lista lista_de_arquivos.
"""
for arquivo in lista_de_arquivos:
    pegar_informacoes(arquivo)

"""
Essas linhas abaixo criam uma tabela com a biblioteca pandas, e essa tabela é convertida
para o formato excel com uma função que está recebendo dois parâmetros, o nome do arquivo 
que será criado e desobrigatoriedade de criar um índice para as linhas que serão inseridas
no arquivo.
"""
tabela = pd.DataFrame(columns = colunas_tabela, data = linhas_tabela)
tabela.to_excel("Notas-Fiscais.xlsx", index = False)
