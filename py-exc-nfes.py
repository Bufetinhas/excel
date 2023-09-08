## Esse código se propõe a receber informações em xml
## e passá-las para uma planilha já organizada no excel

import xmltodict
import os
import pandas as pd


def pegar_infos(nome_arquivo, valores): #Pegando as informações dos arquivos
    #print(f'pegou as informações {nome_arquivo}')
    with open(fr'C:\Users\PC\Desktop\douglas\projeto freela\nfs/{nome_arquivo}', 'rb') as arquivo_xml: 
#Estou abrindo o arquivo "nome_arquivo" e armazenando na variável "arquivo_xml"
        dic_arquivo = xmltodict.parse(arquivo_xml) # Transforma xml em um dicionário python


        if 'NFe' in dic_arquivo:
            infos_nf = dic_arquivo['NFe']['infNFe'] #esse infNFe tem todas as informações dentro
        else:
            infos_nf = dic_arquivo['nfeProc']['NFe']['infNFe']
        numero_nota = infos_nf["@Id"]
        empresa_emissora = infos_nf['emit']['xNome']
        nome_cliente = infos_nf['dest']['xNome']
        endereco = infos_nf['dest']['enderDest']
        if 'vol' in infos_nf['transp']:
            peso = infos_nf['transp']['vol']['pesoB']
        else:
            peso = 'Não informado'
        valores.append([numero_nota, empresa_emissora, nome_cliente, endereco,peso]) #Esse cara é a lista de valores


lista_arquivos = os.listdir(r'C:\Users\PC\Desktop\douglas\projeto freela\nfs') #Listar as informações dentro de um diretório

colunas = ['numero_nota','empresa_emissora', 'nome_cliente','endereco','peso']
valores = []


for arquivo in lista_arquivos: #Para cada arquivo na minha lista de arquivos, eu vou pegar as infos de cada
    pegar_infos(arquivo, valores)

tabela = pd.DataFrame(columns=colunas, data=valores)
tabela.to_excel('NotasFiscais.xlsx', index=False)
