import xml.etree.ElementTree as ET
import os
import csv


def leituraXML(caminho_xml):

    # Defina o namespace
    namespace = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}

    # Carregue o arquivo XML
    tree = ET.parse(caminho_xml)

    # Obtenha o elemento raiz
    root = tree.getroot()

    notafiscal= root.find('.//nfe:ide', namespace)

    # Use o namespace ao buscar elementos
    # Exemplo para extrair informações do emitente e do destinatário:
    emitente = root.find('.//nfe:emit', namespace)

    '''
    root.find(): Este método é usado para encontrar o primeiro subelemento que satisfaz um caminho especificado. O caminho é dado como um argumento para o método find().

    './/nfe:emit': Este é o caminho XPath que estamos usando para localizar o elemento <emit>. O . no início do caminho significa que estamos começando a busca a partir do elemento atual (que é o elemento raiz neste caso). Os dois pontos barra // significam que estamos procurando em todos os descendentes do elemento atual, não apenas nos filhos imediatos. nfe:emit significa que estamos procurando por elementos com a tag emit no namespace associado ao prefixo nfe, que foi definido no dicionário namespace.

    namespace: Este é o dicionário de namespaces que definimos anteriormente. No exemplo, ele foi definido como {'nfe': 'http://www.portalfiscal.inf.br/nfe'}.

    Portanto, o objetivo dessa linha é encontrar o primeiro elemento <emit> no documento XML e atribuir esse elemento à variável emitente. Depois disso, podemos usar a variável emitente para acessar os filhos e os atributos desse elemento e extrair informações específicas sobre o emitente da nota fiscal.

    '''
    if notafiscal is not None:
        num_nota_fiscal = notafiscal.find('nfe:nNF', namespace).text
        data_entrada_saida=notafiscal.find('nfe:dhSaiEnt', namespace).text

        


    if emitente is not None:
        nome_emitente = emitente.find('nfe:xNome', namespace).text
        #print(f'Nome do Emitente: {nome_emitente}')

    destinatario = root.find('.//nfe:dest', namespace)
    if destinatario is not None:
        nome_destinatario = destinatario.find('nfe:xNome', namespace).text
        #print(f'Nome do Destinatário: {nome_destinatario}')

    chave_nota_fiscal=root.find('.//nfe:protNFe/nfe:infProt', namespace)
    if chave_nota_fiscal is not None:
        chave_nfe = chave_nota_fiscal.find('nfe:chNFe', namespace).text    

    # Exemplo para extrair informações do produto:
    produtos = root.findall('.//nfe:det/nfe:prod', namespace)
    for produto in produtos:
        descricao = produto.find('nfe:xProd', namespace).text
        NCM = produto.find('nfe:NCM', namespace).text
        CFOP= produto.find('nfe:CFOP', namespace).text
        unidadeMedida= produto.find('nfe:uCom', namespace).text
        qtd= produto.find('nfe:vUnCom', namespace).text
        valorUnitario = produto.find('nfe:vUnCom',namespace).text
        valor = produto.find('nfe:vProd', namespace).text

        #print(f'Fornecedor: {nome_emitente}, Descrição do Produto: {descricao}, Valor unitário {valorUnitario}, Quantidade {qtd}, Und.Medida: {unidadeMedida}, Valor total: {valor}, CFOP: {CFOP}, NCM: {NCM}, NFE Num.: {num_nota_fiscal}, Data: {data_entrada_saida[:10]}')

    boletos = root.findall('.//nfe:cobr/nfe:dup', namespace)
    for duplicatas in boletos:
        num_duplicata= duplicatas.find('nfe:nDup', namespace).text
        vencimento= duplicatas.find('nfe:dVenc', namespace).text    
        valor_duplicata= duplicatas.find('nfe:vDup', namespace).text

        historico="Vr. Pagto ref. doc. "+num_duplicata+"-"+num_nota_fiscal

        writer.writerow([num_duplicata, vencimento, valor_duplicata, nome_emitente,num_nota_fiscal,historico,chave_nfe])
        #print(f'Número duplicata: {num_duplicata}, Valor: {valor_duplicata}, Vencimento: {vencimento}, Fornecedor:{nome_emitente}')    


def extensao_arquivo(caminho):
    return os.path.splitext(caminho)[1]


def ler_todos_arquivos_xml(diretorio):
    # Percorra todos os arquivos no diretório
    for raiz, subdiretorios, arquivos in os.walk(diretorio):
        for nome_arquivo in arquivos:
            caminho_completo = os.path.join(raiz, nome_arquivo)
            caminho_atualizado = caminho_completo.replace('\\', '\\\\')
            if ('.XML' in extensao_arquivo(caminho_completo).upper()):
                try:

                    leituraXML(caminho_completo)
                except:
                    pass

# Especifique o caminho do diretório que você deseja percorrer
diretorio = "C:\\Users\\Gabriel\\Desktop\\chromedriver_win32\\072023XMLFILIAL\\"
with open('C:\\Users\\Gabriel\\Desktop\\chromedriver_win32\\duplicatas.csv', mode='w', newline='', encoding='utf-8') as file:
    # Cria um objeto writer para escrever no arquivo CSV
    writer = csv.writer(file)
    writer.writerow(["PARCELAS", "VENCIMENTO", "VALOR DUPLICATA", "FORNECEDOR","NOTA FISCAL", "HISTORICO","CHAVE NOTA FISCAL"])
    ler_todos_arquivos_xml(diretorio)

   