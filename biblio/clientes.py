def cadastrar(arq, nome, cidade, pagam):
    try:
        a = open(arq, 'at')
    except:
        print('erro ao abrir arquivo')
    else:
        try:
            a.write(f'{nome};{cidade};{pagam}\n')
        except:
            print('Erro ao adicionar dados')
        else:
            print(f'Novo registro de {nome} adicionado.')
            a.close()

def lerArquivo(nome):
    try:
        a = open(nome, 'rt')
    except:
        print('erro ao ler arquivo')
    else:
        for linha in a:
            for item in linha:
                print(item)
    finally:
        a.close()

def itensArquivo(cod):
    import os
    diretorio = os.getcwd()
    caminho = f'{diretorio}\{"cadastrosclientes.txt"}'
    try:
        arq = open(fr'{caminho}')
    except:
        print('Arquivo não encontrado')
    linhas = arq.readlines()
    linha = linhas[cod].split(';')
    return linha

def linhasArquivo(arquivo):
    import os
    diretorio = os.getcwd()
    caminho = f'{diretorio}\{arquivo}'
    cont = 0
    try:
        arq = open(fr'{caminho}','rt')
    except:
        print('Arquivo não encontrado')
    else:
        for linha in arq:
            dado = linha.split(';')
            dado[1] = dado[1].replace('\n', '')
            print(f'{cont:<3} {dado[0]:<40} {dado[1]:<20} {dado[2]}')
            cont += 1

