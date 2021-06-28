def cadastrar(arq, razao, nome, cidade, forma, prazo):
    try:
        a = open(arq, 'at')
    except:
        print('erro ao abrir arquivo')
    else:
        try:
            a.write(f'{razao};{nome};{cidade};{forma},{prazo}\n')
        except:
            print('Erro ao adicionar dados')
        else:
            print(f'Novo registro de {razao} adicionado.')
            a.close()


class Pessoa:
    def __init__(self,nome,razao,prazo,forma,cidade):
        self.nome = nome
        self.razao = razao
        self.prazo = prazo
        self.forma = forma
        self.cidade = cidade
    pass


# cadastrar("cadastrosclientes.txt",'Kleber da Silva Pereira','Kleber','Carlos Chagas','Ch.','28')


"""cont = 1
while True:
    nome = input('Digite o nome da pessoa: ')
    razao = input('Digite a razão: ')
    prazo = input('Prazo: ')
    forma = input('Forma de Pagamento: ')
    cidade = input('Cidade: ')
    pessoa2 = Pessoa(nome,razao,prazo,forma,cidade)
    cont += 1
    print(pessoa2.nome,pessoa2.prazo)
"""
print(len('Pérola do Mucuri Sup. Dist. Alim.'))