from biblio import clientes
arquivo = 'cadastrosclientes.txt'
while True:
    print('='*30)
    print('NOVO CADASTRO'.center(30))
    print('='*30)
    nome = input('Nome: ')
    #endereço = input('Endereço: ')
    cidade = input('Cidade: ')
    #telefone = input('Fone: ')
    #cnpj = input('CNPJ: ')
    #ie = input('Inscrição estadual: ')
    pagam = input('Forma de Pagamento: ')
    #mail = input('E-mail: ')
    clientes.cadastrar(arquivo,nome,cidade,pagam)
    try:
        perg = str(input('Iniciar um novo cadastro?: [S/N]')).upper().strip()[0]
        if perg in ('N'):
            break
    except:
        continue
