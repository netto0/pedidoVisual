"""CPF = 127.900.056-21

1 * 10 = 10
2 * 9 = 18
7 * 8 = 56
9 * 7 = 63
0 * 6 = 0
0 * 5 = 0
0 * 4 = 0
5 * 3 = 15
6 * 2 = 12

174
11 - (174 % 11) = 2

1 * 11 = 11
2 * 10 = 20
7 * 9 = 63
9 * 8 = 72
0 * 7 = 0
0 * 6 = 0
0 * 5 = 0
5 * 4 = 20
6 * 3 = 18
2 * 2 = 4

208
11 - (208 % 11) = 1

cpf = 12790005621

"""
from random import randint


def validador(cpf_cliente):
    vrf = cpf_cliente.isnumeric()
    if vrf and len(cpf_cliente) == 11:
        soma = 0
        soma2 = 0
        for c, i in enumerate(range(10,1,-1)):
            soma += int(cpf_cliente[c]) * int(i)
        dig1 = 11 - (soma % 11)
        for n, m in enumerate(range(11,1,-1)):
            soma2 += int(cpf_cliente[n]) * int(m)
        dig2 = 11 - (soma2 % 11)
        novo_cpf = str(cpf_cliente[0:9])+str(dig1)+str(dig2)
        if int(novo_cpf) == int(cpf_cliente):
            return True
        else:
            return False
    else:
        return False


while True:
    cpf_cliente = randint(100000000, 999999999)
    var = str(cpf_cliente)
    if len(var) == 9:
        soma = 0
        soma2 = 0
        for c, i in enumerate(range(9,1,-1)):
            soma += int(var[c]) * int(i)
        dig1 = 11 - (soma % 11)
        cpf_dig_1 = str(var[0:9])+str(dig1)
        for n, m in enumerate(range(11,1,-1)):
            soma2 += int(cpf_dig_1[n]) * int(m)
        dig2 = 11 - (soma2 % 11)
        novo_cpf = str(cpf_dig_1[0:9])+str(dig2)
        if validador(novo_cpf):
            print(f'CPF GERADO: {novo_cpf}')
            break
        else:
            continue
    else:
        print('Digite um cpf válido')
        continue