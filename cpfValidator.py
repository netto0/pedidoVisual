while True:
    cpf_cliente = input('Digite seu cpf(somente números): ')
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
            print('CPF Válido')
            break
        else:
            print('CPF Inválido')
            continue
    else:
        print('Digite um cpf válido')
        continue