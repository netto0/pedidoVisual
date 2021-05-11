data = '07/12/22'
def data_valida(data):
    lista = []
    verif = True
    mes = True
    try:
        if int(data[3:5]) <= 12:
            None
        else:
            mes = False
    except:
        print('Data inválida')
    for c in data:
        if c.isnumeric() == True:
            lista.append(c)
    for n in lista:
        if n.isnumeric() == True and len(lista) == 6 and mes:
            None
        else:
            verif = False

        return verif

print(data_valida(data))

"""for n in lista:
    if n.isnumeric() == False:
        return False"""