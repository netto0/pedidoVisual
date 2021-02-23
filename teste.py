from biblio import clientes



codigo = 5
cliente = clientes.itensArquivo(codigo)
nome_cliente = cliente[0]
cidade1 = cliente[1]
pagamento = cliente[2]
if pagamento.index('Bol') == True:
      print('vdd')
print(nome_cliente,
      cidade1,
      pagamento)