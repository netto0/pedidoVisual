def buscar_contato(contato):
    campo_pesquisa = driver.find_element_by_xpath('//div[contains(@class,"copyable-text selectable-text")]')
    time.sleep(3)
    campo_pesquisa.click()
    campo_pesquisa.send_keys(contato)
    campo_pesquisa.send_keys(Keys.ENTER)


def enviar_mensagem(msg):
    campo_mensagem = driver.find_elements_by_xpath('//div[contains(@class,"copyable-text selectable-text")]')
    campo_mensagem[1].click()
    time.sleep(3)
    campo_mensagem[1].send_keys(msg)
    #campo_mensagem[1].send_keys(Keys.ENTER)


def testar(funcao):
    while True:
        try:
            funcao
            break
        except:
            print('ERRO!')


# Definir contatos
contatos = ['Antônio Pena (Tabela1)', 'Gilvan Carlos Chagas', 'Ilma Carlos Chagas', 'Bena Damasceno', 'Renato (Tabela 2)',
             'Silvana Dainner (Tabela1)', 'Keila Pérola', 'Terezinha', 'Robinson (Tabela1)', 'Ailton Itanhém', 'Débora (Zé Baiano) Itanhém',
             'Neguinho Itanhém', 'Lidão Itanhem', 'Celso Itanhém', 'Reinaldo Itanhém', 'Rodomarck', 'Manoelton Itanhém (Semirames)',
             'Aldo Itanhém', 'Jailson Itupeva', 'Itupeva - Jaílton', 'Roberto Lajedão', 'Alas Medeiros Neto', 'Cal (Mercadinho Barão)', 'Edmo Medeiros Neto',
             'Jorvane Antônio Lima Medeiros Neto', 'Nélia (Panela Da Terra)', 'Ildeu Medeiros Neto', 'Núbia Medeiros Neto',
             'Silvio Medeiros Neto', 'Tiago Pardim Medeiros Neto', 'Douglas (Tabela2)', 'Da Hora (Tabela2)', 'Silvani', 'Raquel (Tabela2)',
             'Laiz Cunhada Silvani', 'André - São Jorge', 'Serra dos Aimorés Ademir Galdino', 'Davirley (Tabela2)', 'Janilson (Tabela2)',
            'Roberto - Tânia (Serra)', 'Dilson Gonzaga (Tabela2)', 'Railda', 'Fabiano Serra(Tabela2)', 'Rosilene (Tabela2)', 'Aline Vereda']
apelidos = ['Antônio', 'Gilvan', 'Ilma', 'Bena', 'Renato', 'Silvana', 'Keila', 'Terezinha', 'Robson', 'Aílton', 'Kinho',
            'Neguinho', 'Lidão', 'Celso', 'Reinaldo', 'Rodomarck', 'Manoelton', 'Aldo', 'Jaílson', 'Jaílton', 'Roberto', 'Alas',
            'Kau', 'Edmo', 'Jorvane', 'Nélia', 'Ildeu', 'Núbia', 'Sílvio', 'Tiago', 'Douglas', 'Da Hora', 'Silvani', 'Raquel', 'Laís', 'André', 'Ademir',
            'Davirlei', 'Janilson', 'Roberto', 'Dilson', 'Railda', 'Fabiano', 'Rosilene', 'Aline']
#contatos = ['Lis', 'WppAir☁']
#apelidos = ['Rodrigues', 'Grupo']

contatosSR = ['João Reta (Tabela2)', 'Hélcio Serra']
apelidosSR = ['Seu João', 'Seu Hélcio']

contatosSRA = ['Dona Aflorides Itanhém', 'Dona Maristela', 'Dona Maria (Udr)', 'Dona Branca']
apelidosSRA = ['Dona Aflorides', 'Dona Maristela', 'Dona Maria', 'Dona Branca']

nao_enviar = []

# preços do dia
while True:
    try:
        precobb = int(input('Preço Barbalho: R$ '))
        precogv = int(input('Preço Goval: R$ '))
        break
    except:
        print('ERRO! Digite um valor válido!')

while True:
    nm = str(input('Qual contato da lista não deverá receber a mensagem? '))
    if nm in nao_enviar:
        print(f'"{nm}" já está na lista de exclusão!')
    try:
        contatos.index(nm)
        nao_enviar.append(nm)
        print(f'{nm} adicionado à lista de exclusão. ')
    except:
        print(f'O contato "{nm}" não está na lista contatos. ')
        try:
            contatosSR.index(nm)
            nao_enviar.append(nm)
            print(f'{nm} adicionado à lista de exclusão. ')
        except:
            print(f'O contato "{nm}" não está na lista contatosSR. ')
            try:
                contatosSRA.index(nm)
                nao_enviar.append(nm)
                print(f'{nm} adicionado à lista de exclusão. ')
            except:
                print(f'O contato "{nm}" não está na lista contatosSRA. ')
    try:
        ask = str(input('Mais algum? [S/N]: ')).strip().upper()[0]
        if ask not in ('SN'):
            print('ERRO! Digite uma resposta válida!')
        elif ask in ('N'):
            print(nao_enviar)
            break
        elif ask in ('S'):
            continue
    except:
        print('ERRO! Digite uma resposta válida!')
    while True:
        try:
            cont = str(input('Esses contatos serão exluídos da lista, continuar? [S/N]: ')).strip().upper()[0]
            if cont not in ('SN'):
                print('ERRO! Digite uma resposta válida!')
            if cont in ('S'):
                break
        except:
            print('ERRO! Digite uma resposta válida!')
# noinspection PyUnreachableCode
for nome in contatos:
    if nome in nao_enviar:
        deletar = contatos.index(nome)
        apelidos.pop(deletar)
        contatos.pop(deletar)

for nome in contatosSR:
    if nome in nao_enviar:
        deletar = contatosSR.index(nome)
        apelidosSR.pop(deletar)
        contatosSR.pop(deletar)

for nome in contatosSRA:
    if nome in nao_enviar:
        deletar = contatosSRA.index(nome)
        apelidosSRA.pop(deletar)
        contatosSRA.pop(deletar)

# Importar bibliotecas
from selenium import webdriver
import time
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
# Navegar até o whatsapp web
driver = webdriver.Chrome(ChromeDriverManager().install())
driver.get('https://web.whatsapp.com/')
time.sleep(20)

# Contatos para não enviar

# Buscar contatos
contador = 0
for contato in contatos:
    msg = f'Bom dia {apelidos[contador]}, você vai precisar de feijão pra semana? Está R$ {precobb} o barbalho e R$ {precogv} o goval.'
    buscar_contato(contato)
    # Enviar mensagem
    enviar_mensagem(msg)
    contador += 1
contadorSR = 0
for contato in contatosSR:
    msgSR = f'Bom dia {apelidosSR[contadorSR]}, o senhor vai precisar de feijão pra semana? Está R$ {precobb} o barbalho e R$ {precogv} o goval.'
    buscar_contato(contato)
    # Enviar mensagem
    enviar_mensagem(msgSR)
    contadorSR += 1
contadorSRA = 0
for contato in contatosSRA:
    msgSRA = f'Bom dia {apelidosSRA[contadorSRA]}, a senhora vai precisar de feijão pra semana? Está R$ {precobb} o barbalho e R$ {precogv} o goval.'
    buscar_contato(contato)
    # Enviar mensagem
    enviar_mensagem(msgSRA)
    contadorSRA += 1

# campo de pesquisa : copyable-text selectable-text
# campo de msg privada: copyable-text selectable-text