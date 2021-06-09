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
    # campo_mensagem[1].send_keys(Keys.ENTER)

# preços do dia
precobb = 200
precogv = 190
horario = 'Bom dia'
# horario = 'Boa tarde'
# Definir contatos

contatos = [{'nome': 'Antônio', 'contato': 'Antônio Pena (Tabela1)'},
            {'nome': 'Gilvan', 'contato': 'Gilvan Carlos Chagas'},
            {'nome': 'Ilma', 'contato': 'Ilma Carlos Chagas'},
            {'nome': 'Bena', 'contato': 'Bena Damasceno'},
            {'nome': 'Renato', 'contato': 'Renato (Tabela 2)'},
            {'nome': 'Silvana', 'contato': 'Silvana Dainner (Tabela1)'},
            {'nome': 'Paula', 'contato': 'Paula Pérola (Tabela1)'},
            {'nome': 'Terezinha', 'contato': 'Terezinha'},
            {'nome': 'Robson', 'contato': 'Robinson (Tabela1)'},
            {'nome': 'Aílton', 'contato': 'Ailton Itanhém'},
            {'nome': 'Kinho', 'contato': 'Débora (Zé Baiano) Itanhém'},
            {'nome': 'Neguinho', 'contato': 'Neguinho Itanhém'},
            {'nome': 'Lidão', 'contato': 'Lidão Itanhem'},
            {'nome': 'Celso', 'contato': 'Celso Itanhém'},
            {'nome': 'Aldo', 'contato': 'Aldo-Itanhém'},
            {'nome': 'Rodomarck', 'contato': 'Rodomarck'},
            {'nome': 'Manoelton', 'contato': 'Manoelton Itanhém (Semirames)'},
            {'nome': 'Reinaldo', 'contato': 'Reinaldo Itanhém'},
            {'nome': 'Jaílson', 'contato': 'Jailson Itupeva'},
            {'nome': 'Jaílton', 'contato': 'Itupeva - Jaílton'},
            {'nome': 'Roberto', 'contato': 'Roberto Lajedão'},
            {'nome': 'Alas', 'contato': 'Alas Medeiros Neto'},
            {'nome': 'Kau', 'contato': 'Cal (Mercadinho Barão)'},
            {'nome': 'Edmo', 'contato': 'Edmo Medeiros Neto'},
            {'nome': 'Jorvane', 'contato': 'Jorvane Antônio Lima Medeiros Neto'},
            {'nome': 'Nélia', 'contato': 'Nélia (Panela Da Terra)'},
            {'nome': 'Ildeu', 'contato': 'Ildeu Medeiros Neto'},
            {'nome': 'Núbia', 'contato': 'Núbia Medeiros Neto'},
            {'nome': 'Sílvio', 'contato': 'Silvio Medeiros Neto'},
            {'nome': 'Tiago', 'contato': 'Tiago Pardim Medeiros Neto'},
            {'nome': 'Douglas', 'contato': 'Douglas (Tabela2)'},
            {'nome': 'Da Hora', 'contato': 'Da Hora (Tabela2)'},
            {'nome': 'Silvani', 'contato': 'Silvani'},
            {'nome': 'Raquel', 'contato': 'Raquel (Tabela2)'},
            {'nome': 'Laís', 'contato': 'Laís'},
            {'nome': 'Viviane', 'contato': 'Viviane- São Jorge'},
            {'nome': 'Ademir', 'contato': 'Serra dos Aimorés Ademir Galdino'},
            {'nome': 'Davirlei', 'contato': 'Davirley (Tabela2)'},
            {'nome': 'Janilson', 'contato': 'Janilson (Tabela2)'},
            {'nome': 'Tânia', 'contato': 'Tânia (Serra)'},
            {'nome': 'Dilson', 'contato': 'Dilson Gonzaga (Tabela2)'},
            {'nome': 'Railda', 'contato': 'Railda'},
            {'nome': 'Fabiano', 'contato': 'Fabiano Serra(Tabela2)'},
            {'nome': 'Rosilene', 'contato': 'Rosilene (Tabela2)'},
            {'nome': 'Seu João', 'contato': 'João Reta (Tabela2)', 'tratamento': 'o senhor'},
            {'nome': 'Seu Hélcio', 'contato': 'Hélcio Serra', 'tratamento': 'o senhor'},
            {'nome': 'Dona Maristela', 'contato': 'Dona Maristela', 'tratamento': 'a senhora'},
            {'nome': 'Dona Maria', 'contato': 'Dona Maria (Udr)', 'tratamento': 'a senhora'},
            {'nome': 'Dona Branca', 'contato': 'Dona Branca', 'tratamento': 'a senhora'}]

# Importar bibliotecas
from selenium import webdriver
import time
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
# Navegar até o whatsapp web
driver = webdriver.Chrome(ChromeDriverManager().install())
driver.get('https://web.whatsapp.com/')
time.sleep(20)

# Buscar contatos
contador = 0
for contato in contatos:
    try:
        msg = f'{horario} {contatos[contador]["nome"]}, {contatos[contador]["tratamento"]} vai precisar de feijão pra semana? Está R$ {precobb} o barbalho e R$ {precogv} o goval.'
    except:
        msg = f'{horario} {contatos[contador]["nome"]}, você vai precisar de feijão pra semana? Está R$ {precobb} o barbalho e R$ {precogv} o goval.'
        pass
    buscar_contato(contatos[contador]["contato"])
    # Enviar mensagem
    enviar_mensagem(msg)
    contador += 1

# campo de pesquisa : copyable-text selectable-text
# campo de msg privada: copyable-text selectable-text

