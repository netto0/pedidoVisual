import PySimpleGUI as sg
precoBB = 190
precoGV = 180

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
            {'nome': 'Aldo', 'contato': 'Aldo Itanhém'},
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
            {'nome': 'Laís', 'contato': 'Laiz Cunhada Silvani'},
            {'nome': 'Viviane', 'contato': 'Viviane- São Jorge'},
            {'nome': 'Ademir', 'contato': 'Serra dos Aimorés Ademir Galdino'},
            {'nome': 'Davirlei', 'contato': 'Davirley (Tabela2)'},
            {'nome': 'Janilson', 'contato': 'Janilson (Tabela2)'},
            {'nome': 'Tânia', 'contato': 'Tânia (Serra)'},
            {'nome': 'Dilson', 'contato': 'Dilson Gonzaga (Tabela2)'},
            {'nome': 'Railda', 'contato': 'Railda'},
            {'nome': 'Fabiano', 'contato': 'Fabiano Serra(Tabela2)'},
            {'nome': 'Rosilene', 'contato': 'Rosilene (Tabela2)'},
            {'nome': 'Aline', 'contato': 'Aline Vereda'}]
nomes = []
for c in range(0,len(contatos)):
    nomes.append(contatos[c]['nome'])
while True:
    def telaBot():
        layout = [
            [sg.T('Barbalho: R$'),sg.I(size=(6,1),default_text=precoBB,key="-PRECOBB-"),sg.B('EDIT $',size=(6,1)),sg.B('ENVIAR',size=(6,1))],
            [sg.T('Goval: R$'),sg.I(size=(6,1),default_text=precoGV,key="-PRECOGV-")],
            [sg.I(key="-SEARCH-",size=(40,1))],
            [sg.T('Selecione para remover da lista de envio',text_color='yellow')],
            [sg.LB(nomes,size=(40,10),key="-CONTATOS-",select_mode='multiple',enable_events=False)],
            [sg.B('Confirmar')],
            [sg.T('Selecionados',text_color='yellow')],
            [sg.LB('',size=(40,10),key="-CONTATOSSELECT-")],
        ]
        return sg.Window('telaBot',layout=layout,finalize=True)


    nova_lista = []
    while True:
        event, values = telaBot().read()

        if event == 'Confirmar':
            nova_lista.append(values["-CONTATOS-"])
            print(nova_lista)
            telaBot().Element("-CONTATOSSELECT-").Update(nova_lista)
            break