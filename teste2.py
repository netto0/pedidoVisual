import PySimpleGUI as sg

categoria = ['Celular', 'Bateria', 'Carregador']

layout = [
          [sg.Frame('pet',layout=categoria)]
    ]

window = sg.Window('CADASTRO DE PRODUTOS', layout, size=(700, 300))

while True:
    event, values = window.read()
    if event in (sg.WIN_CLOSED, 'Cancelar'):
        break
    if event == 'Cadastrar':
        window['-NOME-'].update(window['-CATEG-'])

window.close()