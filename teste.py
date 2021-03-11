import PySimpleGUI as sg

categoria = ['Celular', 'Bateria', 'Carregador']
marca = ['Iphone', 'Motorola', 'LG']
cor = ['Branco', 'Verde', 'Preto']
fonte = 20

layout = [[sg.Text('Código', font=fonte), sg.Input(key='-COD-', font=fonte, size=(20, 1))],
          [sg.Text('Unidade', font=fonte), sg.InputText(key='-UNID-', font=fonte, size=(10, 1))],
          [sg.Text('Nome', font=fonte), sg.Input(key='-NOME-', size=(30, 1))],
          [sg.Text('Categoria', font=fonte), sg.Combo(categoria, font=fonte,default_value=categoria[0], key='-CATEG-', size=(30, 1))],
          [sg.Text('Marca', font=fonte), sg.Combo(marca, font=fonte, key='-MARCA-')],
          [sg.Text('Cor/Estampa', font=fonte), sg.Combo(cor, font=fonte, key='-COR-')],
          [sg.Text('')],
          [sg.Button('Cadastrar', font=fonte), sg.Button('Cancelar', font=fonte)]]

window = sg.Window('CADASTRO DE PRODUTOS', layout, size=(700, 300))

while True:
    event, values = window.read()
    if event in (sg.WIN_CLOSED, 'Cancelar'):
        break
    if event == 'Cadastrar':
        window['-NOME-'].update(window['-CATEG-'])

window.close()
