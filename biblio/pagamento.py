def janelaPagamento():
    from PySimpleGUI import PySimpleGUI as sg

    formaDePagamento = ('Ch.', 'Bol.')
    prazoPagamento = ('À Vista', '14', '28', '30', '35', '40', '45', '50', '28/35', '30/40', '30/45', '30/60', '35/45', '30/40/50')

    sg.theme('Reddit')
    layout = [
        [sg.Radio('Cheque','formaPag',key='-PAGCHEQUE-'), sg.Radio('Boleto','formaPag',key='-PAGBOL-',default=True)],
        [sg.Listbox(prazoPagamento,size=(30,len(prazoPagamento)),key='-PRAZO-')],
        [sg.Button('OK'),sg.Button('CANCELAR')],
    ]
    window = sg.Window('telaPedido', layout=layout, finalize=True)

    while True:
        event,values = window.read()
        if event == sg.WIN_CLOSED:
            break
        if values['-PAGCHEQUE-'] == True:
            formaPag = 'Ch.'
        if values['-PAGBOL-'] == True:
            formaPag = 'Bol.'
        if event == 'OK':
            prazoPagamento = values['-PRAZO-'][0]
            sg.popup(f'O modo de pagamento escolhido foi: {formaPag} {prazoPagamento} dias')
            return f'Cond. de Pag.: {formaPag} {prazoPagamento} dias'
            break
        if event == 'CANCELAR':
            break

    window.close()
#janelaPagamento()
