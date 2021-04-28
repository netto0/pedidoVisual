import PySimpleGUI as sg
import re
import warnings
import pandas as pd


class ListboxWithSearch:

    def __init__(self, values, key='', select_mode='single',
                 size=(None, None), sort_fun=False, bind_return_key=False,
                 is_single_mode=True):
        if not is_single_mode:
            select_mode = 'extended'
            warnings.warn('Ev: is_single_mode is going to be deprecated use '
                          'select_mode instead', DeprecationWarning)
        self._key = key
        self._sort = sort_fun if sort_fun else lambda x: list(x)
        self._input_key = key + '_input'
        self._select_all_key = key + '_select_all'
        self._deselect_all_key = key + '_deselect_all'
        self._clear_search_key = key + '_clear_search'
        self._values = values
        self._selected = set()
        self._displayed_secret = values
        self._el = sg.Listbox(values=self._sort(self._displayed),
                              size=self._initialise_size(size),
                              key=key,
                              select_mode=select_mode,
                              default_values=[],
                              bind_return_key=bind_return_key)
        self._i = sg.I(key=self._input_key, enable_events=True,
                       tooltip='''
Start a string with = and everything after will be considered a regexp
otherwise it will be a simple match:
eg enter:
    =.*world$
vs entering
    world''')
        buttons = []
        if select_mode != 'single':
            buttons.append(sg.B('Select all', key=self._select_all_key))
        buttons.append(sg.B('Deselect all', key=self._deselect_all_key))
        buttons.append(sg.B('Clear search', key=self._clear_search_key))
        self.layout = sg.Column([
            [self._i],
            buttons,
            [self._el]])

    def frame_layout(self, name):
        return sg.Frame(name, layout=[[self.layout]])

    def _initialise_size(self, size):
        size = list(size)
        if size[0] is None and len(self._values) > 0:
            size[0] = max(len(x) for x in self._values) + 1
        if size[1] is None and len(self._values) > 0:
            size[1] = len(self._values) + 1
        return size

    @property
    def _displayed(self):
        return self._displayed_secret

    @property
    def selected(self):
        return tuple(self._selected)

    @_displayed.setter
    def _displayed(self, values):
        self._displayed_secret = (values if isinstance(values, dict)
                                  else set(values))
        self._el.Update(values=self._sort(self._displayed_secret),
                        set_to_index=0)

    def update(self, values):

        original_displayed = tuple(self._displayed)
        is_regexp = not(len(values[self._input_key]) > 0 and
                        values[self._input_key][0] == '=')

        if not is_regexp:
            search_string = values[self._input_key][1:]
        else:
            search_string = '.*' + re.escape(values[self._input_key]) + '.*'
        selected = values[self._key]

        def match_fun(s):
            try:
                return re.match(search_string, s, re.I)
            except re.error:
                return True

        # update displayed
        if isinstance(self._values, dict):
            self._displayed = {s: y for s, y in self._values.items()
                               if match_fun(s)}
        else:
            self._displayed = [s for s in self._values if match_fun(s)]

        self._update_selection(selected, original_displayed)

    def _update_selection(self, selected, original_displayed):
        # update selection
        if self._el.SelectMode in ['multiple', 'extended']:
            self._selected = self._selected - set(original_displayed)
            self._selected.update(selected)
        elif self._el.SelectMode == 'single':
            if len(selected) > 0:
                self._selected = set(selected)
        else:
            raise ValueError(self._el.SelectMode,
                             'expected "single" or "multiple"')
        selected_and_displayed = self._selected.intersection(self._displayed)

        # self._el.Update(values=self._sort(self._displayed), set_to_index=0)
        self._el.SetValue(selected_and_displayed)

    def _select_all_displayed(self):

        self._selected.update(self._displayed)
        self._el.SetValue(self._sort(self._displayed))

    def _deselect_all_displayed(self):

        for el in self._displayed:
            self._selected.discard(el)

        self._el.SetValue([])

    def set_values(self, values, selected=None):
        self._values = values
        self._displayed = values
        if selected is None:
            self._selected = set()
            self._el.SetValue([])
        else:
            if isinstance(selected, str):
                selected = [selected]
            self._selected = set(selected)
            self._el.SetValue(self._sort(selected))

    def _clear_search(self, values):
        selected = values[self._key]
        original_displayed = tuple(self._displayed)
        self._update_selection(selected, original_displayed)
        self._i.Update(value='')
        self.update({self._input_key: '',
                     self._key: tuple(self._selected)})

    def manage_events(self, event, values):
        if event == self._select_all_key:
            self._select_all_displayed()
        elif event == self._deselect_all_key:
            self._deselect_all_displayed()
        elif event == self._input_key:
            self.update(values)
        elif event == self._clear_search_key:
            self._clear_search(values)
        elif event is None:
            pass
        else:
            selected = values[self._key]
            original_displayed = tuple(self._displayed)
            self._update_selection(selected, original_displayed)


def get_date(title=None):
    if title is None:
        title = 'Choose Date'

    layout = [
        [sg.Text('Enter Date (YYYY-MM-DD) format')],
        [sg.CalendarButton('Pick Date', target='date', key='cal_button'),
         sg.Input(key='date', enable_events=True)],
        [sg.Button('Ok'), sg.Button('Cancel')]
    ]
    win = sg.Window(title, layout=layout)

    while True:
        event, values = win.Read()
        if event is None or event == 'Cancel':
            win.Close()
            return
        if event == 'date':
            date = values[event][:10]  # keep only YYYY-MM-DD
            win.Element(event).Update(value=date)
            try:
                date = pd.Timestamp(date)
                # Currently not supported
                # win.Element('cal_button').Update(
                #     default_date_m_d_y=(date.month, date.day, date.year))
            except ValueError:
                pass

        elif event == 'Ok':
            break
    win.Close()
    return pd.Timestamp(values['date']).to_pydatetime()


def show_hidden_files_button(win):
    """
    To be used with layouts that include sg.FileBrowser, the purpose is to
    allow hidden files to not be shown. Inspired from:
        https://stackoverflow.com/a/54068050/1764089
        and
        https://github.com/PySimpleGUI/PySimpleGUI/issues/1830
    example usecase:
        import PySimpleGui as sg
        win = sg.Window('Test', layout=[[sg.FileBrowser('Load file'),
                                         sg.Button('ok')]])
        show_hidden_files_button(win)
        while True:
            event, values = win.Read()
            if event is None:
                win.Close()
                break
    """
    # set up TKroot
    win.Read(timeout=0)

    # from https://stackoverflow.com/a/54068050/1764089
    try:
        win.TKroot.tk.call('tk_getOpenFile', '-foobarbaz')
    except sg.tk.TclError:
        pass
    win.TKroot.tk.call('set', '::tk::dialog::file::showHiddenBtn', '1')
    win.TKroot.tk.call('set', '::tk::dialog::file::showHiddenVar', '0')


if __name__ == '__main__':

    values = ["0 | Adriano Braga dos Santos","1 | Antônio Carlos Amaral","2 | Gilvan Delfino de Oliveira","3 | Ilma Delfino de Oliveira","4 | João Ferreira dos Santos","5 | José dos Reis Silva","6 | Lindraci Mendes Damascena","7 | Renato Moura Trindade","8 | Rodrigues & Pinheiro Ltda","9 | Santos & Souza Refeições Ltda","10| Pérola do Mucuri Sup. Dist. Alim.","11| Terezinha Joaquina Silva & Cia","12| Hélia Lacerda Machado","13| M. dos Santos Pereira & Cia Ltda","14| Maria Cleuza Costa & Cia Ltda","15| Adão Afonso da Silva","16| Prado & Prado Com. de Alim.","17| Débora Gil Alves Silva","18| Eunilto Maia Santos","19| Lidiomar Chaves Resende","20| Lucilene Dias Barbosa","21| Reinaldo Ferreira da Silva","22| Roberlan Medeiros","23| Manoelton Santos de Araújo","24| Supermercado Brás Ltda","25| Valdívio Leles dos Santos","26| Jaílson de Jesus","27| Jaílton Ferreira dos Santos","28| José Clemente de Jesus","29| Roberto Gomes da Silva","30| Alas da Silva Santos","31| Claudio Souza Cortes","32| Edenaldo Santana Souza","33| Ismar Costa Mendes Lima","34| Jorvane Antônio Lima","35| José Afonso Faria","36| Josenélia Farias Lucas","37| Maria Emília Silva de Souza","38| Márcio Simão da Silva","39| Mercearia Mineira Ltda","40| Núbia C. B. Leite","41| Oliveira & Leite Ltda","42| Sílvio Cláudio Com. Prod. Alim.","43| Valdomiro Oliveira","44| Adão Brandão dos Santos","45| Adilson Ramos Pereira","46| Afrodízio Tenencio de Brito","47| Da Hora e Soares Ltda","48| Fábio de Souza Bom Jardim","49| João Souza dos Santos","50| Raquel Pereira Mota","51| Laís Lima Brito","52| Maria da Glória de Jesus Viana","53| Ademir Galdino","54| Davirlei de Jesus Costa","55| Detrez Azevedo Com. Prod. Alim.","56| Edgard Rocha Santos","57| Eric Melo de Oliveira","58| Hélcio Ramos Sobral","59| Janilson Soares Araújo","60| José Roberto Dias do Amaral","61| Valdirene Bremer Ramalho","62| Milton Alves de Almeida","63| Zilberto Freitas Meireles","64| Fabiano Folgado","65| Rosilene L. Espíndola","66| Aline Almeida Lacerda","67| Valdemir Pereira Santos",
]
    listbox1 = ListboxWithSearch(values, 'mylistbox', size=(55,20))
    listbox2 = ListboxWithSearch(values, 'mylistbox0', select_mode='extended',size=(55,20))
    layout = [[listbox1.layout],
              [listbox2.layout]]
    win = sg.Window('test', layout=layout, resizable=True)
    while True:
        event, values = win.Read()
        print(event, values)
        if event is None:
            break
        else:
            listbox1.manage_events(event, values)
            listbox2.manage_events(event, values)