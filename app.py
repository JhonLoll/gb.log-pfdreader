import tabula
from functions.functions import *
from PySimpleGUI import PySimpleGUI as sg

def number_pages(layout):
    file_path = values['_FILEBROWSE_']
    layout_ = layout
    
    tables = tabula.read_pdf(file_path, pages='all')
    number_tables = len(tables)
    
    att_optionmenu(layout_, number_tables)

def att_optionmenu(layout, number_pages):
    layout[1][1].update(values = list(range(1, number_pages + 1)))

def retorno_conclusao(retorno):
    sg.popup("Processo finalizado", f'{retorno}')
    
sg.theme('BrownBlue')

layout = [
    # first-line
    [
        sg.Text('Selecione o arquivo PDF: '), 
        sg.Input(
            size=(10, 1),
            enable_events=True,
            key='_FILEBROWSE_'
            ),
        sg.FileBrowse(
            button_text='Buscar',
            auto_size_button=True,
            key='file_path'
            ),
    ],
    # second-line
    [
        sg.Text('Selecione a página: '),
        sg.OptionMenu(
            ['1'],
            default_value='',
            enable_events=True,
            key='_PAGE_'
            )
    ],
    # third-line
    [
        sg.Button('Converter')
    ]
]

window = sg.Window('Conversor PDF para XLSX', layout)

while True:
    events, values = window.read()
    
    if events == sg.WINDOW_CLOSED:
        break
    
    elif events == '_FILEBROWSE_':
        number_pages(layout)
        
    elif events == 'Converter':
        retorno = init_convertion(values['_FILEBROWSE_'], values['_PAGE_'])
        retorno_conclusao(retorno)
        
window.close()