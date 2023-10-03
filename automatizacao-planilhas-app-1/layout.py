import PySimpleGUI as sg
import pandas as pd
import openpyxl
import tkinter as tk
from tkinter import ttk
from tkinter.simpledialog import askstring
from tkinter import filedialog

from functions import edit_file, load_default_sheet, generate, save_sheet


# Definir a função para criar a janela inicial
def create_main_window():
    layout_centro_de_trabalho = [
        [sg.Text('Informações:', font=('Arial', 12))],
        [sg.Text('Número de linhas: ', size=(15, 2)), sg.Text('', size=(10, 2), key='-NUM_ROWS_CT-')],
        [sg.Text('Número de colunas: ', size=(15, 2)), sg.Text('', size=(10, 2), key='-NUM_COLUMNS_CT-')],
        [sg.Text('Contagem registro CR:', size=(15, 2))],
        [sg.Multiline('', size=(24, 24), autoscroll=True, key='-COLUMN_COUNT_CT-')],
    ]

    layout_codigos_dc = [
        [sg.Text('Informações:', font=('Arial', 12))],
        [sg.Text('Número de linhas: ', size=(15, 2)), sg.Text('', size=(10, 2), key='-NUM_ROWS_CD-')],
        [sg.Text('Número de colunas: ', size=(15, 2)), sg.Text('', size=(10, 2), key='-NUM_COLUMNS_CD-')],
        [sg.Text('Contagem de registros débito e crédito por centro de custo:', size=(50, 2))],
        [sg.Multiline('', size=(24, 8), autoscroll=True, key='-COLUMN_COUNT_CD-')],
    ]

    layout_folha_de_pagamento = [
        [sg.Text('Informações:', font=('Arial', 12))],
        [sg.Text('Número de linhas: ', size=(15, 2)), sg.Text('', size=(10, 2), key='-NUM_ROWS_FP-')],
        [sg.Text('Número de colunas: ', size=(15, 2)), sg.Text('', size=(10, 2), key='-NUM_COLUMNS_FP-')],
        [sg.Text('Valor total das colunas importadas:', size=(35, 2)), sg.Text('', size=(10, 2), key='-COLUMN_SUM_FP-')],
        [sg.Multiline('', size=(80, 10), autoscroll=True, key='-MULTILINE_FP_INICIAL-')],
        [sg.Text('Resultado da correlação entre as planilhas:', size=(50, 2))], 
        [sg.Text('', size=(50, 2), key='-RESULT_CORRELATION-')],
        [sg.Text('Resumo da Situação. Valor de cada centro de custo após correlação:', size=(75, 2)), sg.Text('', size=(10, 2), key='-RESULT_SITUACAO-')],
        [sg.Multiline('', size=(80, 10), autoscroll=True, key='-MULTILINE_FP_SITUACAO-'),
         sg.Button('Gerar planilha', font=('Helvetica', 16), key='-GENERATE-')],

    ]    

    layout = [
        [sg.Text('Selecione qual arquivo deseja importar ou visualizar:')],
        [sg.Combo(['Centro de Trabalho', 'Códigos de Débito e Crédito', 'Folha de Pagamento'], key='-COMBO-', default_value='Folha de Pagamento'),
         sg.Text('Informe a data de referência da folha de pagamento (ex. 31072023) :'), sg.InputText(key='-DATA_REF-', size=(15,2))],
        [sg.Text('Informe o caminho do arquivo:'), sg.InputText(key='-FILE_PATH-'), sg.FileBrowse()],
        [sg.Button('Salvar Planilha', key='-SALVE-'),
         sg.Button('Visualizar Planilha', key='-EDIT-'),
         sg.Button('Carregar Informações', key='-ATT-')],
    ]    

    layout += [
        [sg.TabGroup([[
            sg.Tab('Centro de Trabalho', layout_centro_de_trabalho),
            sg.Tab('Códigos de Débito e Crédito', layout_codigos_dc),
            sg.Tab('Folha de Pagamento', layout_folha_de_pagamento)
            ]], key='-TAB GROUP-', expand_x=True, expand_y=True)],
    ]

    # Criar a janela
    window = sg.Window('Padronização de Arquivos', layout, size=(800, 800))

    return window

# Criar a janela inicial
main_window = create_main_window()

# Loop de eventos
while True:
    event, values = main_window.read()

    if event == sg.WINDOW_CLOSED:
        break

    elif event == '-ATT-':
        data_ref = values['-DATA_REF-']
        
        if data_ref:
            load_default_sheet(main_window, 'Centro de Trabalho', data_ref)
            load_default_sheet(main_window, 'Códigos de Débito e Crédito', data_ref)
            load_default_sheet(main_window, 'Folha de Pagamento', data_ref)
        else:
            sg.popup('Por favor, informe a data carregamento da folha de pagamento.')

    elif event == '-EDIT-':
        data_ref = values['-DATA_REF-']
        select_combo = values['-COMBO-']

        if select_combo != 'Folha de Pagamento':
            edit_file(main_window, select_combo, '')
        elif select_combo == 'Folha de Pagamento' and data_ref:
            edit_file(main_window, select_combo, data_ref)
        else:
            sg.popup('Por favor, selecione a data de referência.')
    
    elif event == '-GENERATE-':
        data_ref = values['-DATA_REF-']

        if data_ref:
            generate(data_ref)
    
    elif event == '-SALVE-':
        file_path = values['-FILE_PATH-']
        select_combo = values['-COMBO-']
        data_ref = values['-DATA_REF-']

        if select_combo != 'Folha de Pagamento' and file_path:
            save_sheet(main_window, file_path, select_combo, '')
        elif select_combo == 'Folha de Pagamento' and data_ref and file_path:
            save_sheet(main_window, file_path, select_combo, data_ref)
        else:
            sg.popup('Atenção: \n 01) Para salvar um arquivo, você deve selecionar o arquivo desejado e escolher o caminho do mesmo. \n 02) Exclusivamente para Folha de Pagamento é obrigatório informar a data de referência')


# Fechar a janela ao sair
main_window.close()
