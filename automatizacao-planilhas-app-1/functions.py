import PySimpleGUI as sg
import pandas as pd
import openpyxl
import tkinter as tk
import shutil


from tkinter import ttk
from tkinter.simpledialog import askstring
from tkinter import filedialog


def create_table_popup(data, main_window, combo, data_ref):
     # Tamanho padrão da janela
    default_width = 800
    default_height = 800

    # Criar uma nova janela pop-up para a tabela
    table_window = tk.Toplevel()
    table_window.title('Editar Planilha')

    # Definir o tamanho padrão da janela
    table_window.geometry(f"{default_width}x{default_height}")

    # Criar uma Treeview para exibir a tabela
    tree = ttk.Treeview(table_window)

    # Criar uma barra de rolagem vertical
    vsb = ttk.Scrollbar(table_window, orient="vertical", command=tree.yview)
    vsb.pack(side='right', fill='y')

    # Criar uma barra de rolagem horizontal
    hsb = ttk.Scrollbar(table_window, orient="horizontal", command=tree.xview)
    hsb.pack(side='bottom', fill='x')

    tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

    tree['columns'] = list(data.columns)

    # Definir cabeçalhos da Treeview
    for col in data.columns:
        tree.heading(col, text=col)
        tree.column(col, width=100)  # Ajuste a largura conforme necessário

    # Inserir os dados na Treeview
    for i, row in data.iterrows():
        tree.insert('', 'end', iid=i, values=tuple(row))

    tree.pack(expand=True, fill='both')

    # Função para editar células
    def edit_cell(event):
        item = tree.selection()[0]
        col = tree.identify_column(event.x)
        col = col.split('#')[-1]
        old_value = tree.item(item, 'values')[int(col)-1]
        new_value = askstring('Editar Célula', f'Editar valor para {data.columns[int(col)-1]}:', initialvalue=old_value)
        if new_value is not None:
            tree.item(item, values=tree.item(item, 'values')[:int(col)-1] + (new_value,) + tree.item(item, 'values')[int(col):])


    def save_changes():

        try:
            if combo == 'Folha de Pagamento':
                name_file = f"FolhaPagto-{data_ref}_0011.xlsx"
                file_type_select = "Excel files", "*.xlsx"
            elif combo == 'Centro de Trabalho':
                name_file = f'centro_de_trabalho.csv'
                file_type_select = "CSV files", "*.csv"
            elif combo == 'Códigos de Débito e Crédito':
                name_file = 'codigos.csv'
                file_type_select = "CSV files", "*.csv"

            # Perguntar ao usuário onde deseja salvar a tabela editada
            file_path = filedialog.asksaveasfilename(filetypes=[file_type_select], title="Salvar Tabela Editada", initialfile=name_file)

            if file_path:
                # Determine the file extension
                if file_path.lower().endswith('.csv'):
                    # Obter os dados da Treeview
                    edited_data = []
                    for item in tree.get_children():
                        values = tree.item(item, 'values')
                        edited_data.append(values)

                    # Criar um novo DataFrame com os dados editados
                    edited_df = pd.DataFrame(edited_data, columns=data.columns)

                    # Salvar o DataFrame como um arquivo CSV
                    edited_df.to_csv(file_path, index=False, encoding='ISO-8859-1')

                elif file_path.lower().endswith('.xlsx'):
                    # Obter os dados da Treeview
                    edited_data = []
                    for item in tree.get_children():
                        values = tree.item(item, 'values')
                        edited_data.append(values)

                    # Criar um novo DataFrame com os dados editados
                    edited_df = pd.DataFrame(edited_data, columns=data.columns)

                    # Salvar o DataFrame como um arquivo XLSX
                    edited_df.to_excel(file_path, index=False)

                sg.popup(f'Tabela editada salva em:\n{file_path}')

                main_window.finalize()

        except Exception as e:
            sg.popup_error(f'Ocorreu um erro ao salvar a tabela editada: {str(e)}')

    def cancel_edit():
        table_window.destroy()
        main_window['-EDIT-'].update(disabled=False) # Para garantir que o botão '-EDIT-' seja ativado
    
    def add_new_data():
        # Adicione uma nova linha vazia à tabela
        tree.insert('', 'end', values=[''] * len(data.columns))
        main_window.finalize()

    # Adicionar botões
    button_frame = tk.Frame(table_window)
    button_frame.pack()

    save_button = tk.Button(button_frame, text="Salvar Alterações", command=save_changes)
    save_button.pack(side=tk.LEFT, padx=5, pady=5)

    add_button = tk.Button(button_frame, text="Adicionar Novos Dados", command=add_new_data)
    add_button.pack(side=tk.LEFT, padx=5, pady=5)

    cancel_button = tk.Button(button_frame, text="Cancelar Edição", command=cancel_edit)
    cancel_button.pack(side=tk.LEFT, padx=5, pady=5)

    tree.bind('<Double-1>', edit_cell)  # Duplo clique para editar células

    table_window.protocol('WM_DELETE_WINDOW', cancel_edit) # Isso vincula a função cancel_edit para ser chamada quando a janela é fechada usando o botão de fechar (X) sg.WINDOW_CLOSED


def save_sheet(main_window, file_path, select_combo, data_ref):

    try:
        if select_combo == 'Folha de Pagamento':
            name_file = f"FolhaPagto-{data_ref}_0011.xlsx"
            file_type_select = "Excel files", "*.xlsx"
        elif select_combo == 'Centro de Trabalho':
            name_file = f'centro_de_trabalho.csv'
            file_type_select = "CSV files", "*.csv"
        elif select_combo == 'Códigos de Débito e Crédito':
            name_file = 'codigos.csv'
            file_type_select = "CSV files", "*.csv"
        
        # Perguntar ao usuário onde deseja salvar a tabela editada
        file_new_path = filedialog.asksaveasfilename(filetypes=[file_type_select], title="Salvar Planilha Selecionada", initialfile=name_file)

        if file_new_path:
            # Determine the file extension
            if file_new_path.lower().endswith('.csv'):
                # Copy the content of the file to the new location
                shutil.copyfile(file_path, file_new_path)

            elif file_new_path.lower().endswith('.xlsx'):
                # Copy the content of the file to the new location
                shutil.copyfile(file_path, file_new_path)

        
            sg.popup(f'Tabela editada salva em:\n{file_new_path}')

            main_window.finalize()

    except Exception as e:
        return sg.popup_error(f'Ocorreu um erro, verifique se os caminhos são diferentes: {str(e)}')


def analyze_centro_trabalho_e_codigos(data, target_column):
    num_linhas = len(data)
    num_colunas = len(data.columns)
    contagem_target_column = data[target_column].value_counts()

    return num_linhas, num_colunas, contagem_target_column


def display_dataframe_in_multiline_inicial(data):
    # Adicionar cada linha do DataFrame formatada com espaços adicionais
    df_str = ""
    for index, row in data.iterrows():
        coluna = row["Coluna"]
        soma = f"{row['Soma']:.2f}"
        spaces = "-" * (80 - len(coluna) - len(soma))
        df_str += f"{coluna} {spaces} {soma}\n"
    
    return df_str


def display_dataframe_in_multiline_final(data):
    # Adicionar cada linha do DataFrame formatada com espaços adicionais
    df_str = ""
    for index, row in data.iterrows():
        coluna = row["CR"]  # Alterado para "CR" para corresponder ao seu exemplo
        situacao = row["SITUAÇÃO"]  # Alterado para "SITUAÇÃO"
        unidade_negocio = row["UNIDADE DE NEGOCIO"]  # Alterado para "UNIDADE DE NEGOCIO"
        valor = f"{row['VALOR']:.2f}"  # Formatação do valor
        # Crie a string formatada para a linha atual
        formatted_row = f"{coluna:<25} {situacao:<25} {unidade_negocio:<25} {valor}\n"
        df_str += formatted_row
    
    return df_str


def analyze_folha_pamento(data_folha):

    colunas_para_somar = data_folha.iloc[:, 6:]
    soma_colunas = colunas_para_somar.sum().round(2)
    
    # Crie um novo DataFrame para mostrar a soma das colunas.
    df_soma_fp = pd.DataFrame({'Coluna': soma_colunas.index, 'Soma': soma_colunas.values})

    num_linhas = len(data_folha)
    num_colunas = len(data_folha.columns)
    df_str = display_dataframe_in_multiline_inicial(df_soma_fp)
    soma_total = soma_colunas.sum().round(2)

    return num_linhas, num_colunas, soma_total, df_str


def generate(data_ref):
    try:
        data_codigos = pd.read_csv('utils\codigos.csv', encoding='ISO-8859-1')

        caminho = f"input\\FolhaPagto-{data_ref}_0011.xlsx"
        data_folha = pd.read_excel(caminho, header=3)

        data_formatada = f"{data_ref[:2]}/{data_ref[2:4]}/{data_ref[4:]}"

        num_linhas, num_colunas, soma_total, df_str, mensagem, resumo, lista_resumo, resultado, resultado_situacao = treatment_folha_de_pagamento(data_folha)

        df_ref = resumo.set_index('CR')
        saida = pd.DataFrame(columns=['DATA', 'TIPO', 'DEBITO', 'CREDITO', 'DESCRICAO COMPLETA', 'VALOR', 'UNIDADE DE NEGOCIO', 'CR'])

        for cr_index in df_ref.index:
            if cr_index == 'CR1002':
                codigos = data_codigos.loc[data_codigos['CENTRO DE CUSTO'] == 'adm']
            elif cr_index == 'CR2020':
                codigos = data_codigos.loc[data_codigos['CENTRO DE CUSTO'] == 'com']
            else:
                codigos = data_codigos.loc[data_codigos['CENTRO DE CUSTO'] == 'custo']

            for _, linha in codigos.iterrows():
                codigo1_correspondente = linha['DEBITO']
                codigo2_correspondente = linha['CREDITO']
                descricao_correspondente = linha['DESCRICAO']
                folha_correspondente = linha['COLUNA DA FOLHA']

                if folha_correspondente in lista_resumo['LISTA'].values:
                    valor_correspondente = df_ref.loc[cr_index, folha_correspondente]
                    unidade_correspondente = df_ref.loc[cr_index, 'UNIDADE DE NEGOCIO']
                    ref_completa = f'{descricao_correspondente} {folha_correspondente}'

                    if valor_correspondente > 0:
                        saida.loc[len(saida)] = [data_formatada, 'MANUAL', codigo1_correspondente, codigo2_correspondente, ref_completa, valor_correspondente, unidade_correspondente, cr_index]
                    else:
                        saida = saida
                else:
                    saida

            resultado_xlsx = f'output/export_folha_de_pagamento_{data_ref}.xlsx'
            resultado_csv = f'output/export_folha_de_pagamento_{data_ref}.csv'

        return saida.to_excel(resultado_xlsx, index=False), saida.to_csv(resultado_csv, index=False)

    except Exception as e:
        return sg.popup_error(f'Ocorreu um erro, veja se o arquivos foram importados: {str(e)}')


def correlation_between_spreadsheets_folha_pagamento_e_centro_trabalho(data_folha):
    data_ct = pd.read_csv('utils\centro_de_trabalho.csv', encoding='ISO-8859-1')
    data_fp_ct = data_folha.merge(data_ct[['DIV. RH', 'CENTRO DE CUSTO', 'CR', 'UNIDADE DE NEGOCIO']], on='DIV. RH', how='left')

    data_fp_ct['SITUAÇÃO'] = data_fp_ct['CENTRO DE CUSTO'].str.cat(data_fp_ct['CR'])
        
    total_registros = data_fp_ct['DIV. RH'].count()
    total_situacao = data_fp_ct['SITUAÇÃO'].count()
    quantidade_com_nulos_situacao = data_fp_ct['SITUAÇÃO'].isnull()
    #valores_nulos_situacao = data_fp_ct[data_fp_ct['SITUAÇÃO'].isnull()]
    #linhas_com_nulos_situacao = valores_nulos_situacao[['CÓDIGO', 'DIV. RH', 'CENTRO DE CUSTO', 'CR', 'SITUAÇÃO']]
    
    colunas_float = data_fp_ct.select_dtypes(include='float64').columns.tolist()
    colunas_condicionais = ['CR', 'SITUAÇÃO', 'UNIDADE DE NEGOCIO']
    colunas_relevantes = data_fp_ct[colunas_float + colunas_condicionais]
    soma_condicional = colunas_relevantes.groupby(colunas_condicionais).sum()
    soma_condicional['VALOR'] = soma_condicional.sum(axis=1)

    resumo = soma_condicional.reset_index()
    select_columns = soma_condicional.columns[1:]
    lista_resumo = pd.DataFrame({'LISTA': select_columns})
    
    resultado_situacao = soma_condicional[['VALOR']].round(2)
    resultado_situacao.reset_index(inplace=True)
    resultado_situacao.columns = ['CR', 'SITUAÇÃO', 'UNIDADE DE NEGOCIO', 'VALOR']
    resultado = resultado_situacao['VALOR'].sum()

    if total_situacao == total_registros:
        mensagem = (f'Sucesso!. Todos os {total_registros} registros foram classificados em "SITUAÇÃO')
    else:
        mensagem = (f'Incosistência!. Dos {total_registros} registros, {total_situacao} foram classificados em "SITUAÇÃO".\n'
                     f'Existem {quantidade_com_nulos_situacao.sum()} registros que não foram corretamente classificados em "SITUAÇÃO".')

    #print(f'Veja a relação a seguir: \n \n {linhas_com_nulos_situacao}')

    return mensagem, resumo, lista_resumo, resultado, resultado_situacao


def treatment_folha_de_pagamento(data_folha):
    colunas_para_alterar = data_folha.columns[6:]
    
    def tratar_valor(valor):
        valor = str(valor).strip()
        if valor == '-':
            return 0
        return valor

    # Loop pelas colunas e aplique a função de tratamento em cada uma delas.
    for coluna in colunas_para_alterar:
        data_folha[coluna] = data_folha[coluna].apply(tratar_valor)

    for coluna in colunas_para_alterar:
        data_folha[coluna] = data_folha[coluna].astype('float')
    
    # Padronizando os valores da coluna 'DIV. RH'
    data_folha['DIV. RH'] = data_folha['DIV. RH'].str.replace('.', '')
    data_folha['DIV. RH'] = data_folha['DIV. RH'].astype('int64')
    data_folha['DIV. RH'].head()

    num_linhas, num_colunas, soma_total, df_str = analyze_folha_pamento(data_folha)
    mensagem, resumo, lista_resumo, resultado, resultado_situacao = correlation_between_spreadsheets_folha_pagamento_e_centro_trabalho(data_folha)

    return num_linhas, num_colunas, soma_total, df_str, mensagem, resumo, lista_resumo, resultado, resultado_situacao


def load_default_sheet(main_window, sheet, data_ref):
    try:
        if (sheet == 'Centro de Trabalho'):
            data = pd.read_csv('utils\centro_de_trabalho.csv', encoding='ISO-8859-1')
            target_column_CT = 'CR'
            
            num_linhas, num_colunas, contagem_target_column = analyze_centro_trabalho_e_codigos(data, target_column_CT)
            main_window['-NUM_ROWS_CT-'].update(num_linhas)
            main_window['-NUM_COLUMNS_CT-'].update(num_colunas)
            main_window['-COLUMN_COUNT_CT-'].update(contagem_target_column)
            
        elif (sheet == 'Códigos de Débito e Crédito'):
            data = pd.read_csv('utils\codigos.csv', encoding='ISO-8859-1')
            target_column_DC = 'CENTRO DE CUSTO'

            num_linhas, num_colunas, contagem_target_column = analyze_centro_trabalho_e_codigos(data, target_column_DC)
            main_window['-NUM_ROWS_CD-'].update(num_linhas)
            main_window['-NUM_COLUMNS_CD-'].update(num_colunas)
            main_window['-COLUMN_COUNT_CD-'].update(contagem_target_column)

        elif (sheet == 'Folha de Pagamento'):
            caminho = f'input\FolhaPagto-{data_ref}_0011.xlsx'
            data_folha = pd.read_excel(caminho, header=3)

            num_linhas, num_colunas, soma_total, df_str, mensagem, resumo, lista_resumo, resultado, resultado_situacao = treatment_folha_de_pagamento(data_folha)

            main_window['-NUM_ROWS_FP-'].update(num_linhas)
            main_window['-NUM_COLUMNS_FP-'].update(num_colunas)
            main_window['-COLUMN_SUM_FP-'].update(soma_total)
            main_window['-MULTILINE_FP_INICIAL-'].update(value=df_str)
            main_window['-RESULT_SITUACAO-'].update(resultado)
            main_window['-RESULT_CORRELATION-'].update(mensagem)
            main_window['-MULTILINE_FP_SITUACAO-'].update(value=display_dataframe_in_multiline_final(resultado_situacao))
    
    except Exception as e:
        sg.popup_error(f'Ocorreu um erro, veja se o arquivo foi importado: {str(e)}')


def edit_file(main_window, combo, data_ref):
    try:
        if combo == 'Centro de Trabalho':
            data = pd.read_csv('utils\centro_de_trabalho.csv', encoding='ISO-8859-1')
        elif combo == 'Códigos de Débito e Crédito':
            data = pd.read_csv('utils\codigos.csv', encoding='ISO-8859-1')
        elif combo == 'Folha de Pagamento':
            data = pd.read_excel(f'input\FolhaPagto-{data_ref}_0011.xlsx')

        main_window['-EDIT-'].update(disabled=True) # Para garantir que o botão '-EDIT-' seja desativado
        create_table_popup(data, main_window, combo, data_ref) # Chamar a função para exibir a tabela
        
    except Exception as e:
        sg.popup_error(f'Ocorreu um erro ao editar o arquivo CSV: {str(e)}')
