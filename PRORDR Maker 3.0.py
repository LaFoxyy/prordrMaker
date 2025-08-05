"""
Código por: Guilherme Amancio & Yan Mello
Ultima atualização: 3.0, 04/01/23
"""


import PySimpleGUI as sg
import pandas as pd
import csv
from fpdf import FPDF
from tkinter import *

#dataframe for cities
dfcities = pd.read_excel("F:\Downloads\Velvet Paws Studio\CidadesPRORDR.xlsx")
cities = dfcities['Nomes'].tolist()

#change theme
#sg.theme('DarkGreen2')

#global variable
clname = ''
tpname = ''
nsname = ''
ctname = ''
selRow = 0

contar = 0

#define the city codes
def get_city_code(city_name):
    city_code = dfcities[dfcities.iloc[:, 0] == city_name].iloc[0, 1]
    return str(city_code)

#define the table variables titled and data
def generate_table_data():
    headings = ['Coordenada(X)', 'Coordenada(Y)', 'Zona', 'Estação', 'Ponto Visado', 
                'Progressiva', 'Cota', 'Ângulo', 'Minuto', 'Sentido', 'Codigo de Acidente']
    data = []

    return headings, data

#Main menu
def main_menu():
    clientName = sg.Input(size=(30,1), key='-NAME-', enable_events=True)
    topName = sg.Input(size=(30,1), key='-TOP-', enable_events=True)
    nsValue = sg.Input(size=(30,1), key='-NS-', enable_events=True)
    cityName = sg.Combo(values=cities, size=(30,1), key='-CITY-', enable_events=True)
    
    leftlayout = [
        [sg.Text("NS:", size=(8,1), justification='right'), nsValue],
        [sg.Text("Cliente:", size=(8,1), justification='right'), clientName]
    ]
    rightlayout = [
        [sg.Text("Topógrafo:", size=(8,1), justification='right'), topName],
        [sg.Text("Cidade:", size=(8,1), justification='right'), cityName]
    ]

    layout = [
        [sg.Text("PRORDR Maker 3.0", font=("Arial Bold", 25), justification='center')],
        
        [sg.Column(leftlayout), sg.VSeparator(), sg.Column(rightlayout)],

        [sg.Button("Carregar"), sg.Button("Continuar")]
    ]

    window = sg.Window('PRORDR Maker', layout, resizable=True, finalize=True)

    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED:
            break
        if event == "Continuar":
            clname = values['-NAME-']
            tpname = values['-TOP-']
            nsname = values['-NS-']
            ctname = values['-CITY-']
            data = []
            #print(clname, tpname, nsname, ctname)
            if len(nsname) == 10 and nsname.isdigit()and clname != '' and tpname != '' and ctname in cities:
                window.close()
                generate_table(nsname, clname, tpname, ctname, data)
            else:
                sg.popup_error(f"NS {nsname} deve possuir 10 digitos e ser numérico! O cliente, cidade e o topógrafo não podem ser nulos!")
        elif event == "Carregar":
            try:
                filename = sg.popup_get_file('Selecione o arquivo a carregar!', file_types=(("CSV Files", "*.csv"),))
                if filename:
                    with open(filename, "r") as csvfile:
                        csvreader = csv.reader(csvfile)
                        dados = list(csvreader)
                        infoload = dados[0]
                        dataload = dados[3:]
                    
                    cl, ct, tp, nus = infoload

                    data = []
                    for row in dataload:
                        x, y, z, e, pv, p, c, a, m, s, ca = row
                        data.append([x, y, z, e, pv, p, c, a, m, s, ca])
                    window.close()
                    #print(nsname, ctname, tpname, nsname)
                    generate_table(nus, tp, cl, ct, data)    
            except Exception as er:
                sg.popup_error(f"Erro ao carregar arquivo {er}")

# Antiga função de edição individual de celulas. Provavelmente não é mais necessário o uso do TKinter sem ela, provavelmente
# TKinter function to display and edit value in cell
'''def edit_cell(window, key, row, col, data, justify='left'):

    global textvariable, edit

    def callback(event, row, col, text, key):
        global edit
        # event.widget gives you the same entry widget we created earlier
        widget = event.widget
        if key == 'Focus_Out':
            # Get new text that has been typed into widget
            text = widget.get()
            # Print to terminal
            #print(text)
            data[row - 1][col] = text
        # Destroy the entry widget
        widget.destroy()
        # Destroy all widgets
        widget.master.destroy()
        # Get the row from the table that was edited
        # table variable exists here because it was called before the callback
        values = list(table.item(row, 'values'))
        # Store new value in the appropriate row and column
        values[col] = text
        table.item(row, values=values)
        edit = False

    if edit or row <= 0:
        return

    edit = True
    # Get the Tkinter functionality for our window
    root = window.TKroot
    # Gets the Widget object from the PySimpleGUI table - a PySimpleGUI table is really
    # what's called a TreeView widget in TKinter
    table = window[key].Widget
    # Get the row as a dict using .item function and get individual value using [col]
    # Get currently selected value
    text = table.item(row, "values")[col]
    # Return x and y position of cell as well as width and height (in TreeView widget)
    x, y, width, height = table.bbox(row, col)

    # Create a new container that acts as container for the editable text input widget
    frame = sg.tk.Frame(root)
    # put frame in same location as selected cell
    frame.place(x=x, y=y, anchor="nw", width=width, height=height)

    # textvariable represents a text value
    textvariable = sg.tk.StringVar()
    textvariable.set(text)
    # Used to acceot single line text input from user - editable text input
    # frame is the parent window, textvariable is the initial value, justify is the position
    entry = sg.tk.Entry(frame, textvariable=textvariable, justify=justify)
    # Organizes widgets into blocks before putting them into the parent
    entry.pack()
    # selects all text in the entry input widget
    entry.select_range(0, sg.tk.END)
    # Puts cursor at end of input text
    entry.icursor(sg.tk.END)
    # Forces focus on the entry widget (actually when the user clicks because this initiates all this Tkinter stuff, e
    # ending with a focus on what has been created)
    entry.focus_force()
    # When you click outside of the selected widget, everything is returned back to normal
    # lambda e generates an empty function, which is turned into an event function 
    # which corresponds to the "FocusOut" (clicking outside of the cell) event
    entry.bind("<FocusOut>", lambda e, r=row, c=col, t=text, k='Focus_Out':callback(e, r, c, t, k))'''

#generate the table and the layout
def generate_table(ns, name, top, city, dados):
    global edit, contar

    edit = False

    if dados == []:
        headings, data = generate_table_data()
    else:
        headings, data = generate_table_data()
        data = dados
    sg.set_options(dpi_awareness=True)

    #individual layout
    lx = [
        [sg.Text('Coordenada(X)')],
        [sg.InputText(key='-X-', size=(12, 1), enable_events=True)]
    ]
    ly = [
        [sg.Text('Coordenada(Y)')], 
        [sg.InputText(key='-Y-', size=(12, 1), enable_events=True)]
    ]
    lz = [
        [sg.Text('Zona')],
        [sg.InputText("23", key='-Z-', size=(10, 1), enable_events=True)]
    ]
    le = [
        [sg.Text('Estação')],
        [sg.InputText(key='-E-', size=(10, 1), enable_events=True)]
    ]
    lpv = [
        [sg.Text('Ponto Visado')],
        [sg.InputText(key='-PV-', size=(10, 1), enable_events=True)]
    ]
    lp = [
        [sg.Text('Progressiva')],
        [sg.InputText(key='-P-', size=(10, 1), enable_events=True)]
    ]
    lc = [
        [sg.Text('Cota')],
        [sg.InputText(key='-C-', size=(10, 1), enable_events=True)]
    ]
    la = [
        [sg.Text('Ângulo')],
        [sg.InputText(key='-A-', size=(10, 1), enable_events=True)]
    ]
    ls = [
        [sg.Text('Sentido')],
        [sg.Combo(['Direito', 'Esquerdo'], key='-S-', size=(10, 1), enable_events=True)]
    ]
    lca = [
        [sg.Text('Codigo de Acidente')],
        [sg.Combo(['000','003', '004', '005', '007', '008','010'], key='-CA-', size=(10, 1), default_value='000', enable_events=True)]
    ]
    lnyon = [
        [sg.Button('Add', key='-ADD-'), sg.Text('Linha selecionada: '), sg.Text(' ', key='twitter')],
        [sg.Button('Insert', key='-INS-'), sg.Button('Delete', key='-DEL-'), sg.Button('Edit', key='-EDIT-')]
    ]

    #Main layout
    layout = [
            [sg.Column(lx), sg.VSeperator(), sg.Column(ly), sg.VSeperator(), sg.Column(lz), sg.VSeperator(),
             sg.Column(le), sg.VSeperator(), sg.Column(lpv), sg.VSeperator(), sg.Column(lp), sg.VSeperator(),
             sg.Column(lc), sg.VSeperator(), sg.Column(la), sg.VSeperator(), sg.Column(ls), sg.VSeperator(),
             sg.Column(lca), sg.Column(lnyon)],
             

            [sg.Table(values=data, headings=headings, max_col_width=25,
                    font=("Arial", 12), auto_size_columns=True,
                    justification='center', num_rows=20, 
                    alternating_row_color=sg.theme_button_color()[1],
                    key='-TABLE-', expand_x=True, expand_y=True,
                    enable_click_events=True, selected_row_colors='red on yellow')],

            [sg.Text('Celula Selecionada:'), sg.T(key='-CLICKED_CELL-')],
            [sg.Button('Salvar'), sg.Button('Exportar'), sg.Button('PDF')],

            [sg.Button('Voltar'), sg.Text("Total de US: "), sg.Text("0", key='xxx')]
    ]


    window = sg.Window('PRORDR Maker', layout, resizable=True, finalize=True)

    while True:
        #print(data)
        global selRow
        event, values = window.read()
        if event in (sg.WIN_CLOSED, 'Exit'):
            break
        # Checks if the event object is of tuple data type, indicating a click on a cell'
        elif event == '-ADD-':
            x = values['-X-']
            y = values['-Y-']
            z = int(values['-Z-'])
            e = values['-E-']
            pv = values['-PV-']
            p = values['-P-']
            c = values['-C-']
            a = values['-A-'] #00
            m = 00
            s = values['-S-'] # 
            ca = values['-CA-'] #010 se nao 000
            # Verificar as condições antes de adicionar os valores a data
            if a == '':
                a = 0
            if len(x) == 6 and x.isnumeric() and len(y) == 7 and y.isnumeric() and z == 23 \
                    and len(e) == 3 and len(pv) == 3 and p.isnumeric() and c.isnumeric() \
                    and 0 <= int(a) < 120 and s in ['Direito', 'Esquerdo', ''] and ca in ['000', '003', '004', '005', '007', '008', '010']:

                data.append([x, y, z, e, pv, p, c, a, m, s, ca])
                window['-TABLE-'].update(values=data)

                kj = ((int)(p) / 1000) * 5
                window['xxx'].update(kj)

                window['-X-'].update("")
                window['-Y-'].update("")
                window['-Z-'].update(23)
                window['-E-'].update("")
                window['-PV-'].update("")
                window['-P-'].update("")
                window['-C-'].update("")
                window['-A-'].update("")
                window['-S-'].update("")
                window['-CA-'].update("000")
            else:
                sg.popup_error('Valores inválidos. Verifique os requisitos.')
        elif '+CLICKED+' in event:
            selRow = event[2][0] #retorno padrão de evento para armanezar a linha selecionada, não me pergunte porque
            if selRow is not None:
                window['twitter'].update(selRow+1) #twitter é a chave para a linha selecionada na interface, calma
            else:
                selRow=0
                window['twitter'].update(selRow)
        elif event == '-INS-':
            print('Soon...')
            print("Finalmente chegou!!!")
            if selRow is not None:
                x = values['-X-']
                y = values['-Y-']
                z = int(values['-Z-'])
                e = values['-E-']
                pv = values['-PV-']
                p = values['-P-']
                c = values['-C-']
                a = values['-A-'] #00
                m = 00
                s = values['-S-'] # 
                ca = values['-CA-']
                if len(x) == 6 and x.isnumeric() and len(y) == 7 and y.isnumeric() and z == 23 \
                    and len(e) == 3 and len(pv) == 3 and p.isnumeric() and c.isnumeric() \
                    and 0 <= int(a) < 120 and s in ['Direito', 'Esquerdo', ''] and ca in ['000', '003', '004', '005', '007', '008', '010']:
                    newData=[x, y, z, e, pv, p, c, a, m, s, ca]
                    data.insert(selRow,newData)
                    window['-TABLE-'].update(values=data)
                    
                    window['-X-'].update("")
                    window['-Y-'].update("")
                    window['-Z-'].update(23)
                    window['-E-'].update("")
                    window['-PV-'].update("")
                    window['-P-'].update("")
                    window['-C-'].update("")
                    window['-A-'].update("")
                    window['-S-'].update("")
                    window['-CA-'].update("000")
        elif event == "-EDIT-":
            if selRow is not None:
                x = values['-X-']
                y = values['-Y-']
                z = int(values['-Z-'])
                e = values['-E-']
                pv = values['-PV-']
                p = values['-P-']
                c = values['-C-']
                a = values['-A-'] #00
                m = 00
                s = values['-S-'] # 
                ca = values['-CA-']

                newData = [x, y, z, e, pv, p, c, a, m, s, ca]
                
                oldData = data[selRow]
                
                for n in range(len(newData)):
                    if newData[n]!=oldData[n]: 
                        if newData[n]=='' or newData[n]==23:
                            continue
                        else:
                            oldData[n]=newData[n]

                data[selRow]=oldData
                window['-TABLE-'].update(values=data)
            else:
                sg.popup_error("Nenhuma Linha Selecionada")
        elif event == '-DEL-':
            if selRow is not None:
                print(selRow)
                del data[selRow]
                window['-TABLE-'].update(values=data)
            else:
                sg.popup_error("Nenhuma Linha Selecionada")
        #antigo if para edição individual das celulas, se mantido na identação ele buga os eventos de baixo(???)    
            """elif isinstance(event, tuple):
            if event[0] == '-TABLE-' and isinstance(event[2][0], int) and event[2][0] > -1:
                cell = row, col = event[2]
                #print(row)
                # Displays that coordinates of the cell that was clicked on
                window['-CLICKED_CELL-'].update(cell)
                edit_cell(window, '-TABLE-', row+1, col, data, justify='right')
            elif event[0] == '-TABLE-' and event[1] == 'Headings Click':
                # Do nothing when table headings are clicked
                pass
            """        
        elif event == 'Salvar':
            if len(data) > 0:
                filename = sg.popup_get_file('Selecione onde Salvar!', save_as=True,file_types=(("CSV Files", "*.csv"),))
                if filename:
                    with open(filename, 'w', newline='') as csvfile:
                        csvwriter = csv.writer(csvfile)

                        csvwriter.writerow([name, city, top, ns])
                        csvwriter.writerow([])

                        csvwriter.writerow(headings)
                        csvwriter.writerows(data)

                    sg.popup(f'Dados salvos com sucesso no nome de {ns}.csv')
            else:
                sg.popup_error('Erro ao salvar! 404 dados não encontrados!')
        elif event == 'Exportar':
            #[12, 12, 2, 3, 3, 6, 6, 3, 2, 1, 3]
            if len(data) > 0:
                try:
                    filename = sg.popup_get_file('Selecione onde Salvar!', save_as=True,file_types=(("Text Files", "*.txt"),))
                    if filename:
                        with open(filename, 'w') as txtfile:
                            for row in data:
                                row[0], row[1] = row[1], row[0]
                                contar = 1
                                formatted_row = []
                                for idx, (value, size) in enumerate(zip(row, [12, 12, 2, 3, 3, 9, 9, 3, 2, 1, 3])):
                                    # Tratamento especial para o valor na coluna 10
                                    if idx == 9:  # Verifica se é a décima coluna (índice 9)
                                        if value == 'Direito':
                                            formatted_value = 'D'.rjust(size)
                                        elif value == 'Esquerdo':
                                            formatted_value = 'E'.rjust(size)
                                        elif value == '':
                                            formatted_value = ' '.rjust(size)
                                        else:
                                            formatted_value = value.rjust(size)
                                    # Formatação para outros valores
                                    elif idx == 4:
                                        formatted_value = str(value).zfill(3).rjust(size + 1)
                                    elif idx == 3 or idx == 7 or idx == 10:
                                        formatted_value = str(value).zfill(3).rjust(size)
                                    elif idx == 2:
                                        formatted_value = str(value).zfill(2).rjust(size)
                                    elif idx == 5 or idx == 6:
                                        formatted_value = '{:.2f}'.format(float(value)).zfill(9).rjust(size)
                                    elif idx == 8:
                                        formatted_value = str(int(float(value))).zfill(2).rjust(size)
                                    elif idx == 0 or idx == 1:
                                        formatted_value = '{:.2f}'.format(float(value)).zfill(12).rjust(size)
                                    formatted_row.append(formatted_value)
                                txtfile.write(' '.join(formatted_row) + '\n')
                        sg.popup(f'Dados exportados com sucesso para {ns}.txt')
                except Exception as e:
                    sg.popup_error(f'Erro ao exportar dados: {e}')
            else:
                sg.popup_error('Erro ao exportar! Nenhum dado encontrado para exportar!')
        elif event == 'PDF':
            if len(data) > 0:
                try:
                    filename = sg.popup_get_file('Selecione onde Salvar PDF!', save_as=True, file_types=(("PDF Files", "*.pdf"),))
                    if filename:
                        # Generate PDF

                        pdf = FPDF(orientation='L')
                        pdf.add_page()
                        pdf.set_font("Arial", size=12)

                        # Add client information
                        client_info = f"Cliente: {name}  |  NS: {ns}"
                        topogr_into = f"Topógrafo: {top} |  Cidade: {city}"  #Código: {get_city_code(city)}"
                        pdf.cell(300, 10, txt=client_info, ln=True, align='C')
                        pdf.cell(300, 10, txt=topogr_into, ln=True, align='C')
                        pdf.ln(10)  # Add a new line for spacing

                        # Calculate column widths based on the content length
                        column_widths = []
                        for header, col_data in zip(headings, zip(*data, headings)):
                            max_width = max(pdf.get_string_width(str(cell)) for cell in col_data)
                            column_widths.append(max(max_width, pdf.get_string_width(header)))

                        # Add table headers
                        pdf.set_fill_color(255, 183, 0)
                        for header, width in zip(headings, column_widths):
                            pdf.cell(width + 5, 6, txt=header, border=2, align='C', fill=True)
                        pdf.ln()
                        
                        # Add table data
                        for row in data:
                            if contar == 1:
                                row[0], row[1] = row[1], row[0]
                            else:
                                row[0], row[1] = row[0], row[1]

                            for cell, width in zip(row, column_widths):
                                pdf.cell(width + 5, 6, txt=str(cell), border=2, align='C')
                            pdf.ln()

                        # Save the PDF to the specified filename
                        pdf.output(filename)
                        sg.popup(f'Dados exportados com sucesso para {filename}')
                except Exception as e:
                    sg.popup_error(f'Erro ao exportar dados para PDF: {e}')
            else:
                sg.popup_error('Erro ao exportar para PDF! Nenhum dado encontrado para exportar!')
        elif event == 'Voltar':
            window.close()
            main_menu()
    window.close()

main_menu()