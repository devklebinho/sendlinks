import openpyxl

wb = openpyxl.load_workbook("item.xlsx")

selected_sheet_name = wb.sheetnames[2]
worksheet = wb[selected_sheet_name]


def read_cells(line, columns):  # read each cell in line and save it in a list
    selected_line = line
    columns = columns
    mensage_variables = []

    for column in columns:
        selected_column = column
        selection = selected_column + selected_line
        cell = worksheet[selection].value

        # validation: no empty cell
        cell = str(cell)
        if cell != 'None':
            mensage_variables.append(cell)
        else:
            print("célula vazia")
    # print(mensage_variables)
    return mensage_variables


# -------------FORMATAÇÃO DA MENSAGEM PARA WHATSAPP------------
def format_message(mensage_variables):
    if mensage_variables != []:
        curso = mensage_variables[0]
        dia = mensage_variables[1]
        data = mensage_variables[2]
        horario = mensage_variables[3]
        #teacher = mensage_variables[4]
        linkZoom = mensage_variables[5]
        idDaReuniao = mensage_variables[6]
        senhaZoom = mensage_variables[7]
        aluno = mensage_variables[8]
        #responsavel = mensage_variables[9]
        #whatsapp = mensage_variables[10]

        whatsapp_mensage = 'Prezados Pais, bom dia!\nSegue link da aula ON LINE do(a) ' + aluno + ' desta semana:\n\n' + curso + '\n' + dia + ' ' + data + ' às ' + horario + '\n\nLink: '+linkZoom+'\nID DA REUNIÃO: ' + idDaReuniao + '\nSENHA: ' + senhaZoom + \
            '\n\n*Comunicado importante:*\n\nTemos o maior interesse no bom aprendizado de todos, e, para isto, solicitamos e contamos com a habitual compreensão e parceria no cumprimento desse horário para que tenham um bom desempenho sem perda de conteúdo. Como sabemos que imprevistos acontecem, teremos tolerância de no máximo 20 minutos de atraso.\n\nContamos com a compreensão de todos!\n\nIvanilda\nAtenciosamente,\nEquipe SuperGeeks Recife\n'

        # print(whatsapp_mensage)
        return whatsapp_mensage
    else:
        print('Mensagem vazia!')
        return ' '
