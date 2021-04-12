import openpyxl
import worksheetreader
import whatsbot


def extact_message():
    # range obrigadório a partir de 3(não existe linha zero no excel e a linha 1 normalmente é título)
    for line in range(4):
        # print(i)
        if line != 0 and line != 1:
            line = str(line)
            data_list = worksheetreader.read_cells(
                line, ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K'])
            message = worksheetreader.format_message(data_list)
            # contato = data_list[10]
    return message


message = ''

contato = 'Testes Automatizados'
whatsbot.abrir_whatsapp()
whatsbot.buscar_contato(contato)
whatsbot.enviar_mensagem(message)
