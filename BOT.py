
import PySimpleGUI as sg

sg.theme('DarkAmber')   # Add a touch of color
# All the stuff inside your window.

btn_active = True

def listafuncoes():
    listar = [
        "TERMO - GERAR TERMO DE ENTREGA/RETIRADA", 
        "PROCESSAR - FECHAR O.S. EM ABERTO COMO CONCLUIDAS", 
        #"ENVIO - PREENCHER O.S. DE ENVIO DE MAQUINAS",
        #"RETIRADA - PREENCHER O.S. DE RETIRADA DE MAQUINAS"
        ]
    return sorted(listar)
    

layout = [

            # FORMULARIO DE LOGIN DA APLICAÇÃO
            [sg.Text('INSIRA OS DADOS DE AUTENTICAÇÃO DO SCAT')],
            [sg.Text('USUÁRIO'), sg.Combo("BRASILCENTER", default_value="BRASILCENTER", s=(20,22), enable_events=True, readonly=True, k='SITE_CBX'), sg.InputText(s=(26,22), focus=True, k='USUARIO_TXB')],
            [sg.Text('SENHA   '), sg.InputText(s=(50,22), password_char="•", k='SENHA_TXB')],
            [sg.Text('')],

            # SELETOR DA FUNÇÃO DO BOT
            [sg.Text('SELECIONE A FUNÇÃO DESEJADA:')],
            [sg.Combo(listafuncoes(), s=(58,0), enable_events=True, readonly=True, k='FUNCAO_CBX',)],
            [sg.Text('')],

            # SELETOR DA PLANILHA DE DADOS
            {sg.Text('SELECIONE A PLANILHA DE DADOS:')},
            [sg.InputText(s=(51,22),key='-file2-'),sg.FileBrowse(target='-file2-',file_types=(('Planilhas do Excel', '*.xlsx'),))],
            [sg.Text('')],

            # ACEITE DE RESPOSABILIZAÇÃO PELAS AÇÕES AUTOMATICAS
            [sg.Checkbox('Me responsabilizo INTEGRALMENTE pelas ações a serem realizadas', enable_events=True,  k='-RESPBACT-', p=0)],
            [sg.Text('')],

            # BOTÕES DE PROCESSAMENTO
            [sg.Button('ENTRAR E PROCESSAR', enable_events=True,  k='BTN_ENTRAR', s=(52,2), button_color=('grey'),disabled_button_color=('grey'))],
            
            # INDICADOR VISUAL DO PROGRESSO | PROGRESS-BAR
            [sg.Text('PROGRESSO')],
            [sg.ProgressBar(10, orientation='h', s=(39,25), k='-PBAR-',bar_color=('green'))],
]

# Create the Window
window = sg.Window('SCABOT - CAS RPO', layout)
# Event Loop to process "events" and get the "values" of the inputs

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Cancel': # if user closes window or clicks cancel
        break
    #print('You entered ', values[0])

    # EVENTO DE LEITURA DO CHECKBOX DE ACEITE DE RESPONSABILIDADE
    if event == '-RESPBACT-':
        btn_active =  not btn_active
        print(btn_active)
        # ATUALIZA UMA PROPRIEDADE DO BOTÃO ENTRAR
        window['BTN_ENTRAR'].update(disabled=btn_active)
        if(btn_active != True):
            window['BTN_ENTRAR'].update(button_color= '#fdcb52')
        else:
            window['BTN_ENTRAR'].update(button_color= 'grey')
    # EVENTO DO BOTÃO DE ENTRAR E PROCESSAR
    if event == 'BTN_ENTRAR':
        window['SITE_CBX'].update(disabled=True)
        window['USUARIO_TXB'].update(text_color='black', disabled=True)
        window['SENHA_TXB'].update(text_color='black', disabled=True)
        window['FUNCAO_CBX'].update(disabled=True)
        window['BTN_ENTRAR'].update(disabled=True, button_color= 'grey')

window.close()