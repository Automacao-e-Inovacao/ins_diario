import logging
import os

caminho_relativo = os.path.dirname(__file__)

log_format = '''\n\n\n\nTempo do evento: %(asctime)s
            \nNome da funcao da chamada: %(funcName)s\nNome do level: %(levelname)s\nLinha do evento: 
            \nDescricao do evento: %(message)s\nNome do arquivo: %(filename)s
            \nTempo corrido em milissegundos: %(msecs)d\nNome do processo%(processName)s'''

logging.basicConfig(filename=os.path.join(caminho_relativo, 'logs.log'),
                    filemode='w',
                    level=logging.WARNING,
                    format=log_format)

logger = logging.getLogger(__name__)
try:
    from win32clipboard import OpenClipboard, CloseClipboard, GetClipboardData, EmptyClipboard
    from time import sleep
    import datetime
    from ahk import AHK
    from fechar_processos import fechar_processo_python_anterior, fechar_processos
    print(fechar_processo_python_anterior())

    with open(os.path.join(caminho_relativo, "processo_anterior.txt"), "w") as arquivo:
        arquivo.write(str(os.getpid()))

    ahk = AHK(
        executable_path=os.path.dirname(os.path.realpath(__file__)) + '\\venv\\' + '\\Autohotkey\\' + 'AutoHotkey.exe')
    from scripts_ccsupervision import esperar, esperar_janela_Edit, Abrir_CCsupervision, Aba_estatisticas, \
        login_ccsupervision, verificar_ip, mudar_ip, Conexao_postgresql, fechar_janela_ACD_e_TEMP
    from scripts_excel import Manipulacao, Manipulacao_2

    caminho = r"C:\ProgramData\Alcatel\CCSupervisor\Excel\TEMP.XLSM"

    lista_dos_caminhos_dos_processos = [os.path.join(caminho_relativo, 'scripts_ccsupervision.py'),
                                        os.path.join(caminho_relativo,
                                                     'scripts_excel.py')]
    lista_dos_nomes_de_processos = ['ccs.exe', 'gsw32.exe', 'EXCEL.EXE', 'AutoHotkey.exe']

    db = Conexao_postgresql('10.6.2.211', 'datamart', '5507392', 'NEFARIAN@1654')

    dia_anterior = datetime.datetime.now() - datetime.timedelta(days=1)

    parte_do_titulo = 'Pilotos'
    cabecalho = ['INS 30seg', 'IAB 30seg', 'ICO', 'ChOf',
                 'ChA', 'TMA']
    qtd_insercao_db = ''
    for indice_qtd in range(len(cabecalho) + 2):
        if indice_qtd != len(cabecalho) + 1:
            qtd_insercao_db += r'%s,'
        else:
            qtd_insercao_db += r'%s'
    tabela = 'insdiarioresumo'
    tabela_2 = 'insdiario'
    cabecalho_2 = [
        'Dia',
        'Perini',
        'Perfim',
        'Tipicidade',
        'ChA<30',
        'ChA>30',
        'Chab<30',
        'Chab>30',
        'ChOc',
        'ChOf',
        'TMA',
        'TME',
        'Diferença',
        'INS',
        'IAb',
        'ICO',
        'Limite'
    ]
    qtd_insercao_db_2 = ''
    for indice_qtd in range(len(cabecalho_2) + 2):
        if indice_qtd != len(cabecalho_2) + 1:
            qtd_insercao_db_2 += r'%s,'
        else:
            qtd_insercao_db_2 += r'%s'
    usuario = 'Everton'
    senha = 12345678

    lista_pilotos_disponiveis = ["eqtl-ma-116", "eqtl-pa-0800", "eqtl-pi-0800", "eqtl-al-0800", "eqtl-ap-116", "csa-ap",
                                 "gc-ma", "gc-pa", "gc-pi", "gc-al", "gc-ap", "tel ma", "tel pa", "tel pi", "tel al",
                                 "tel ap", "ouv-ma", "ouv-pa", "ouv-pi", "ouv-al", "ouv-ap", "ag-ma", "ag-pa", "ag-pi",
                                 "ag-al", "ag-ap", "gd-ma", "gd-pa", "gd-pi", "gd-al", "ng-ma-eqtl", "ng-pa-eqtl"]
    lista_formatos_disponiveis = ["eqtl ma", "eqtl pa", "eqtl pi", "eqtl al", "eqtl ap", "torp ilhas 30s"]

    (x, y) = (462, 121)


    def Janela_edicao(piloto_disponivel, formato_disponivel, dia, mes, ano):
        ahk.run_script(f'''
                        piloto_disponivel:= "{piloto_disponivel}"
                        formato_disponivel:= "{formato_disponivel}"
                        dia:= "{dia}"
                        mes:= "{mes}"
                        ano:= "{ano}"
                        ''' +
                       '''
                        SetControlDelay - 1
                        controlgettext, data_completa_da_janela_edit, SysDateTimePick321, Edi
                        lista_provisoria:= strsplit(data_completa_da_janela_edit, "/")
                        if substr(dia, 1, 1) = 0
                        {
                            dia:= substr(dia, 2)
                        }
                        if substr(lista_provisoria[1], 1, 1) = 0
                        {
                            lista_provisoria[1]:= substr(lista_provisoria[1], 2)
                        }
                        if substr(mes, 1, 1) = 0
                        {
                            mes:= substr(mes, 2)
                        }
                        if substr(lista_provisoria[2], 1, 1) = 0
                        {
                            lista_provisoria[2]:= substr(lista_provisoria[2], 2)
                        }
                        loop
                        {
                            if lista_provisoria[1] = dia and lista_provisoria[2] = mes and lista_provisoria[3] = ano
                            {
                                break
                            }
                            else
                            {
                                if lista_provisoria[3] > ano
                                {
                                    ControlClick, x142 y433, Edi
                                    sleep, 500
                                    Controlsend, SysDateTimePick321, {down}, Edi
                                }
                                else
                                {
                                    if lista_provisoria[3] < ano
                                    {
                                        ControlClick, x142 y433, Edi
                                        sleep, 500
                                        Controlsend, SysDateTimePick321, {up}, Edi
                                    }
                                }
                                if lista_provisoria[2] > mes
                                {
                                    ControlClick, x119 y433, Edi
                                    sleep, 500
                                    Controlsend, SysDateTimePick321, {down}, Edi
                                }
                                else
                                {
                                    if lista_provisoria[2] < mes
                                    {
                                        ControlClick, x119 y433, Edi
                                        sleep, 500
                                        Controlsend, SysDateTimePick321, {up}, Edi
                                    }
                                }
                                if lista_provisoria[1] > dia
                                {
                                    ControlClick, x102 y433, Edi
                                    sleep, 500
                                    Controlsend, SysDateTimePick321, {down}, Edi
                                }
                                else
                                {
                                    if lista_provisoria[1] < dia
                                    {
                                        ControlClick, x102 y433, Edi
                                        sleep, 500
                                        Controlsend, SysDateTimePick321, {up}, Edi
                                    }
                                }
                            }
                            controlgettext, data_completa_da_janela_edit, SysDateTimePick321, Edi
                            lista_provisoria:= strsplit(data_completa_da_janela_edit, "/")
                            if substr(dia, 1, 1) = 0
                            {
                                dia:= substr(dia, 2)
                            }
                            if substr(lista_provisoria[1], 1, 1) = 0
                            {
                                lista_provisoria[1]:= substr(lista_provisoria[1], 2)
                            }
                            if substr(mes, 1, 1) = 0
                            {
                                mes:= substr(mes, 2)
                            }
                            if substr(lista_provisoria[2], 1, 1) = 0
                            {
                                lista_provisoria[2]:= substr(lista_provisoria[2], 2)
                            }
                        }
                        PostMessage , 0x0185 , 1 , -1 , ListBox2 , Edi
                        ControlClick, Button27, Edi
                        PostMessage , 0x0185 , 1 , -1 , ListBox1 , Edi
                        ControlClick, Button8, Edi
                        PostMessage , 0x0185 , 0 , -1 , ListBox5 , Edi
                        PostMessage , 0x0185 , 0 , -1 , ListBox6 , Edi
                        Control, Check,,Button1, Edi
                        Control, Check,, Button3, Edi
                        Control, Check,, Button17, Edi
                        Control, Check,, Button14, Edi
                        Control, UnCheck,, Button37, Edi
                        Control, Uncheck,, Button12, Edi
                        Control, Uncheck,, Button13, Edi
                        Control, Uncheck,, Button38, Edi
                        Control, Uncheck,, Button30, Edi
                        Control, Uncheck,, Button39, Edi
                        Control, ChooseString, %piloto_disponivel% , ListBox5, Edi
                        Control, ChooseString, %formato_disponivel% , ListBox6, Edi
                        ControlClick, Button20, Edi
                    ''', blocking=True)


    fechar_processos(lista_dos_caminhos_dos_processos, lista_dos_nomes_de_processos)
    Abrir_CCsupervision(ahk)

    while 1 == 1:
        try:
            OpenClipboard()
            EmptyClipboard()
            CloseClipboard()
            sleep(2)
            verificar_ip(ahk)
            OpenClipboard()
            clipboard_texto = GetClipboardData()
            print(str(clipboard_texto))
            if 'CCsupervision' in clipboard_texto:
                if not '10.101.19.27' in clipboard_texto:
                    try:
                        if ahk.win_get('Iniciar sess').exist:
                            (x_mudar_ip, y_mudar_ip, x_mudar_ip1, y_mudar_ip1) = 157, 38, 196, 137
                        else:
                            (x_mudar_ip, y_mudar_ip, x_mudar_ip1, y_mudar_ip1) = 473, 39, 526, 304
                    except:
                        if ahk.win_get('Iniciar sess'):
                            (x_mudar_ip, y_mudar_ip, x_mudar_ip1, y_mudar_ip1) = 157, 38, 196, 137
                        else:
                            (x_mudar_ip, y_mudar_ip, x_mudar_ip1, y_mudar_ip1) = 473, 39, 526, 304
                    mudar_ip(ahk, x_mudar_ip, y_mudar_ip, x_mudar_ip1, y_mudar_ip1)
                    fechar_processos(lista_dos_caminhos_dos_processos, lista_dos_nomes_de_processos)
                    Abrir_CCsupervision(ahk)
                break
            CloseClipboard()
        except Exception as exc:
            try:
                EmptyClipboard()
            except:
                pass
            try:
                CloseClipboard()
            except:
                pass
    try:
        CloseClipboard()
    except:
        pass

    login_ccsupervision(ahk, usuario, senha)

    Aba_estatisticas(ahk, x, y, parte_do_titulo)
    esperar_janela_Edit(ahk)
    for contagem_lista_pilotos_disponiveis, piloto_disponivel in enumerate(lista_pilotos_disponiveis):

        if contagem_lista_pilotos_disponiveis <= len(lista_formatos_disponiveis) - 1:
            formato_disponivel = lista_formatos_disponiveis[contagem_lista_pilotos_disponiveis]

        fechar_janela_ACD_e_TEMP(ahk)

        Janela_edicao(piloto_disponivel, formato_disponivel, dia_anterior.day, dia_anterior.month, dia_anterior.year)

        esperar(ahk)

        matriz_principal, matriz = Manipulacao(caminho, cabecalho, piloto_disponivel, dia_anterior)
        db.manipular(f"""insert into "{tabela}" values ({qtd_insercao_db})""", matriz_principal)
        matriz_principal_2 = Manipulacao_2(cabecalho_2, piloto_disponivel, dia_anterior, matriz)
        for linha_2 in matriz_principal_2:
            db.manipular(f"""insert into "{tabela_2}" values ({qtd_insercao_db_2})""", linha_2)

    fechar_processos(lista_dos_caminhos_dos_processos, lista_dos_nomes_de_processos)
    Abrir_CCsupervision(ahk)
    usuario = 'poa_rpa'
    senha = 'Rpa@1234'

    while 1 == 1:
        try:
            OpenClipboard()
            EmptyClipboard()
            CloseClipboard()
            sleep(2)
            verificar_ip(ahk)
            OpenClipboard()
            clipboard_texto = GetClipboardData()
            print(str(clipboard_texto))
            if 'Aplicativo CCsupervision' in clipboard_texto:
                if not '10.48.98.70' in clipboard_texto:
                    try:
                        if ahk.win_get('Iniciar sess').exist:
                            (x_mudar_ip, y_mudar_ip, x_mudar_ip1, y_mudar_ip1) = 157, 38, 196, 137
                        else:
                            (x_mudar_ip, y_mudar_ip, x_mudar_ip1, y_mudar_ip1) = 473, 39, 526, 304
                    except:
                        if ahk.win_get('Iniciar sess'):
                            (x_mudar_ip, y_mudar_ip, x_mudar_ip1, y_mudar_ip1) = 157, 38, 196, 137
                        else:
                            (x_mudar_ip, y_mudar_ip, x_mudar_ip1, y_mudar_ip1) = 473, 39, 526, 304
                    mudar_ip(ahk, x_mudar_ip, y_mudar_ip, x_mudar_ip1, y_mudar_ip1)
                    fechar_processos(lista_dos_caminhos_dos_processos, lista_dos_nomes_de_processos)
                    Abrir_CCsupervision(ahk)
                break
            CloseClipboard()
        except Exception as exc:
            try:
                EmptyClipboard()
            except:
                pass
            try:
                CloseClipboard()
            except:
                pass
    try:
        CloseClipboard()
    except:
        pass
    login_ccsupervision(ahk, usuario, senha)
    sleep(20)
    Aba_estatisticas(ahk, x, y, parte_do_titulo)
    esperar_janela_Edit(ahk)
    lista_pilotos_disponiveis = ['eqtl-rs-0800', 'ouv-rs']
    lista_formatos_disponiveis = ['eqtl rs', 'torp ilhas 30s']
    for contagem_lista_pilotos_disponiveis, piloto_disponivel in enumerate(lista_pilotos_disponiveis):

        formato_disponivel = lista_formatos_disponiveis[contagem_lista_pilotos_disponiveis]

        fechar_janela_ACD_e_TEMP(ahk)

        Janela_edicao(piloto_disponivel, formato_disponivel, dia_anterior.day, dia_anterior.month, dia_anterior.year)
        esperar(ahk)

        matriz_principal, matriz = Manipulacao(caminho, cabecalho, piloto_disponivel, dia_anterior)
        db.manipular(f"""insert into "{tabela}" values ({qtd_insercao_db})""", matriz_principal)
        matriz_principal_2 = Manipulacao_2(cabecalho_2, piloto_disponivel, dia_anterior, matriz)
        for linha_2 in matriz_principal_2:
            db.manipular(f"""insert into "{tabela_2}" values ({qtd_insercao_db_2})""", linha_2)

except Exception as e:
    logger.error(e)
