from psycopg2 import connect
import psutil


class Conexao_postgresql(object):
    def __init__(self, mhost, db, usr, pwd):
        self._db = connect(host=mhost, database=db, user=usr, password=pwd)

    def manipular(self, sql, _Vars):
        cur = self._db.cursor()
        cur.execute(sql, _Vars)
        cur.close()
        self._db.commit()

    def query(self, sql):
        cur = self._db.cursor()
        cur.execute(sql)
        cur.close()
        self._db.commit()

    def consultar(self, sql):
        rs = None
        cur = self._db.cursor()
        cur.execute(sql)
        rs = cur.fetchall()
        return rs

    def proximaPK(self, tabela, chave):
        sql = 'select max(' + chave + ') from ' + tabela
        rs = self.consultar(sql)
        pk = rs[0][0]
        if pk == None:
            return 0
        else:
            return pk + 1

    def fechar(self):
        self._db.close()


def fechar_janela_ACD_e_TEMP(ahk):
    ahk.run_script(
        '''
    loop
    {
        ifwinexist, ACD
        {
            winclose, ACD
        }
        else
        {
            break
        }
    }
    loop
    {
        ifwinexist, ahk_exe EXCEL.exe
        {
            winkill, ahk_exe EXCEL.exe
        }
        else
        {
            break
        }
    }
    ''', blocking=True)


def fechar_tudo(lista_dos_caminhos_dos_processos, lista_dos_nomes_de_processos):
    for proc in psutil.process_iter():
        try:
            for linha in lista_dos_caminhos_dos_processos:
                if linha in proc.cmdline():
                    proc.kill()
                    continue
            for linha in lista_dos_nomes_de_processos:
                if linha in proc.exe():
                    proc.kill()
                    continue
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass


def Abrir_CCsupervision(ahk):
    ahk.run_script('''
                   ifwinnotexist, ahk_exe ccs.exe
                   {
                        run, C:\Program Files (x86)\Alcatel\A4400 Call Center Supervisor\ccs.exe
                        winwait, ahk_exe ccs.exe
                        winwait, Excel
                        winwait, Iniciar sess
                        loop
                        {
                            ifwinexist, ahk_exe EXCEL.exe
                            {
                                winkill, ahk_exe EXCEL.exe
                            }
                            else
                            {
                                break
                            }
                        }
                   }
                   ''', blocking=True)


def login_ccsupervision(ahk, usuario, senha):
    ahk.run_script(f'''
                    usuario:= "{usuario}"
                    senha:= "{senha}"
                    '''
                   +
                   '''
                    ControlSetText, Edit1, %usuario%,Iniciar sess
                    ControlSetText, Edit2, %senha%,Iniciar sess
                    SetControlDelay - 1
                    loop
                    {
                        ifwinexist, Iniciar sess
                        {
                            ControlClick, Button1, Iniciar sess
                            sleep, 5000
                        }
                        else
                        {
                            break
                        }
                    }
                   ''', blocking=True)


def Aba_estatisticas(ahk, x, y, titulo):
    ahk.run_script(f'''
    titulo_recebido:= "{titulo}"
    x:= {x}
    y:= {y}
    ''' +
                   '''
    loop
    {
        ifwinexist, Informa
        {
            winkill, Informa
        }
        else
        {
            break
        }
    }
    WinGetTitle, titulo, ahk_exe ccs.exe
    IfNotInString, titulo, % titulo_recebido
    {
        winactivate, CCsupervision
        winwaitactive, CCsupervision
        click, 170, 39
        sleep, 500
        mousemove, 230, 96
        sleep, 500
        click, %x%, %y%
    }
    ''', blocking=True)


def esperar(ahk):
    ahk.run_script('''
    loop
    {
        winwait, TEMP,,10
        ifwinexist, TEMP
        {
            break
        }
    }
    sleep, 5000
    ''', blocking=True)


def esperar_janela_Edit(ahk):
    ahk.run_script('''
                winwait, ahk_exe ccs.exe ahk_class #32770
                winactivate, ahk_exe ccs.exe ahk_class #32770 
                ''', blocking=True)


def verificar_ip(ahk):
    ahk.run_script(
        '''
                    loop
                    {
                        ifwinexist, Edi
                        {
                            winkill, Edi
                        }
                        else
                        {
                            break
                        }
                    }
                    winactivate, CCsupervision
                    winwaitactive, CCsupervision
                    click, 512, 67
                    winwait, Informa
                    winwaitactive, Informa
                    sleep, 1000
                    loop
                    {
                        try
                        {
                            WinGetText, Clipboard, Informa
                            break
                        }
                        catch
                        {
                            
                        }
                    }
                    winclose, Informa
                   ''', blocking=False)


def mudar_ip(ahk, x, y, x1, y1):
    ahk.run_script(f'''
                    x:= {x}
                    y:= {y}
                    x1:= {x1}
                    y1:= {y1}
                    '''
                   +
                   '''
                    SetControlDelay - 1
                    loop
                    {
                        ifwinexist, Informa
                        {
                            winkill, Informa
                        }
                        else
                        {
                            break
                        }
                    }
                    winactivate, CCsupervision
                    winwaitactive, CCsupervision
                    click, %x%, %y%
                    sleep, 500
                    click, %x1%, %y1%
                    winwait, Personaliza
                    sleep, 500
                    winactivate, Personaliza
                    loop
                    {
                        click, 320, 316
                        pixelsearch,,, 222, 311, 699, 323, 0xFFFFFF, 0, fast rgb
                    }
                    until (errorlevel = 0)
                    click, 203, 328
                    sleep, 500
                    ControlClick, Button10, Personaliza
                    winwait, CCS
                    ControlClick, Button2, CCS
                    winwait, ACD
                    winclose, ACD
                    winwaitactive, Personaliza
                    ControlClick, Button1, Personaliza
                   ''')
