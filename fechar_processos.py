import os

import psutil


def fechar_processo_python_anterior():
    caminho_relativo = os.path.dirname(__file__)
    with open(os.path.join(caminho_relativo, "processo_anterior.txt"), "r") as arquivo:
        last_pid = int(arquivo.read())

    try:
        parent_process = psutil.Process(last_pid)
    except:
        return 'Processo anterior já foi finalizado'
    print('Processo anterior ainda está aberto')
    child_processes = parent_process.children(recursive=True)

    parent_process.kill()
    for child in child_processes:
        try:
            child.kill()
        except:
            print(child.name() + ' Não foi finalizado')

    return 'Processo anterior finalizado com sucesso!'

def fechar_processos(lista_dos_caminhos_dos_processos, lista_dos_nomes_de_processos):
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
