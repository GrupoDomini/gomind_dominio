import math
import re
import time
from typing import Literal, Union, Self
from gomind_excel import converter_xls_para_xlsx
import pyautogui as py
import gomind_automation as automation
from time import sleep
from gomind_web_browser import WebBrowser
import pyperclip
import os


class DominioWeb(WebBrowser):
    def __init__(self) -> None:
        super().__init__()

    def run_login_steps(self, user: str, password: str) -> None:
        print("Vai rodar os passos")
        self.open().login(user, password).aceitar_plugin_dominio()

    def open(self) -> Self:
        print("[step] (open) Abrindo dominio web")
        self.nav.get("https://www.dominioweb.com.br/")
        self.nav.maximize_window()

        try:
            self.wait_for_element(
                '//*[@id="trid-auth"]/form/label[1]/span[2]/input', timeout=20
            )
        except Exception as e:
            print("Acesso ao Domínio Web não disponível", e)
            raise Exception("Acesso ao Domínio Web não disponível")

        return self

    def login(self, user: str, password: str) -> Self:
        print("[step] (login) Abrindo dominio web")

        self.send_keys('//*[@id="trid-auth"]/form/label[1]/span[2]/input', user)
        sleep(0.5)
        self.send_keys('//*[@id="trid-auth"]/form/label[2]/span[2]/input', password)
        sleep(0.5)
        self.click('//*[@id="enterButton"]')
        sleep(3)

        # Verificar se login ou senha estao incorretos
        teve_erro_no_login = self.verify_if_element_exists(
            "//*[contains(text(), 'inválidos')]"
        )

        if teve_erro_no_login:
            print("Usuário e/ou senha inválidos")
            self.encerrar_navegador()
            raise Exception("Erro ao fazer login no dominio web")

        else:
            sleep(2)
            trid_auth = '//*[@id="trid-auth"]/form/div[2]/div/button'

            if self.verify_if_element_exists(trid_auth):
                self.click(trid_auth)

            print("login DW = OK")
        return self

    def aceitar_plugin_dominio(self) -> Self:
        print("[step] (aceitar_plugin_dominio) Aceitando o plugin do dominio")

        try:
            """ Ativa Plug-in do Domínio Web """
            print("Tentando encontrar plugin dominioweb em portugues")
            automation.clicar_na_imagem_ate_sumir("plugin_dominioweb.png")
            return self
        except Exception as e:
            print("Erro ao aceitar plugin dominio")
            raise e


def login(user: str, password: str):
    """Logon Módulo - Domínio Web"""
    print("[step] (login) Fazendo login no dominio desktop")
    print("Fazendo login no Dominio Web Plugin")
    if not automation.esperar_imagem("user_dominiow.png", 80, 1):
        print("Plugin do Domínio Web não disponível")
        raise Exception("Plugin do Domínio Web não disponível")

    py.write(user)
    py.press("tab")
    sleep(1)
    py.write(password)
    py.press("enter")

    sleep(10)

    intervalo = 1
    tentativas = math.ceil(60 / intervalo)
    print("Esperando dominio abrir")
    for tentativa in range(tentativas):
        if automation.encontrar_imagem(
            "dominio_dark_background.png", ignorar_erro=True
        ):
            print(f"Dominio abriu na tentativa {tentativa}/{tentativas}")
            print("Finalizou login")
            return True

        print(f"Dominio nao abriu na tentativa {tentativa}/{tentativas}")
        forcar_escape()
        sleep(intervalo)

    raise Exception("Dominio nao abriu completamente apos o login")


def forcar_escape():
    print("Forcando escape")
    for _ in range(3):
        py.press(["esc", "n"] * 4)


def checar_dominioweb(segundos: int = 60, intervalo: float = 0.5):
    """Verificador de Domínio Web Desktop esta aberto"""
    return automation.esperar_imagem(
        "dominio_web_desktop_logo.png", segundos, intervalo
    )


def checar_dominio(segundos: int = 60, intervalo: float = 0.5):
    """Verificador de Domínio Sistemas aberto"""
    print("Verificando se dominio ja esta aberto")

    if automation.esperar_imagem("window_dominioweb.png", segundos, intervalo):
        print("Dominio ja esta aberto")
        return True

    print("Dominio nao esta aberto")
    return False


def escolher_modulo(coluna: int = 1, linha: int = 1) -> None:
    """Escolhe o modulo do Dominio Web a partir de uma coluna e uma linha,
    como Folha por exemplo

    Args:
        coluna (int): Comeca da coluna 1 adiante
        linha (int): Comeca da linha 1 adiante
    """
    if coluna == 1 and linha == 1:
        py.press("right", presses=1, interval=0.2)
        py.press("left", presses=1, interval=0.2)
    else:
        py.press("right", presses=coluna - 1, interval=0.2)
        py.press("down", presses=linha - 1, interval=0.2)
    time.sleep(1)
    py.press("enter", interval=0.2)
    py.hotkey("win", "d")


def fechar_dominio_sistemas():
    """Fechar o aplicativo - Domínio Sistemas"""
    # Janela principal        
    try:
        os.system('TASKKILL /F /IM ' + 'AppController.exe')
        os.system('TASKKILL /F /IM ' + 'gg-client.exe')
        os.system('TASKKILL /F /IM ' + 'UpdateService.exe')
        os.system('TASKKILL /F /IM ' + 'TRInternetMonitor.exe')
    except:
        pass


def alterar_cliente_dominio(codigo_cliente: Union[str, int]) -> bool:
    print("Vai trocar o cliente")

    if not codigo_cliente:
        raise Exception("O parâmetro 'codigo_cliente' é obrigatório")

    # region =========== AGUARDAR BUTTON TROCAR EMPRESA =======
    try:
        automation.clicar_na_imagem("btn_troca_dominioweb.png", 15, clicks=2)
    except Exception as _:
        print("Não foi possível selecionar o cliente no Domínio Sistemas")
        raise Exception("Não foi possível selecionar o cliente no Domínio Sistemas")
    # endregion

    # region =========== AGUARDAR BUTTON TROCAR CODIGO =======
    automation.clicar_na_imagem("cod_troca_dominioweb.png", 60, clicks=2)
    # endregion

    # region =========== TROCAR/ESCREVER CODIGO ===============
    sleep(1)
    py.write(str(codigo_cliente), interval=0.2)
    sleep(1)
    py.press("enter")
    # endregion

    sleep(5)
    forcar_escape()

    if not automation.esperar_imagem(
        "aviso1_dominioweb.png", segundos=3
    ):  # ==> VERIFICAR ERROS
        print(f"Existem parâmetros cadastrados para a empresa {codigo_cliente}")
        return True  # Retorna True se conseguiu trocar o código
    else:
        print(f"Não existem parâmetros cadastrados para a empresa {codigo_cliente}")
        return False  # Returna False se nao conseguiu trocar o codigo


def abrir_menu(letra: Literal["c", "a", "m", "r", "u", "f", "j"]):
    if not letra:
        raise Exception("Coluna é um parâmetro obrigatório")

    py.hotkey("alt", letra)


def baixar_relatorio_como_excel():
    print("Vai baixar o relatorio como excel")
    sleep(2)
    automation.clicar_na_imagem("btn_salvar_como_excel.png", clicks=3)

    try:
        # Adicionei esse clique para ter certeza de que vai abrir
        automation.clicar_na_imagem("btn_salvar_como_excel.png", segundos=5)
    except Exception as _:
        print("Erro ao clicar novamente no botao salvar como excel", "warning")

    sleep(10)  # Esperar tela abrir
    py.press("tab", presses=4, interval=0.5)
    py.press(["down", "r", "c", "c", "enter"], interval=0.5)
    sleep(2)
    py.press("tab", presses=3, interval=0.5)
    sleep(2)
    py.press(["r", "enter"], interval=2)
    sleep(2)
    print("Selecionou a pasta RPA")
    py.press("tab", presses=4, interval=0.7)
    sleep(2)
    py.press("enter")
    print("Baixou o relatorio como excel")
    sleep(10)  # Tempo para esperar o arquivo EXCEL baixar


def baixar_relatorio_como_xlsx(caminho: str, letra_do_cliente: str, converter_xlsx: bool = False):
    """A partir de um xls ele transformara em xlsx no mesmo caminho, apagando o arquivo xls no final do processo"""
    print("Vai baixar o relatorio como excel")
    sleep(2)
    automation.clicar_na_imagem("btn_salvar_como_excel.png", clicks=3, segundos=20)

    try:
        # Adicionei esse clique para ter certeza de que vai abrir
        automation.clicar_na_imagem("btn_salvar_como_excel.png", segundos=5)
    except Exception as _:
        print("Erro ao clicar novamente no botao salvar como excel", "warning")

    sleep(10)  # Esperar tela abrir
    automation.special_write(caminho)
    print("Selecionou a pasta RPA")
    py.press("tab", presses=2, interval=1)
    sleep(1)
    py.press("enter")
    print("Baixou o relatorio como excel")

    automation.esperar_imagem('icon_openofice.png', 140)
    #clica no centro da tela
    py.click(500, 500)
    time.sleep(1)
    py.hotkey('ctrl', 'q')

    tempo_inicial = time.time()
    while tempo_inicial - time.time() < 200:
        achou = automation.esperar_imagem("btn_salvar_como_excel.png")
        print("ACHOU ======= {}".format(achou))
        if not achou:
            break
        automation.clicar_na_imagem("btn_salvar_como_excel.png", clicks=3, segundos=20)
        time.sleep(5)
    else:
        raise Exception("Passou do tempo total de espera")

    caminho_cliente = re.sub(
        r"^\w:\\\\",
        letra_do_cliente.replace("\\", "\\\\"),
        caminho.replace("\\", "\\\\"),
    ).replace("\\\\", "\\")

    caminho_xlsx = caminho_cliente
    if converter_xlsx:
        caminho_xlsx = "{}.xlsx".format(caminho_cliente)
        caminho_xls = "{}.xls".format(caminho_cliente)
        converter_xls_para_xlsx(caminho_xls, caminho_xlsx, True)

    while True:
        position = automation.esperar_imagem("btn_salvar_como_excel.png")
        if position:
            break
        py.press("esc")
        sleep(1)

    return caminho_xlsx


def baixar_relatorio_como_pdf():
    print("Vai baixar o relatorio como pdf")
    py.hotkey("ctrl", "d", interval=5)
    py.press("down", presses=4, interval=0.2)
    py.press("enter", interval=0.3)

    automation.clicar_na_imagem("pasta_rpa_dominioweb.png", segundos=20)

    sleep(1)
    py.press("tab", presses=2, interval=0.4)
    py.press("enter", interval=0.3)
    print("Baixou o relatorio como pdf")
    if automation.esperar_imagem('localiza_pdf.png', 60, 0.99):
        automation.clicar_no_centro_da_tela() #--> Clica no centro da tela
        py.hotkey('ctrl', 'q')
    sleep(3)


def abrir_dominio(
    user_desktop,
    password_desktop,
    user_web: str,
    password_web: str,
    modulo_linha: int = 1,
    modulo_coluna=1,
    tentativas: int = 3,
):
    dominio_web = DominioWeb()

    try:
        dominio_web.run_login_steps(user_web, password_web)
    except Exception as e:
        print(e)
    finally:
        print("Encerrando o navegador")
        time.sleep(2)
        dominio_web.encerrar_navegador()

    if not checar_dominioweb(segundos=80):  # Verifica se o Plugin DominioWeb foi aberto
        fechar_dominio_sistemas()
        if tentativas > 0:
            tentativas -= 1
            abrir_dominio(user_desktop, password_desktop, user_web, password_web, modulo_linha, modulo_coluna, tentativas)
            return
        raise Exception("Não conseguiu abrir o dominio")

    print(f"Escolhendo modulo LINHA={modulo_linha}, COLUNA={modulo_coluna}")
    sleep(1)

    escolher_modulo(linha=modulo_linha, coluna=modulo_coluna)
    login(user_desktop, password_desktop)

    if not checar_dominio(10):
        print("Erro ao abrir o dominio", "error")
        raise Exception("Não conseguiu abrir o dominio")
    else:
        print("Dominio foi aberto corretamente")

    return 

def apura_geraparcela_dominio(mes_ant, ano):
    try:
        width_screen, height_screen = py.size()
        time.sleep(.5)
        py.doubleClick(width_screen /2, height_screen /2)

        automation.clicar_na_imagem('apuracao_dominioweb.png',20, 0.5)
        if btn_processa_periodo := automation.esperar_imagem('proce_apuracao_dominioweb.png', 10, 0.5):
            sleep(2)
            py.write(f"{mes_ant}{ano}")  # define período inicial de apuração
            sleep(.5)
            py.press("tab")
            sleep(.5)
            py.write(f"{mes_ant}{ano}")  # define período final de apuração
        
            automation.clicar_na_imagem('proce_apuracao_dominioweb.png', 3, 0.5, clicks=2)

        sleep(3)
        automation.esperar_imagem_sumir('processo_apuracao.png',180, 0.5)
        if automation.esperar_imagem('aviso_apuracao.png', 5, 0.5):
            py.press('esc', presses=3)
            sleep(1)
            print("Erro na Apuração do Período para a empresa")
        else:
            print("Apuração do Período para a empresa realizado com sucesso")
        
        while not automation.esperar_imagem("apuracao_selecao_periodo.png", 5, 0.5):
            py.press("n", presses=2)
            sleep(0.5)
        py.press(['n', 'esc'], presses= 5, interval= .3)
        print("Encerrou janela Apuração do Período")
    except:
        print("Erro na Apuração do Período para a empresa")

def apuracao_dominio(mes_ant, ano):
    automation.clicar_na_imagem('apuracao_dominioweb.png',20, 0.5)
    automation.esperar_imagem('proce_apuracao_dominioweb.png', 5, 0.5)
    py.hotkey("ctrl", "c")
    data_obtida = pyperclip.paste()
    sleep(1)
    data_obtidaStr = str(data_obtida)

    data_final = f"{mes_ant}/{ano}"

    if data_obtidaStr != data_final:
        py.write(f"{mes_ant}{ano}")  # define período inicial de apuração
        sleep(2)
        py.press("tab")
        sleep(2)
        py.write(f"{mes_ant}{ano}")  # define período final de apuração
        automation.clicar_na_imagem("proce_apuracao_dominioweb.png", 60, 0.5)
        sleep(2)

    automation.esperar_imagem_sumir('processo_apuracao.png',180, 0.5)
    if automation.esperar_imagem('aviso_apuracao.png', 5, 0.5):
        py.press('esc', presses=3)
        sleep(1)

    while not automation.esperar_imagem("apuracao_selecao_periodo.png", 5, 0.5):
        py.press("n", presses=2)
        sleep(0.5)
    py.press("n", presses=2)
    sleep(0.5)
    py.press("esc", presses=3, interval=0.5)

    if not checar_dominio(5):
        print("Erro ao abrir o dominio", "error")
        raise Exception("Não conseguiu abrir o dominio")
    else:
        print("Dominio foi aberto corretamente")
