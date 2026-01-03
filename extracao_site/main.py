# versão 1.0.2
# Criado por automação e inovação - Contact Center
# Colaboraram: Ricardo Gomes (TATE5507392); Everton Barreto (U5511121);

import os
import datetime
import time
from playwright.sync_api import sync_playwright


class main:
    meses = ['jan', 'fev', 'mar', 'abr', 'mai', 'jun', 'jul', 'ago', 'set', 'out', 'nov', 'dez']
    usuario_optimus = '5510004'
    senha_optimus = 'SFJ@n26ina00@'

    def __init__(self, data_dia_anterior) -> None:
        self.tempo_espera = time.time()
        self.caminho_relativo = os.path.dirname(__file__)
        self.data_atual = datetime.datetime.now()
        self.dia_anterior = self.data_atual - datetime.timedelta(days=1)

    def configuracao_playwright(self, p):
        caminho_driver = r"C:\Users\U5511670\AppData\Local\ms-playwright\chromium-1187\chrome-win\chrome.exe"
        browser = p.chromium.launch(headless=False, executable_path=caminho_driver)
        context1 = browser.new_context(accept_downloads=True)
        page = context1.new_page()
        page.set_default_timeout(0)
        return page

    def primeira_pagina_optimus(self, page):
        page.goto("http://10.6.1.160/")
        with page.expect_popup() as popup_info:
            page.locator("text=Entrar").click()
        page1 = popup_info.value
        return page1

    def login_optimus(self, page):
        page.locator("input[name=\"txtLogin\"]").fill(self.__class__.usuario_optimus)
        page.locator("input[name=\"txtSenha\"]").fill(self.__class__.senha_optimus)
        with page.expect_navigation():
            page.locator("text=Entrar").click()

    def menu_optimus(self, page):
        while True:
            page.locator("a[id=\"1\"]").click()
            try:
                page.locator("[id=\"\\31 6\"]").click(timeout=300)
                break
            except:
                pass

        page.locator("[id=\"\\32 0\"]").click()
        page.locator("id=853 >> nth=1").click()

    def checagem_inicial_da_tela_exporta_monitorias(self, page):
        mes = self.__class__.meses[self.dia_anterior.month - 1]
        if self.dia_anterior.year != datetime.datetime.now().year:
            page.frame_locator("iframe[name=\"navMain\"]").locator("th[class=\"datepickerGoPrev\"]").click()
        page.frame_locator("iframe[name=\"navMain\"]").locator("#rdItem_121").check()
        page.frame_locator("iframe[name=\"navMain\"]").locator("a:has-text(\"" + mes + "\")").click()
        page.frame_locator("iframe[name=\"navMain\"]").locator("#rdItem_150").check()
        page.frame_locator("iframe[name=\"navMain\"]").locator("#rdItem_121").check()
        page.frame_locator("iframe[name=\"navMain\"]").locator("#rdItem_85").check()
        page.frame_locator("iframe[name=\"navMain\"]").locator("#rdItem_201").check()
        page.frame_locator("iframe[name=\"navMain\"]").locator("#rdItem_166").check()
        # page.frame_locator("iframe[name=\"navMain\"]").locator("input[name=\"chkFiltarOpPrincipal\"]").check()
        page.frame_locator("iframe[name=\"navMain\"]").locator("select[name=\"grupo_2\"]").wait_for()
        page.frame_locator("iframe[name=\"navMain\"]").locator("input[name=\"chkTodos2\"]").check()
        page.frame_locator("iframe[name=\"navMain\"]").locator("select[name=\"grupo_6\"]").wait_for()
        page.frame_locator("iframe[name=\"navMain\"]").locator("input[name=\"chkTodos6\"]").check()
        # page.frame_locator("iframe[name=\"navMain\"]").locator("select[name=\"grupo_15\"]").wait_for()
        # page.frame_locator("iframe[name=\"navMain\"]").locator("input[name=\"chkTodos15\"]").check()
        page.frame_locator("iframe[name=\"navMain\"]").locator("select[name=\"grupo_32\"]").wait_for()
        page.frame_locator("iframe[name=\"navMain\"]").locator("input[name=\"chkTodos32\"]").check()
        page.frame_locator("iframe[name=\"navMain\"]").locator("select[name=\"grupo_36\"]").wait_for()

    def obtendo_lista_de_fichas(self, select_element) -> list:
        option_elements = select_element.all_inner_texts()
        option_elements = str(option_elements)
        option_elements = option_elements.split('\\n')
        while ''.join(filter(str.isalnum, option_elements[0][0])) == '':
            option_elements[0] = option_elements[0][1:]

        while ''.join(filter(str.isalnum, option_elements[len(option_elements) - 1][
            len(option_elements[len(option_elements) - 1]) - 1])) == '':
            option_elements[len(option_elements) - 1] = option_elements[len(option_elements) - 1][
                                                        :len(option_elements[len(option_elements) - 1]) - 1]

        lista_fichas = [str(linha) for linha in option_elements if not 'INATIVA'.upper() in linha.upper()]

        return lista_fichas

    def extracao_site_optimus(self):
        with sync_playwright() as p:

            page = self.configuracao_playwright(p=p)

            page = self.primeira_pagina_optimus(page=page)

            self.login_optimus(page=page)

            self.menu_optimus(page=page)

            self.checagem_inicial_da_tela_exporta_monitorias(page=page)

            select_element = page.frame_locator("iframe[name=\"navMain\"]").locator("select[name=\"grupo_36\"]")
            lista_de_fichas = self.obtendo_lista_de_fichas(select_element=select_element)

            for ficha in lista_de_fichas:
                time.sleep(3)

                page.frame_locator("iframe[name=\"navMain\"]").locator("select[name=\"grupo_36\"]").select_option(ficha)

                page.once("dialog", lambda dialog: dialog.dismiss())
                with page.expect_popup() as popup_info:
                    page.frame_locator("iframe[name=\"navMain\"]").locator("img[alt=\"Selecionar\"]").nth(1).click()

                time.sleep(3)

                page_1 = popup_info.value
                page_1.set_default_timeout(0)

                try:
                    with page_1.expect_download() as download_info:
                        page_1.wait_for_event("download")

                    download = download_info.value
                    file_name = '{:02d}-{:04d}'.format(self.dia_anterior.month,
                                                       self.dia_anterior.year) + ' ' + ficha + '.xls'
                    destination_folder_path = r"\\55ASPDCARQ01\55Atende\Administrativo\06 - Gerência Contact Center\03 - Call Center Sao Luis\02 - Monitoria de Qualidade\10.BASES"

                    time.sleep(3)

                    download.save_as(os.path.join(destination_folder_path, str(file_name).replace('/', '')))
                except:
                    pass
                page_1.close()
