"""
Módulo: extracao_site/main.py
Versão: 1.0.2
Descrição: Automatiza extração de fichas de monitoria do sistema Optimus usando Playwright
Autor: Automação e Inovação - Contact Center
Colaboradores: Ricardo Gomes (TATE5507392); Everton Barreto (U5511121)
Sistema: Optimus (10.6.1.160)
"""

import os
import datetime
import time
from playwright.sync_api import sync_playwright


class main:
    """
    Classe para automação de extração de fichas do sistema Optimus.
    
    Funcionalidades:
    1. Acessa sistema Optimus via web
    2. Faz login automatizado
    3. Navega até tela de exportação de monitorias
    4. Configura filtros (mês, tipos de ficha, operações)
    5. Baixa cada ficha como arquivo .xls
    
    Tecnologia: Playwright (automação de navegador)
    """
    
    # Lista de abreviações dos meses para navegação no site
    meses = ['jan', 'fev', 'mar', 'abr', 'mai', 'jun', 'jul', 'ago', 'set', 'out', 'nov', 'dez']
    
    # Credenciais de acesso ao Optimus
    usuario_optimus = '5510004'
    senha_optimus = 'SFJ@n26ina00@'

    def __init__(self, data_dia_anterior) -> None:
        """
        Inicializa o extrator.
        
        Args:
            data_dia_anterior (datetime): Data de referência para extração
            
        Atributos:
            tempo_espera: Timestamp para medição de tempo
            caminho_relativo: Diretório do script
            data_atual: Datetime agora
            dia_anterior: Data D-1 (referência para os dados)
        """
        self.tempo_espera = time.time()
        self.caminho_relativo = os.path.dirname(__file__)
        self.data_atual = datetime.datetime.now()
        self.dia_anterior = self.data_atual - datetime.timedelta(days=1)

    def configuracao_playwright(self, p):
        """
        Configura e inicializa o navegador Playwright.
        
        Args:
            p: Instância do Playwright
            
        Returns:
            page: Objeto de página do navegador
            
        Configurações:
        - headless=False: Navegador visível (útil para debug)
        - accept_downloads=True: Permite downloads automáticos
        - timeout=0: Sem timeout (espera indefinida)
        
        Nota:
            headless=False é importante para ambientes que não suportam
            modo headless ou quando precisa visualizar a execução
        """
        browser = p.chromium.launch(headless=False)
        
        # Cria contexto com permissão para downloads
        context1 = browser.new_context(accept_downloads=True)
        
        # Cria nova página
        page = context1.new_page()
        
        # Remove timeout padrão (espera indefinida)
        page.set_default_timeout(0)
        
        return page

    def primeira_pagina_optimus(self, page):
        """
        Acessa página inicial do Optimus e clica no botão de entrada.
        
        Args:
            page: Objeto de página do Playwright
            
        Returns:
            page1: Nova página que abre após clicar (popup)
            
        Comportamento:
        - Acessa http://10.6.1.160/
        - Aguarda popup abrir (janela de login)
        - Retorna a página do popup
        """
        page.goto("http://10.6.1.160/")
        
        # expect_popup captura a nova janela que vai abrir
        with page.expect_popup() as popup_info:
            page.locator("text=Entrar").click()
        
        # Obtém a página do popup
        page1 = popup_info.value
        return page1

    def login_optimus(self, page):
        """
        Realiza login no sistema Optimus.
        
        Args:
            page: Página de login
            
        Passos:
        1. Preenche campo de usuário
        2. Preenche campo de senha
        3. Clica no botão Entrar
        4. Aguarda navegação (carregamento da página principal)
        """
        page.locator("input[name=\"txtLogin\"]").fill(self.__class__.usuario_optimus)
        page.locator("input[name=\"txtSenha\"]").fill(self.__class__.senha_optimus)
        
        # expect_navigation aguarda o carregamento completo após o clique
        with page.expect_navigation():
            page.locator("text=Entrar").click()

    def menu_optimus(self, page):
        """
        Navega pelo menu do Optimus até a tela de exportação.
        
        Args:
            page: Página principal do Optimus
            
        Navegação:
        - Menu principal (id=1)
        - Submenu (id=16) - com retry em loop
        - Item (id=20)
        - Item final (id=853, segundo elemento)
        
        Nota:
            O loop while True com try/except é para aguardar o menu
            carregar completamente antes de clicar no próximo item
        """
        # Clica no menu principal
        while True:
            page.locator("a[id=\"1\"]").click()
            try:
                # Tenta clicar no submenu (timeout de 300ms)
                # Se conseguir, sai do loop
                page.locator("[id=\"\\31 6\"]").click(timeout=300)
                break
            except:
                # Se falhar, tenta novamente
                pass

        # Continua navegação no menu
        page.locator("[id=\"\\32 0\"]").click()
        
        # Pega o segundo elemento com id=853 (nth=1 é zero-indexed, então 2º elemento)
        page.locator("id=853 >> nth=1").click()

    def checagem_inicial_da_tela_exporta_monitorias(self, page):
        """
        Configura todos os filtros na tela de exportação de monitorias.
        
        Args:
            page: Página de exportação
            
        Configurações aplicadas:
        1. Seleciona mês anterior
        2. Marca checkboxes de tipos de ficha:
           - rdItem_121
           - rdItem_150
           - rdItem_85
           - rdItem_201
           - rdItem_166
        3. Marca "Todos" nos grupos de operações:
           - grupo_2
           - grupo_6
           - grupo_32
           - grupo_36 (este é o principal - lista de fichas)
        
        Nota sobre frames:
            Todos os seletores usam .frame_locator("iframe[name=\"navMain\"]")
            porque o conteúdo está dentro de um iframe
        """
        # Obtém abreviação do mês anterior
        mes = self.__class__.meses[self.dia_anterior.month - 1]
        
        # Se o mês é de ano anterior, clica na seta de voltar no calendário
        if self.dia_anterior.year != datetime.datetime.now().year:
            page.frame_locator("iframe[name=\"navMain\"]").locator("th[class=\"datepickerGoPrev\"]").click()
        
        # Marca checkboxes de tipos de ficha
        page.frame_locator("iframe[name=\"navMain\"]").locator("#rdItem_121").check()
        page.frame_locator("iframe[name=\"navMain\"]").locator("a:has-text(\"" + mes + "\")").click()
        page.frame_locator("iframe[name=\"navMain\"]").locator("#rdItem_150").check()
        page.frame_locator("iframe[name=\"navMain\"]").locator("#rdItem_121").check()
        page.frame_locator("iframe[name=\"navMain\"]").locator("#rdItem_85").check()
        page.frame_locator("iframe[name=\"navMain\"]").locator("#rdItem_201").check()
        page.frame_locator("iframe[name=\"navMain\"]").locator("#rdItem_166").check()
        
        # Checkbox comentado (desabilitado)
        # page.frame_locator("iframe[name=\"navMain\"]").locator("input[name=\"chkFiltarOpPrincipal\"]").check()
        
        # Marca "Todos" nos grupos de operações
        # Cada .wait_for() garante que o select carregou antes de marcar
        
        page.frame_locator("iframe[name=\"navMain\"]").locator("select[name=\"grupo_2\"]").wait_for()
        page.frame_locator("iframe[name=\"navMain\"]").locator("input[name=\"chkTodos2\"]").check()
        
        page.frame_locator("iframe[name=\"navMain\"]").locator("select[name=\"grupo_6\"]").wait_for()
        page.frame_locator("iframe[name=\"navMain\"]").locator("input[name=\"chkTodos6\"]").check()
        
        # Grupo 15 comentado (desabilitado)
        # page.frame_locator("iframe[name=\"navMain\"]").locator("select[name=\"grupo_15\"]").wait_for()
        # page.frame_locator("iframe[name=\"navMain\"]").locator("input[name=\"chkTodos15\"]").check()
        
        page.frame_locator("iframe[name=\"navMain\"]").locator("select[name=\"grupo_32\"]").wait_for()
        page.frame_locator("iframe[name=\"navMain\"]").locator("input[name=\"chkTodos32\"]").check()
        
        # Aguarda o select principal (grupo_36) carregar
        page.frame_locator("iframe[name=\"navMain\"]").locator("select[name=\"grupo_36\"]").wait_for()

    def obtendo_lista_de_fichas(self, select_element) -> list:
        """
        Extrai lista de fichas disponíveis do elemento select.
        
        Args:
            select_element: Elemento <select> do Playwright
            
        Returns:
            list: Lista de nomes de fichas (excluindo INATIVAS)
            
        Processamento:
        1. Obtém todo o texto interno do select
        2. Separa por quebras de linha
        3. Remove caracteres não alfanuméricos no início/fim
        4. Filtra fichas que contêm "INATIVA"
        
        Lógica de limpeza:
            Playwright retorna o texto com caracteres especiais e espaços,
            então precisa limpar o início e fim de cada opção
        """
        # Obtém todo texto interno do select
        option_elements = select_element.all_inner_texts()
        option_elements = str(option_elements)
        
        # Separa por quebras de linha
        option_elements = option_elements.split('\\n')
        
        # Remove caracteres não alfanuméricos do INÍCIO da primeira linha
        while ''.join(filter(str.isalnum, option_elements[0][0])) == '':
            option_elements[0] = option_elements[0][1:]

        # Remove caracteres não alfanuméricos do FINAL da última linha
        while ''.join(filter(str.isalnum, option_elements[len(option_elements) - 1][
            len(option_elements[len(option_elements) - 1]) - 1])) == '':
            option_elements[len(option_elements) - 1] = option_elements[len(option_elements) - 1][
                                                        :len(option_elements[len(option_elements) - 1]) - 1]

        # Filtra apenas fichas ativas (remove as que contêm "INATIVA")
        lista_fichas = [str(linha) for linha in option_elements if not 'INATIVA'.upper() in linha.upper()]

        return lista_fichas

    def extracao_site_optimus(self):
        """
        Método principal que executa todo o fluxo de extração.
        
        Fluxo completo:
        1. Inicia Playwright
        2. Configura navegador
        3. Acessa Optimus
        4. Faz login
        5. Navega até tela de exportação
        6. Configura filtros
        7. Obtém lista de fichas
        8. Para cada ficha:
           a. Seleciona a ficha
           b. Clica em "Selecionar"
           c. Captura download
           d. Salva com nome padronizado
        
        Formato do arquivo:
            "MM-AAAA Nome da Ficha.xls"
            Exemplo: "01-2025 FICHA RECEPTIVO.xls"
        
        Diretório destino:
            C:\\Users\\sautomaeqtls\\Documents\\Docs monitorias
        """
        print("🔹 Iniciando extração...")
        
        with sync_playwright() as p:
            print("🔹 Configurando navegador...")
            page = self.configuracao_playwright(p=p)
            
            print("🔹 Acessando Optimus...")
            page = self.primeira_pagina_optimus(page=page)
            
            print("🔹 Fazendo login...")
            self.login_optimus(page=page)
            
            print("🔹 Navegando menu...")
            self.menu_optimus(page=page)
            
            print("🔹 Configurando filtros...")
            self.checagem_inicial_da_tela_exporta_monitorias(page=page)
            
            print("🔹 Obtendo lista de fichas")
            select_element = page.frame_locator("iframe[name=\"navMain\"]").locator("select[name=\"grupo_36\"]")
            lista_de_fichas = self.obtendo_lista_de_fichas(select_element=select_element)
            
            print(f"🔹 Encontradas {len(lista_de_fichas)} fichas:")
            print(lista_de_fichas)

            # Loop principal: processa cada ficha
            for ficha in lista_de_fichas:
                print(f"🔹 Processando: {ficha}")
                time.sleep(3)  # Pausa para estabilizar

                # Seleciona a ficha no dropdown
                page.frame_locator("iframe[name=\"navMain\"]").locator("select[name=\"grupo_36\"]").select_option(ficha)

                # Prepara para fechar possível dialog/alert
                page.once("dialog", lambda dialog: dialog.dismiss())
                
                # Clica no botão "Selecionar" e aguarda popup
                with page.expect_popup() as popup_info:
                    page.frame_locator("iframe[name=\"navMain\"]").locator("img[alt=\"Selecionar\"]").nth(1).click()

                time.sleep(3)

                # Obtém página do popup
                page_1 = popup_info.value
                page_1.set_default_timeout(0)

                try:
                    # Aguarda download iniciar
                    with page_1.expect_download() as download_info:
                        page_1.wait_for_event("download")

                    download = download_info.value
                    
                    # Monta nome do arquivo: "MM-AAAA Nome da Ficha.xls"
                    file_name = '{:02d}-{:04d}'.format(self.dia_anterior.month,
                                                       self.dia_anterior.year) + ' ' + ficha + '.xls'
                    
                    destination_folder_path = r"\\EQTSPDSRCL01\planejamento_e_trafego\Automacao e Inovacao\Fichas_monitorias"
                    time.sleep(3)

                    # Salva arquivo (remove barras do nome se houver)
                    download.save_as(os.path.join(destination_folder_path, str(file_name).replace('/', '')))
                    print(f" Arquivo salvo: {file_name}")
                    
                except Exception as e:
                    print(f" Erro ao baixar {ficha}: {str(e)}")
                    
                # Fecha popup
                page_1.close()
            
            print("✅ Extração concluída!")
