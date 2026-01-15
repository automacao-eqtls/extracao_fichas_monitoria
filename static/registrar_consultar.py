"""
Módulo: registrar_consultar.py
Descrição: Camada de abstração para operações no banco de dados do Datamart
Autor: Automação e Inovação - Contact Center
Recursos: Inserção dinâmica, controle de execução, gerenciamento de histórico
"""

import datetime
from static.postgresql import Conexao_postgresql
from psycopg2.extensions import AsIs  # Permite inserir SQL literal de forma segura
import os


class registers:
    """
    Classe para gerenciar registros e consultas no banco Datamart.
    
    Funcionalidades:
    - Inserção dinâmica a partir de dicionários
    - Controle de histórico de execução (evita reprocessamento)
    - Atualização de registros
    - Limpeza de dados mensais
    """
    
    # Credenciais do banco Datamart (variáveis de classe - compartilhadas)
    usuario_datamart = '5511670'
    senha_datamart = 'Aequatorial85408@'
    nome_do_banco_de_dados_postgresql = 'datamart'
    ip_conexao_postgresql = '10.6.2.211'

    # Captura o usuário do Windows logado no momento
    usuario = os.getlogin()

    def __init__(self) -> None:
        """
        Inicializa a classe e estabelece conexão com o banco.
        
        Atributos criados:
            data_atual: Data de hoje (para controle de execução)
            conexao: Objeto de conexão PostgreSQL
        """
        self.data_atual = datetime.date.today()

        # Cria instância da conexão usando as credenciais da classe
        self.conexao = Conexao_postgresql(self.__class__.ip_conexao_postgresql,
                                          self.__class__.nome_do_banco_de_dados_postgresql,
                                          self.__class__.usuario_datamart, self.__class__.senha_datamart)

    def registro_sucesso(self, dicionario: dict, tabela: str) -> int:
        """
        Insere um registro no banco a partir de um dicionário.
        
        Args:
            dicionario (dict): Dados a inserir {'coluna': valor}
            tabela (str): Nome da tabela (ex: 'monitoria.fichas')
            
        Returns:
            int: ID do registro inserido (se configurado RETURNING)
            
        Exemplo:
            registro_sucesso(
                dicionario={'nome': 'João', 'idade': 30},
                tabela='public.usuarios'
            )
            
            Gera SQL: INSERT INTO public.usuarios (nome, idade) VALUES ('João', 30)
            
        Vantagens:
        - Dinâmico: Não precisa escrever SQL manualmente
        - Seguro: Usa parametrização
        - Limpo: Código mais legível
        """
        # Template SQL com placeholders
        sql_insert = f'insert into {tabela} (%s) values %s'

        # Extrai as chaves do dicionário (nomes das colunas)
        colunas_string = dicionario.keys()
        colunas_string = ','.join(colunas_string)  # 'nome,idade,cidade'

        # AsIs: Insere as colunas como SQL literal (não como string)
        # tuple(dicionario.values()): Converte valores em tupla
        return self.conexao.manipular(sql_insert, (AsIs(colunas_string), tuple(dicionario.values())))

    def registro_sucesso_list(self, tabela, colunas_string, linha) -> None:
        """
        Insere registro usando lista de valores (mais performático em loops).
        
        Args:
            tabela (str): Nome da tabela
            colunas_string (str): String com colunas separadas por vírgula
            linha (list): Lista de valores na mesma ordem das colunas
            
        Uso:
            Otimizado para inserção em massa dentro de loops
            Evita recriar a string de colunas a cada iteração
            
        Exemplo:
            colunas = 'nome,idade,cidade'
            for linha in dados:
                registro_sucesso_list('usuarios', colunas, linha)
        """
        sql_insert = f'insert into {tabela} (%s) values %s'
        self.conexao.manipular(sql_insert, (AsIs(colunas_string), tuple(linha)))

    def atualizar_registro(self, dicionario: dict, tabela: str, id) -> None:
        """
        Atualiza um registro existente pelo ID.
        
        Args:
            dicionario (dict): Colunas e novos valores {'coluna': novo_valor}
            tabela (str): Nome da tabela
            id (int): ID do registro a atualizar
            
        Exemplo:
            atualizar_registro(
                dicionario={'nome': 'João Silva', 'idade': 31},
                tabela='public.usuarios',
                id=5
            )
            
            Gera SQL: UPDATE public.usuarios SET nome=%s, idade=%s WHERE id = 5
        """
        sql_update = " update {} set {} where id = {} "

        # Extrai colunas do dicionário
        string_set = dicionario.keys()
        
        # Cria lista: ['nome=%s', 'idade=%s']
        lista = ['{}={}'.format(coluna, '%s') for coluna in string_set]
        
        # Junta com vírgula: 'nome=%s,idade=%s'
        string_set = ','.join(lista)
        
        # Formata SQL completo
        sql_1 = sql_update.format(tabela, string_set, id)

        # Executa com valores do dicionário
        self.conexao.manipular(sql_1, tuple(dicionario.values()))

    def consultar_notas(self, sql):
        """
        Executa consulta SELECT e retorna resultados.
        
        Args:
            sql (str): Query SQL completa
            
        Returns:
            list[dict]: Lista de dicionários com resultados
            
        Nota:
            Nome 'consultar_notas' é legado, mas serve para qualquer SELECT
        """
        lista = self.conexao.consultar(sql)
        return lista

    def procurar_historico_execucao(self, nome_do_relatorio):
        """
        Verifica se o RPA já foi executado hoje com sucesso.
        
        Args:
            nome_do_relatorio (str): Identificador único do RPA
            
        Returns:
            bool: True = pode executar, False = já executado ou limite de tentativas
            
        Lógica:
        1. Consulta hist_bases buscando execuções de hoje
        2. Se não há registros → Primeira execução do dia → Retorna True
        3. Se há registro concluído → Já executou com sucesso → Retorna False
        4. Se há falhas mas < 3 tentativas → Pode tentar novamente → Retorna True
        5. Se >= 3 tentativas falhadas → Limite atingido → Retorna False
        
        Efeito colateral:
            Define self.qtd_tentativa com o número de tentativas atuais
            
        Uso:
            Evita reprocessamento e controla retry automático
        """
        global qtd_tentativa  # Aviso: uso de global (não recomendado)
        
        # Formata data de hoje no formato PostgreSQL
        data_atual = '{:04d}-{:02d}-{:02d} 00:00:00'.format(self.data_atual.year, self.data_atual.month,
                                                            self.data_atual.day)
        
        # Busca execuções deste relatório a partir de hoje 00:00
        lista_consulta_datamart = self.conexao.consultar(
            f""" select tentativa, concluido from hist_bases where nome_do_relatorio = '{nome_do_relatorio}' and carimbo_tempo > '{data_atual}' """)
        
        # Caso 1: Sem registros de hoje → Primeira execução
        if not lista_consulta_datamart:
            self.qtd_tentativa = 0
            return True
        else:
            # Caso 2 e 3: Há registros de hoje
            for linha in lista_consulta_datamart:
                concluido = linha['concluido']
                qtd_tentativa = linha['tentativa']
                
                # Se encontrou registro concluído → Já executou com sucesso
                if concluido:
                    return False
            else:
                # Se chegou aqui: nenhum registro está concluído (todas tentativas falharam)
                qtd_tentativa = int(qtd_tentativa)
                
                # Limite de 3 tentativas
                if qtd_tentativa > 20:
                    return False  # Desiste após 3 falhas
                else:
                    # Pode tentar novamente
                    self.qtd_tentativa = qtd_tentativa
                    return True

    def resetando_mes(self, tabela, mes, ano):
        """
        Remove todos os dados de um mês/ano específico de uma tabela.
        
        Args:
            tabela (str): Nome da tabela
            mes (int): Mês (1-12)
            ano (int): Ano (ex: 2025)
            
        Uso:
            Chamado antes de reprocessar dados de um mês
            Garante que não haverá duplicatas
            
        Exemplo:
            resetando_mes('public.fichas_monitoria', 1, 2025)
            # Remove todos os registros de janeiro/2025
            
        Atenção:
            Operação destrutiva! Não há como reverter.
        """
        self.conexao.query(f'delete from {tabela} where ano = {ano} and mes = {mes}')
