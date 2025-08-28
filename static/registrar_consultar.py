import datetime
from static.postgresql import Conexao_postgresql
from psycopg2.extensions import AsIs
import os


class registers:
    usuario_datamart = '5511670'
    senha_datamart = 'Aequatorial85408@'
    nome_do_banco_de_dados_postgresql = 'datamart'
    ip_conexao_postgresql = '10.6.2.211'

    usuario = os.getlogin()

    def __init__(self) -> None:
        self.data_atual = datetime.date.today()

        self.conexao = Conexao_postgresql(self.__class__.ip_conexao_postgresql,
                                          self.__class__.nome_do_banco_de_dados_postgresql,
                                          self.__class__.usuario_datamart, self.__class__.senha_datamart)

    def registro_sucesso(self, dicionario: dict, tabela: str) -> int:
        sql_insert = f'insert into {tabela} (%s) values %s'

        colunas_string = dicionario.keys()
        colunas_string = ','.join(colunas_string)

        return self.conexao.manipular(sql_insert, (AsIs(colunas_string), tuple(dicionario.values())))

    def registro_sucesso_list(self, tabela, colunas_string, linha) -> None:
        sql_insert = f'insert into {tabela} (%s) values %s'
        self.conexao.manipular(sql_insert, (AsIs(colunas_string), tuple(linha)))

    def atualizar_registro(self, dicionario: dict, tabela: str, id) -> None:
        sql_update = " update {} set {} where id = {} "

        string_set = dicionario.keys()
        lista = ['{}={}'.format(coluna, '%s') for coluna in string_set]
        string_set = ','.join(lista)
        sql_1 = sql_update.format(tabela, string_set, id)

        self.conexao.manipular(sql_1, tuple(dicionario.values()))

    def consultar_notas(self, sql):
        lista = self.conexao.consultar(sql)
        return lista

    def procurar_historico_execucao(self, nome_do_relatorio):
        global qtd_tentativa
        data_atual = '{:04d}-{:02d}-{:02d} 00:00:00'.format(self.data_atual.year, self.data_atual.month,
                                                            self.data_atual.day)
        lista_consulta_datamart = self.conexao.consultar(
            f""" select tentativa, concluido from hist_bases where nome_do_relatorio = '{nome_do_relatorio}' and carimbo_tempo > '{data_atual}' """)
        if not lista_consulta_datamart:
            self.qtd_tentativa = 0
            return True
        else:
            for linha in lista_consulta_datamart:
                concluido = linha['concluido']
                qtd_tentativa = linha['tentativa']
                if concluido:
                    return False
            else:
                qtd_tentativa = int(qtd_tentativa)
                if qtd_tentativa > 3:
                    return False
                else:
                    self.qtd_tentativa = qtd_tentativa
                    return True

    def resetando_mes(self, tabela, mes, ano):
        self.conexao.query(f'delete from {tabela} where ano = {ano} and mes = {mes}')
