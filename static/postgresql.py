"""
Módulo: postgresql.py
Descrição: Classe para gerenciar conexões e operações no PostgreSQL
Autor: Automação e Inovação - Contact Center
Recursos: Reconexão automática, proteção contra SQL injection, retorno de dicionários
"""

from typing import Any

import psycopg2.extras
from psycopg2 import connect


class Conexao_postgresql(object):
    """
    Classe de conexão PostgreSQL com gerenciamento automático de reconexão.
    
    Características:
    - Reconexão automática em caso de perda de conexão
    - Uso de cursores parametrizados (proteção contra SQL injection)
    - Retorno de consultas como lista de dicionários
    - Commit automático após cada operação
    """
    
    def __init__(self, mhost, db, usr, pwd):
        """
        Inicializa a conexão com o PostgreSQL.
        
        Args:
            mhost (str): Host/IP do servidor PostgreSQL
            db (str): Nome do banco de dados
            usr (str): Usuário do banco
            pwd (str): Senha do usuário
        """
        # Armazena credenciais para reconexão automática
        self.pwd = pwd
        self.usr = usr
        self.db = db
        self.mhost = mhost
        
        # Estabelece a conexão inicial
        self._db = connect(host=mhost, database=db, user=usr, password=pwd)

    def manipular(self, sql, _Vars):
        """
        Executa comandos SQL de manipulação (INSERT, UPDATE, DELETE).
        
        Args:
            sql (str): Comando SQL com placeholders (%s)
            _Vars (tuple): Valores para os placeholders
            
        Comportamento:
        - INSERT: Verifica se é insert nos primeiros 10 caracteres
        - Outros comandos: UPDATE, DELETE, etc.
        - Commit automático após execução
        - Reconexão automática se conexão perdida
        
        Raises:
            AssertionError: Se houver erro na execução e conexão estiver ativa
            
        Nota:
            Código comentado mostra possibilidade de retornar ID do registro inserido
        """
        try:
            # Detecta se é um INSERT (verifica primeiros 10 caracteres)
            if 'insert' in sql[:10]:
                # Código comentado: retornaria o ID do registro inserido
                # if 'RETURNING id' not in sql:
                #     sql = sql + ' RETURNING id '
                cur = self._db.cursor()
                cur.execute(sql, _Vars)  # Execute com proteção contra SQL injection
                # fk = cur.fetchone()[0]  # Pegaria o ID retornado
                cur.close()
                self._db.commit()  # Confirma a transação
                # return fk
            else:
                # Para UPDATE, DELETE e outros comandos
                cur = self._db.cursor()
                cur.execute(sql, _Vars)
                cur.close()
                self._db.commit()
                
        except Exception as e:
            # self._db.closed retorna:
            # 0 = conexão aberta
            # != 0 = conexão fechada/perdida
            if self._db.closed != 0:
                # Reconecta automaticamente
                self._db = connect(host=self.mhost, database=self.db, user=self.usr, password=self.pwd)
                # Tenta executar novamente (recursão)
                return self.manipular(sql=sql, _Vars=_Vars)
            else:
                # Se conexão está ativa mas deu erro, é erro real
                e = str(e)
                try:
                    cur.close()  # Tenta fechar cursor se existir
                except:
                    pass
                raise AssertionError(e)  # Propaga o erro

    def query(self, sql):
        """
        Executa comandos SQL diretos sem parâmetros (menos seguro).
        
        Args:
            sql (str): Comando SQL completo
            
        Uso:
            Ideal para comandos DDL (CREATE, DROP, ALTER, TRUNCATE)
            ou comandos sem dados externos
            
        Exemplo:
            conn.query("DELETE FROM tabela WHERE mes = 12")
            conn.query("TRUNCATE TABLE temp_table")
            
        Aviso:
            Não use com dados do usuário (risco de SQL injection)
        """
        try:
            cur = self._db.cursor()
            cur.execute(sql)  # Executa SQL direto
            cur.close()
            self._db.commit()
            
        except Exception as e:
            # Mesma lógica de reconexão do método manipular()
            if self._db.closed != 0:
                self._db = connect(host=self.mhost, database=self.db, user=self.usr, password=self.pwd)
                return self.query(sql=sql)
            else:
                e = str(e)
                try:
                    cur.close()
                except:
                    pass
                raise AssertionError(e)

    def consultar(self, sql) -> list[dict[Any, Any]]:
        """
        Executa SELECT e retorna resultados como lista de dicionários.
        
        Args:
            sql (str): Query SELECT
            
        Returns:
            list[dict]: Lista de dicionários onde cada dict é uma linha
            
        Exemplo de retorno:
            [
                {'id': 1, 'nome': 'João', 'idade': 30},
                {'id': 2, 'nome': 'Maria', 'idade': 25}
            ]
            
        Vantagens do DictCursor:
        - Acesso por nome de coluna: row['nome']
        - Mais legível que índices numéricos
        - Fácil de converter para JSON
        - Evita erros ao mudar ordem das colunas
        """
        try:
            rs = None
            # DictCursor: retorna linhas como dicionários ao invés de tuplas
            cur = self._db.cursor(cursor_factory=psycopg2.extras.DictCursor)
            cur.execute(sql)
            rs = cur.fetchall()  # Busca todos os resultados
            
            # Converte cada linha (DictRow) para dict Python puro
            ans = []
            for row in rs:
                ans.append(dict(row))
            cur.close()
            return ans
            
        except Exception as e:
            # Mesma lógica de reconexão
            if self._db.closed != 0:
                self._db = connect(host=self.mhost, database=self.db, user=self.usr, password=self.pwd)
                return self.consultar(sql=sql)
            else:
                e = str(e)
                try:
                    cur.close()
                except:
                    pass
                raise AssertionError(e)

    def fechar(self):
        """
        Fecha a conexão com o banco de dados.
        
        Uso:
            Deve ser chamado ao final do script ou quando a conexão
            não é mais necessária para liberar recursos.
            
        Nota:
            Na maioria dos casos do projeto, a conexão fica aberta
            durante toda a execução do RPA.
        """
        self._db.close()
