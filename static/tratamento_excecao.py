"""
Módulo: tratamento_excecao.py
Descrição: Captura e formata informações detalhadas de exceções Python
Autor: Automação e Inovação - Contact Center
Uso: Chamado em blocos except para registrar erros no banco de dados
"""

import sys, traceback

def tratamento_excecao() -> str:
    """
    Captura detalhes completos de uma exceção que ocorreu.
    
    Retorna:
        str: String formatada com todos os detalhes do erro
        Formato: "chave:valor|chave:valor|..."
        
    Exemplo de retorno:
        "arquivo:main.py|linha:45|nome:extracao_site_optimus|tipo_erro:ValueError|
         mensagem_erro:invalid literal|arquivo_next:['main.py','postgresql.py']|..."
    """
    
    # sys.exc_info() retorna informações sobre a exceção mais recente
    # exc_type: Tipo da exceção (ex: ValueError, KeyError)
    # exc_value: Mensagem da exceção
    # exc_traceback: Objeto traceback com a pilha de chamadas
    exc_type, exc_value, exc_traceback = sys.exc_info()
    
    # traceback.extract_tb() converte o traceback em uma lista de objetos FrameSummary
    # Cada FrameSummary contém: filename, lineno, name, line
    lista_traceback = traceback.extract_tb(exc_traceback)
    
    # Listas para armazenar informações de TODOS os níveis da pilha de erros
    lista_arquivo_next = []
    lista_linha_next = []
    lista_nome_next = []
    
    # Itera sobre cada frame da pilha de execução
    for contagem, dicionario in enumerate(lista_traceback):
        
        # contagem == 0: Captura informações do PRIMEIRO erro (origem do problema)
        if contagem == 0:
            arquivo = dicionario.filename        # Caminho completo do arquivo
            linha = dicionario.lineno            # Número da linha onde ocorreu o erro
            nome = dicionario.name               # Nome da função onde ocorreu
            tipo_erro = str(exc_type.__name__)   # Nome da classe da exceção (ex: "ValueError")
            mensagem_erro = str(exc_value)       # Mensagem descritiva do erro
            
            # Extrai apenas o nome do arquivo (remove o caminho completo)
            # Exemplo: "C:\\Users\\user\\projeto\\main.py" -> "main.py"
            arquivo = str(arquivo).split('\\')
            arquivo = arquivo[len(arquivo) - 1]
            
            # Dicionário com detalhes do erro INICIAL
            detalhes_do_erro =  {
                                'arquivo' : arquivo,           # Nome do arquivo
                                'linha'  : linha,              # Linha do erro
                                'nome'    : nome,              # Nome da função
                                'tipo_erro'    : tipo_erro,    # Tipo da exceção
                                'mensagem_erro' : mensagem_erro, # Mensagem do erro
                                }
        
        # Para TODOS os frames (incluindo o primeiro), coleta a sequência completa
        arquivo_next = dicionario.filename
        arquivo_next = str(arquivo_next).split('\\')
        arquivo_next = arquivo_next[len(arquivo_next) - 1]
        
        # Adiciona informações de cada nível da pilha
        lista_arquivo_next.append(arquivo_next)     # Sequência de arquivos
        lista_linha_next.append(dicionario.lineno)  # Sequência de linhas
        lista_nome_next.append(dicionario.name)     # Sequência de funções
    
    # Dicionário com a SEQUÊNCIA COMPLETA da pilha de erros
    # Permite rastrear o caminho completo do erro através dos arquivos
    detalhes_do_erro_next =  {
                                    'arquivo_next' : lista_arquivo_next,
                                    'linha_next'  : lista_linha_next,
                                    'nome_next' : lista_nome_next,
                                    }
    
    # Operador | (união de dicionários - Python 3.9+)
    # Combina detalhes do erro inicial com a sequência completa
    detalhes_do_erro = detalhes_do_erro | detalhes_do_erro_next
    
    
    # Formata o dicionário como string no padrão: "chave:valor|chave:valor"
    colunas = detalhes_do_erro.keys()
    valores = ['{}:{}'.format(coluna, detalhes_do_erro[coluna]) for coluna in colunas]
    
    # Une tudo com pipe (|) como separador
    str_detalhes_do_erro = '|'.join(valores)
    
    return str_detalhes_do_erro
