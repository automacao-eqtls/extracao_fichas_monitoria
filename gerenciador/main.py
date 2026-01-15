"""
Módulo: gerenciador/main.py
Descrição: Orquestrador principal que executa todo o fluxo do RPA
Autor: Automação e Inovação - Contact Center
Função: Coordena extração → importação com controle de erros e histórico
"""

import sys
import os

# Configura paths para importar módulos do projeto
caminho_relativo = os.path.dirname(__file__)
caminho_relativo = os.path.dirname(caminho_relativo)
sys.path.append(caminho_relativo)

# Ordem de diretórios a adicionar ao path
ordem_das_etapas = ['static', 'insercao_datamart', 'extracao_site', 'venv']
for etapa in ordem_das_etapas:
    sys.path.append(os.path.join(caminho_relativo, etapa))

try:
    # Imports dos módulos principais
    from extracao_site.main import main as site
    from insercao_datamart.main import main as insercao
    from static.tratamento_excecao import tratamento_excecao
    from static.registrar_consultar import registers
    import time
    import datetime

    # ETAPA 1: EXTRAÇÃO DO SITE
    
    tabela = 'public.hist_bases'
    nome_do_relatorio = 'optimus_monitoria_site'

    # Dicionário base para registro de histórico
    dicionario = {
        'carimbo_tempo': datetime.datetime.now(),
        'nome_do_relatorio': nome_do_relatorio,
        'tempo_de_extracao_seg': time.time()  # Marca início
    }

    # Data de referência: D-1
    data_dia_anterior = dicionario['carimbo_tempo'] - datetime.timedelta(days=1)

    # Instancia o extrator do site
    inst_main_extracao = site(data_dia_anterior)

    # Instancia o registrador
    inst_registers = registers()

    # Verifica se pode executar (não executou hoje ou tem tentativas disponíveis)
    if inst_registers.procurar_historico_execucao(nome_do_relatorio=nome_do_relatorio):
        dicionario['tentativa'] = inst_registers.qtd_tentativa
        
        try:
            # EXECUÇÃO: Extrai fichas do site Optimus
            # NOTA: Linha comentada no código original (não executa)
            inst_main_extracao.extracao_site_optimus()
            
            # Calcula tempo de execução
            dicionario['tempo_de_extracao_seg'] = time.time() - dicionario['tempo_de_extracao_seg']
            
            # Marca como concluído
            dicionario['concluido'] = True
            
            # Registra sucesso no histórico
            inst_registers.registro_sucesso(dicionario=dicionario, tabela=tabela)
            
        except:
            # Captura erro detalhado
            str_erro = tratamento_excecao()
            
            # Calcula tempo até o erro
            dicionario['tempo_de_extracao_seg'] = time.time() - dicionario['tempo_de_extracao_seg']
            
            # Adiciona informações do erro
            dicionario['msg_erro'] = str_erro
            dicionario['tentativa'] = inst_registers.qtd_tentativa + 1
            
            # Registra falha no histórico
            inst_registers.registro_sucesso(dicionario=dicionario, tabela=tabela)

    # ETAPA 2: IMPORTAÇÃO PARA DATAMART
    
    tabela = 'public.hist_bases'
    nome_do_relatorio = 'fichas_importacao'

    # Novo dicionário para segunda etapa
    dicionario = {
        'carimbo_tempo': datetime.datetime.now(),
        'nome_do_relatorio': nome_do_relatorio,
        'tempo_de_extracao_seg': time.time()
    }

    # Instancia o importador
    inst_main_extracao = insercao(data_extracao=data_dia_anterior)

    # Nova instância do registrador
    inst_registers = registers()

    # Verifica se pode executar
    if inst_registers.procurar_historico_execucao(nome_do_relatorio=nome_do_relatorio):
        dicionario['tentativa'] = inst_registers.qtd_tentativa
        
        try:
            # EXECUÇÃO: Processa fichas e insere no Datamart
            inst_main_extracao.run()
            
            # Calcula tempo de execução
            dicionario['tempo_de_extracao_seg'] = time.time() - dicionario['tempo_de_extracao_seg']
            
            # Marca como concluído
            dicionario['concluido'] = True
            
            # Registra sucesso
            inst_registers.registro_sucesso(dicionario=dicionario, tabela=tabela)
            
        except:
            # Captura erro
            str_erro = tratamento_excecao()
            
            # Calcula tempo até erro
            dicionario['tempo_de_extracao_seg'] = time.time() - dicionario['tempo_de_extracao_seg']
            
            # Adiciona informações do erro
            dicionario['msg_erro'] = str_erro
            dicionario['tentativa'] = inst_registers.qtd_tentativa + 1
            
            # Registra falha
            inst_registers.registro_sucesso(dicionario=dicionario, tabela=tabela)

except Exception as excecao:
    # Captura exceções não tratadas (erros nos imports, etc)
    print(excecao)