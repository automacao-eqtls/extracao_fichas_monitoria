"""
Módulo: insercao_datamart/main.py
Descrição: Processa arquivos Excel de monitorias e insere no Datamart
Autor: Automação e Inovação - Contact Center
Fluxo: Leitura Excel → Tratamento → Padronização → Inserção no PostgreSQL
"""

import os
import sys
import pandas

# Adiciona o diretório pai ao path para permitir importar 'static'
# Necessário porque a estrutura tem pastas separadas (static, insercao_datamart, etc)
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from static.registrar_consultar import registers

class main:
    """
    Classe principal para processamento de fichas de monitoria.
    
    Responsabilidades:
    1. Ler arquivos Excel (.xls) de um diretório específico
    2. Identificar e extrair colunas relevantes
    3. Tratar tipos de dados (datas, números)
    4. Padronizar nomes de colunas
    5. Inserir no banco Datamart
    """
    
    # Caminho onde os arquivos Excel são salvos pela extração
    caminho_relativo = r'C:\Users\sautomaeqtls\Documents\Docs monitorias'
    
    def __init__(self, data_extracao) -> None:
        """
        Inicializa o processador de fichas.
        
        Args:
            data_extracao (datetime): Data de referência para buscar os arquivos
            
        Lógica:
        - Lista todos os arquivos .xls do diretório
        - Filtra apenas os do mês/ano da data_extracao
        - Padrão de nome: "MM-AAAA Nome da Ficha.xls"
        """
        self.data_atual = data_extracao
        
        # Lista todos os arquivos do diretório
        lista_de_arquivos = os.listdir(self.__class__.caminho_relativo)
        
        # Filtra apenas .xls que começam com o mês/ano correto
        # Exemplo: "01-2025 FICHA RECEPTIVO.xls"
        self.lista_de_arquivos = [
            os.path.join(self.__class__.caminho_relativo, linha) 
            for linha in lista_de_arquivos 
            if linha.endswith('.xls') and linha.startswith('{:02d}-{:04d}'.format(self.data_atual.month, self.data_atual.year))
        ]
        
        self.conexao = registers()  # Instância para comunicação com banco
        self.colunas = ''  # Será preenchida na primeira iteração
        self.tabela = 'public.fichas_monitoria'  # Tabela destino
    
    
    def leitura(self, caminho, nome_ficha):
        """
        Lê arquivo Excel e extrai apenas as colunas relevantes.
        
        Args:
            caminho (str): Caminho completo do arquivo .xls
            nome_ficha (str): Nome da ficha para tratamentos especiais
            
        Returns:
            pandas.DataFrame: DataFrame com colunas padronizadas
            
        Lógica complexa:
        1. Lê HTML do Excel (arquivos .xls são na verdade HTML)
        2. Identifica colunas dinâmicas: assertividade, distribuidora, protocolo
        3. Monta lista de colunas fixas + dinâmicas
        4. Trata exceções para fichas específicas
        5. Adiciona coluna 'protocolo' se não existir
        6. Adiciona coluna 'distribuidora' vazia se não existir
        """
        # pandas.read_html lê tabelas de arquivos HTML/Excel
        # keep_default_na=False: não converte strings vazias em NaN
        # decimal=',': vírgula como separador decimal (padrão BR)
        # thousands='.': ponto como separador de milhar
        leitura = pandas.read_html(caminho, keep_default_na=False, decimal=',', thousands='.')
        
        # Validação: deveria ter apenas 1 planilha
        if len(leitura) != 1:
            pass  # Continua mesmo se tiver mais de uma
        
        leitura = leitura[0]  # Pega a primeira planilha
        
        # Normaliza nomes das colunas para lowercase
        colunas = leitura.columns.tolist()
        colunas = [str(coluna).lower() for coluna in colunas]
        leitura.columns = colunas
        
        # Busca colunas que contêm palavras-chave (case-insensitive)
        assertividade = leitura.columns.str.contains('assertividade', case=False).tolist()
        distribuidora = leitura.columns.str.contains('distribuidora', case=False).tolist()
        protocolo = leitura.columns.str.contains('protocolo', case=False).tolist()
        
        # Converte listas booleanas em índices das colunas encontradas
        # Exemplo: [False, True, False] -> [1]
        assertividade = [contagem for contagem, linha in enumerate(assertividade) if linha]
        distribuidora = [contagem for contagem, linha in enumerate(distribuidora) if linha]
        protocolo = [contagem for contagem, linha in enumerate(protocolo) if linha]
        
        # Colunas fixas esperadas em todas as fichas
        lista_de_colunas = [
            'matricula', 
            'nome_funcionario', 
            'data da monitoria', 
            'data_ligacao', 
            'cod_monitoria', 
            'num_monitoria',
            'perfil_monitoria', 
            'nome_monitor'
        ]
        
        # TRATAMENTO ESPECIAL: Algumas fichas têm múltiplas colunas de assertividade
        if len(assertividade) != 1:
            # Fichas conhecidas que têm duplicatas
            if nome_ficha == 'CNR - COBE - REGIONAL 2022':
                assertividade = [assertividade[0]]  # Pega só a primeira
            elif nome_ficha == 'CNR - SCOB - REGIONAL 2022':
                assertividade = [assertividade[0]]
            else:
                # Se não é ficha conhecida, alerta e retorna sem processar
                print(f'Assertividade diferente que o previsto = {len(assertividade)} - {nome_ficha}')
                return leitura
        
        # TRATAMENTO ESPECIAL: Algumas fichas têm múltiplas colunas de distribuidora
        if len(distribuidora) > 1:
            # Bug no código original: o 'or' sempre retorna True
            # Deveria ser 'if nome_ficha in [...]'
            if nome_ficha in ['FICHA - HABILIDADE DE TRATAMENTO 2025 - NOTA RC', 'FICHA DA REC. HABILIDADE DE TRATAMENTO - NOTA RC']:
                distribuidora = [distribuidora[0]]
            else:
                print(f'Distribuidora maior que o previsto = {len(distribuidora)} - {nome_ficha}')
                return leitura
        
        # Adiciona as colunas dinâmicas encontradas
        for linha in assertividade:
            lista_de_colunas.append(leitura.columns[linha])
        for linha in distribuidora:
            lista_de_colunas.append(leitura.columns[linha])
        
        # Tratamento para protocolo (opcional)
        if len(protocolo) > 0:
            protocolo = [protocolo[0]]  # Pega só o primeiro
            for linha in protocolo:
                lista_de_colunas.append(leitura.columns[linha])
        else:
            # Se não tem coluna de protocolo, cria uma vazia
            leitura['protocolo'] = [None] * len(leitura)
            lista_de_colunas.append('protocolo')
        
        
        # Seleciona apenas as colunas da lista_de_colunas (descarta o resto)
        leitura = leitura.loc[:, lista_de_colunas]
        
        # Se não tem coluna distribuidora, cria uma vazia
        if len(distribuidora) == 0:
            lista_nova_coluna = ['' for linha in range(leitura.shape[0])]
            leitura['distribuidora'] = lista_nova_coluna
        

        return leitura
    
    
    def tratamento_do_dataframe(self, df):
        """
        Converte tipos de dados das colunas.
        
        Args:
            df (DataFrame): DataFrame após leitura
            
        Returns:
            DataFrame: Com tipos corretos
            
        Conversões:
        - data da monitoria → datetime com hora
        - data_ligacao → date (sem hora)
        - matricula → numérico
        - cod_monitoria → numérico
        - num_monitoria → numérico
        - assertividade (penúltima coluna -3) → numérico
        """
        # Converte strings de data para datetime
        df['data da monitoria'] = pandas.to_datetime(df['data da monitoria'], format='%d/%m/%Y %H:%M')
        df['data_ligacao'] = pandas.to_datetime(df['data_ligacao'], format='%d/%m/%Y').dt.date
        
        # Converte strings para números
        df['matricula'] = pandas.to_numeric(df['matricula'])
        df['cod_monitoria'] = pandas.to_numeric(df['cod_monitoria'])
        df['num_monitoria'] = pandas.to_numeric(df['num_monitoria'])
        
        # Assertividade está sempre na posição -3 (terceira coluna do final)
        # Estrutura: [...colunas fixas..., assertividade, distribuidora, protocolo]
        df[df.columns[len(df.columns) - 3]] = pandas.to_numeric(df[df.columns[len(df.columns) - 3]])
        
        return df
    
    
    def colunas_adicionais(self, df, nome_ficha : str):
        """
        Adiciona coluna 'tipo_da_ficha' com o nome da ficha.
        
        Args:
            df (DataFrame): DataFrame processado
            nome_ficha (str): Nome da ficha
            
        Returns:
            DataFrame: Com coluna adicional
            
        Propósito:
            Permite identificar de qual ficha vieram os dados após inserção
        """
        # Cria lista com o mesmo valor repetido para todas as linhas
        lista_nova_coluna = [nome_ficha for linha in range(df.shape[0])]
        df['tipo_da_ficha'] = lista_nova_coluna
        
        return df
    
    
    def deixando_no_modelo_datamart(self, df):
        """
        Padroniza nomes de colunas para o modelo do Datamart.
        
        Args:
            df (DataFrame): DataFrame com colunas dinâmicas
            
        Returns:
            DataFrame: Com nomes de colunas padronizados
            
        Lógica:
        - Identifica as 3 últimas colunas (sempre nessa ordem):
          posição -1: protocolo
          posição -2: distribuidora
          posição -3: assertividade
        - Renomeia para nomes fixos esperados pelo banco
        """
        # Identifica as 3 últimas colunas
        protocolo = df.columns[len(df.columns) - 1]
        distribuidora = df.columns[len(df.columns) - 2]
        assertividade = df.columns[len(df.columns) - 3]
        
        # Renomeia para padrão do Datamart
        df = df.rename(columns={
            'nome_funcionario': 'nome_do_funcionario',
            'data da monitoria': 'data_da_monitoria',
            assertividade: 'assertividade',  # Nome dinâmico → fixo
            distribuidora: 'distribuidora',  # Nome dinâmico → fixo
            protocolo: 'protocolo'  # Nome dinâmico → fixo
        })
        
        return df
    
    
    def run(self):
        """
        Método principal que executa todo o fluxo de processamento.
        
        Fluxo:
        1. Limpa dados do mês/ano no banco (resetando_mes)
        2. Para cada arquivo Excel:
           a. Extrai nome da ficha do nome do arquivo
           b. Lê e processa o Excel
           c. Valida se tem 11 colunas (padrão esperado)
           d. Trata tipos de dados
           e. Padroniza nomes
           f. Adiciona metadados (tipo_ficha, ano, mes)
           g. Insere linha por linha no banco
        """
        # Passo 1: Remove dados antigos do mês/ano
        self.conexao.resetando_mes(self.tabela, self.data_atual.month, self.data_atual.year)
        
        # Passo 2: Processa cada arquivo
        for caminho_ficha in self.lista_de_arquivos:
            # Skip para ficha específica (provavelmente problemática)
            if 'FICHA DA REC. HABILIDADE DE TRATAMENTO' in caminho_ficha:
                pass  # Não faz nada, mas continua o loop
            
            # Extrai o nome da ficha do caminho completo
            # Exemplo: "C:\path\01-2025 FICHA RECEPTIVO.xls" → "FICHA RECEPTIVO"
            caminho_relativo_ficha = os.path.dirname(caminho_ficha)
            nome_ficha = caminho_ficha.replace(caminho_relativo_ficha, '')
            nome_ficha = nome_ficha.replace('.xls', '')
            nome_ficha = nome_ficha.replace('\\', '')
            nome_ficha = nome_ficha.replace('{:02d}-{:04d} '.format(self.data_atual.month, self.data_atual.year), '')
            
            # Lê o Excel
            df = self.leitura(caminho_ficha, nome_ficha)
            if df.shape[1] != 11:
                print(f"❌ PULANDO: '{nome_ficha}' tem {df.shape[1]} colunas (esperado: 11)")
                print(f"   Colunas encontradas: {df.columns.tolist()}")
                continue
            
            # Pipeline de transformação
            df = self.tratamento_do_dataframe(df=df)
            df = self.deixando_no_modelo_datamart(df=df)
            df = self.colunas_adicionais(df=df, nome_ficha=nome_ficha)
            
            # Adiciona metadados de período
            df['ano'] = self.data_atual.year
            df['mes'] = self.data_atual.month
            
            # Na primeira iteração, captura os nomes das colunas
            if self.colunas == '':
                self.colunas = df.columns.tolist()
                colunas_string = ','.join(self.colunas)
            
            # Converte DataFrame para lista de listas
            lista_insercao = df.values.tolist()
            
            # Insere cada linha no banco
            for linha in lista_insercao:
                self.conexao.registro_sucesso_list(self.tabela, colunas_string, linha)
