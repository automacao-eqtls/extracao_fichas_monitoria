"""
Módulo: att_importacao_datamart/main.py
Descrição: Processa fichas já extraídas e organiza no schema 'monitoria' do banco
Autor: Automação e Inovação - Contact Center
Diferencial: Estrutura normalizada com múltiplas tabelas (fichas, colunas, registros, registro_coluna)
"""

# %%
import sys
import os

# Configura caminhos para importar módulos do projeto
caminho_relativo = os.path.dirname(__file__)
caminho_relativo = os.path.dirname(caminho_relativo)
sys.path.append(caminho_relativo)

# Define ordem de diretórios no path
ordem_das_etapas = ['static', 'insercao_datamart', 'extracao_site', 'venv']
for etapa in ordem_das_etapas:
    sys.path.append(os.path.join(caminho_relativo, etapa))

import datetime
from static.registrar_consultar import registers
import pandas

# Inicializa classe de conexão
inst_register = registers()

# Carrega todas as tabelas do schema 'monitoria' em memória
# Isso é feito para verificações rápidas antes de inserir
fichas = inst_register.consultar_notas('select * from monitoria.fichas')
registro_coluna = inst_register.consultar_notas('select * from monitoria.registro_coluna')
colunas = inst_register.consultar_notas('select * from monitoria.colunas')
registros = inst_register.consultar_notas('select * from monitoria.registros')

# Converte listas de dicionários em DataFrames para manipulação com Pandas
colunas = pandas.DataFrame(colunas)
registros = pandas.DataFrame(registros)
fichas = pandas.DataFrame(fichas)
registro_coluna = pandas.DataFrame(registro_coluna)

# %%
# Configuração de caminhos e datas
caminho_fichas_excel = r'\\55aspdcarq01\55atende\Administrativo\06 - GerÃªncia Contact Center\03 - Call Center Sao Luis\02 - Monitoria de Qualidade\10.BASES'
data_atual = datetime.date.today() - datetime.timedelta(days=1)  # D-1

# Lista arquivos .xls do mês/ano atual
lista_de_arquivos = os.listdir(caminho_fichas_excel)
lista_de_arquivos = [
    os.path.join(caminho_fichas_excel, linha) 
    for linha in lista_de_arquivos 
    if linha.endswith('.xls') and linha.startswith('{:02d}-{:04d}'.format(data_atual.month, data_atual.year))
]


def inserir_valor(ficha_dataframe, num_monitoria_fk):
    """
    Insere registros e valores de colunas no banco.
    
    Args:
        ficha_dataframe (DataFrame): Dados da ficha com colunas já renomeadas para FK
        num_monitoria_fk (int): FK da coluna 'NUM_MONITORIA'
        
    Estrutura de dados:
        monitoria.registros: Cada linha da ficha (identificada por num_monitoria)
        monitoria.registro_coluna: Valores de cada coluna para cada registro
        
    Lógica:
        1. Para cada linha da ficha:
           a. Verifica se registro já existe (por num_monitoria + ficha_fk)
           b. Se não existe, insere em monitoria.registros
           c. Para cada coluna da linha:
              - Verifica se já existe em monitoria.registro_coluna
              - Se não existe, insere o valor
    
    Nota:
        Usa DataFrames em memória para verificação rápida ao invés de queries repetidas
    """
    for indice, linha in ficha_dataframe.iterrows():
        # Busca ou cria registro
        if registros.empty:
            # Se tabela registros está vazia, insere novo
            registro_fk = inst_register.registro_sucesso(
                dicionario={
                    'num_monitoria': linha[num_monitoria_fk],
                    'ficha_fk': ficha_fk
                }, 
                tabela='monitoria.registros'
            )
        else:
            # Busca registro existente
            resultado_busca = registros.loc[
                (registros['ficha_fk'] == ficha_fk) & 
                (registros['num_monitoria'] == linha[num_monitoria_fk])
            ]
            
            if resultado_busca.empty:
                # Registro não existe, insere
                registro_fk = inst_register.registro_sucesso(
                    dicionario={
                        'num_monitoria': linha[num_monitoria_fk],
                        'ficha_fk': ficha_fk
                    }, 
                    tabela='monitoria.registros'
                )
            else:
                # Registro existe, pega o FK
                registro_fk = resultado_busca.values[0][0]
    
        # Insere valores de cada coluna
        for coluna_fk_valor, valor in linha.items():
            if registro_coluna.empty:
                # Se tabela vazia, insere direto
                inst_register.registro_sucesso(
                    dicionario={
                        'coluna_fk': coluna_fk_valor,
                        'registro_fk': registro_fk,
                        'valor': valor
                    }, 
                    tabela='monitoria.registro_coluna'
                )
            else:
                # Verifica se combinação coluna+registro já existe
                resultado_busca = registro_coluna.loc[
                    (registro_coluna['coluna_fk'] == coluna_fk_valor) & 
                    (registro_coluna['registro_fk'] == registro_fk)
                ]
                
                if resultado_busca.empty:
                    # Não existe, insere
                    inst_register.registro_sucesso(
                        dicionario={
                            'coluna_fk': coluna_fk_valor,
                            'registro_fk': registro_fk,
                            'valor': valor
                        }, 
                        tabela='monitoria.registro_coluna'
                    )


def leitura_excel(caminho, ficha_fk):
    """
    Lê arquivo Excel e mapeia colunas para FK do banco.
    
    Args:
        caminho (str): Caminho do arquivo .xls
        ficha_fk (int): FK da ficha em monitoria.fichas
        
    Returns:
        tuple: (DataFrame com colunas renomeadas, FK da coluna NUM_MONITORIA)
        
    Processo:
        1. Lê Excel
        2. Para cada coluna do Excel:
           a. Verifica se coluna já existe em monitoria.colunas
           b. Se não existe, insere
           c. Renomeia coluna do DataFrame para o FK
        3. Identifica qual FK é da coluna NUM_MONITORIA
        
    Importante:
        As colunas do DataFrame são renomeadas para os IDs do banco,
        facilitando a inserção posterior
    """
    # Lê Excel (arquivos .xls são HTML)
    ficha_dataframe = pandas.read_html(caminho, keep_default_na=False, decimal=',', thousands='.')
    
    # Validação: deve ter exatamente 1 planilha
    if len(ficha_dataframe) != 1:
        raise AssertionError('Pandas leu o excel mas o tamanho de planilha foi diferente de 1')

    ficha_dataframe = ficha_dataframe[0]
    
    # Obtém lista de colunas do Excel
    colunas_excel = ficha_dataframe.columns.values.tolist()

    # Para cada coluna do Excel
    for coluna in colunas_excel:
        # Busca ou cria a coluna no banco
        if colunas.empty:
            # Tabela vazia, insere nova coluna
            coluna_fk = inst_register.registro_sucesso(
                dicionario={
                    'nome_coluna': coluna,
                    'ficha_fk': ficha_fk
                }, 
                tabela='monitoria.colunas'
            )
        else:
            # Busca coluna existente
            resultado_busca = colunas.loc[
                (colunas['ficha_fk'] == ficha_fk) & 
                (colunas['nome_coluna'] == coluna)
            ]
            
            if resultado_busca.empty:
                # Não existe, insere
                coluna_fk = inst_register.registro_sucesso(
                    dicionario={
                        'nome_coluna': coluna,
                        'ficha_fk': ficha_fk
                    }, 
                    tabela='monitoria.colunas'
                )
            else:
                # Existe, pega FK
                coluna_fk = resultado_busca.values[0][0]
        
        # Identifica se é a coluna NUM_MONITORIA (chave primária da ficha)
        if coluna == 'NUM_MONITORIA':
            num_monitoria_fk = coluna_fk
        
        # IMPORTANTE: Renomeia coluna do DataFrame para o FK do banco
        # Exemplo: 'MATRICULA' vira 123 (se FK da coluna MATRICULA for 123)
        ficha_dataframe.rename(columns={coluna: coluna_fk}, inplace=True)
        
    return (ficha_dataframe, num_monitoria_fk)

# %%
"""
Loop principal: Processa cada arquivo Excel
"""
for indice_lista_arquivos, ficha in enumerate(lista_de_arquivos):
    # Extrai nome da ficha do caminho
    split_ficha = ficha.split('\\')
    ficha = split_ficha[len(split_ficha) - 1]
    ficha = ficha.replace('.xls', '')
    ficha = ficha.replace('{:02d}-{:04d}'.format(data_atual.month, data_atual.year), '')
    ficha = ficha.strip()
    
    # Busca ou cria a ficha no banco
    if fichas.empty:
        # Tabela vazia, insere nova ficha
        ficha_fk = inst_register.registro_sucesso(
            dicionario={
                'nome_ficha': ficha
            }, 
            tabela='monitoria.fichas'
        )
    else:
        # Busca ficha existente (por nome + mês + ano)
        resultado_busca = fichas.loc[
            (fichas['nome_ficha'] == ficha) & 
            (fichas['mes'] == data_atual.month) & 
            (fichas['ano'] == data_atual.year)
        ]
        
        if resultado_busca.empty:
            # Não existe, insere
            ficha_fk = inst_register.registro_sucesso(
                dicionario={
                    'nome_ficha': ficha,
                    'mes': data_atual.month,
                    'ano': data_atual.year
                }, 
                tabela='monitoria.fichas'
            )
        else:
            # Existe, pega FK
            ficha_fk = resultado_busca.values[0][0]
    
    # Lê o Excel e mapeia colunas
    (dataframe, num_monitoria_fk) = leitura_excel(
        caminho=lista_de_arquivos[indice_lista_arquivos], 
        ficha_fk=ficha_fk
    )
    
    # Insere registros e valores
    inserir_valor(ficha_dataframe=dataframe, num_monitoria_fk=num_monitoria_fk)


"""
ARQUITETURA DE DADOS:

Schema: monitoria

Tabelas:
1. fichas
   - id (PK)
   - nome_ficha
   - mes
   - ano

2. colunas
   - id (PK)
   - nome_coluna
   - ficha_fk (FK → fichas.id)

3. registros
   - id (PK)
   - num_monitoria
   - ficha_fk (FK → fichas.id)

4. registro_coluna (tabela de relacionamento)
   - id (PK)
   - coluna_fk (FK → colunas.id)
   - registro_fk (FK → registros.id)
   - valor (TEXT)

VANTAGENS:
- Estrutura normalizada (3ª forma normal)
- Flexível: suporta fichas com colunas diferentes
- Histórico preservado: mes/ano nas fichas
- Queries complexas possíveis (joins entre tabelas)

DESVANTAGENS:
- Performance: muitas queries para inserir
- Complexidade: precisa joins para consultas simples
- Não otimizado para análises (melhor seria star schema)
"""
