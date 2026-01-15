"""
Módulo: __init__.py
Descrição: Arquivo especial que marca um diretório como pacote Python
Localização: Presente em static/, extracao_site/, insercao_datamart/, att_importacao_datamart/
Autor: Automação e Inovação - Contact Center
"""

# Arquivo vazio intencionalmente

"""
O QUE É __init__.py?

O arquivo __init__.py é um arquivo especial do Python que transforma um diretório
comum em um "pacote Python". Isso permite que você importe módulos desse diretório
usando a sintaxe de ponto.

EXEMPLO DE ESTRUTURA:
projeto/
├── static/
│   ├── __init__.py          ← Marca 'static' como pacote
│   ├── postgresql.py
│   └── registrar_consultar.py
└── gerenciador/
    └── main.py

SEM __init__.py:
    from static.postgresql import Conexao_postgresql  ❌ ERRO

COM __init__.py:
    from static.postgresql import Conexao_postgresql  ✅ FUNCIONA


QUANDO ESTÁ VAZIO (COMO ESTE):
- Serve apenas para marcar o diretório como pacote
- Não executa nenhum código adicional
- É o caso mais comum

QUANDO TEM CÓDIGO:
Você pode colocar código que será executado quando o pacote for importado.

Exemplos de uso com código:

1. Importar módulos para facilitar acesso:
   # __init__.py
   from .postgresql import Conexao_postgresql
   from .registrar_consultar import registers
   
   # Agora pode importar direto do pacote:
   from static import Conexao_postgresql, registers

2. Definir __all__ (lista de exports públicos):
   # __init__.py
   __all__ = ['Conexao_postgresql', 'registers']

3. Inicialização de pacote:
   # __init__.py
   print("Pacote static foi importado")
   configuracao = carregar_config()

4. Versionamento:
   # __init__.py
   __version__ = '1.0.2'


IMPORTANTE NO PYTHON 3.3+:
Tecnicamente, desde Python 3.3, o __init__.py não é mais obrigatório
(existe o conceito de "namespace packages"). MAS:

✅ BOA PRÁTICA: Sempre incluir __init__.py
   - Compatibilidade com código mais antigo
   - Deixa explícito que é um pacote
   - Permite adicionar código de inicialização no futuro
   - Alguns IDEs e ferramentas ainda esperam o arquivo

NESTE PROJETO:
Os __init__.py estão vazios porque:
- Não há necessidade de código de inicialização
- Os imports são feitos diretamente dos módulos
- Serve apenas para marcar os diretórios como pacotes Python válidos


LOCALIZAÇÃO DOS __init__.py NESTE PROJETO:
- static/__init__.py           (módulos auxiliares)
- extracao_site/__init__.py    (extração Playwright)
- insercao_datamart/__init__.py (importação para public.fichas_monitoria)
- att_importacao_datamart/__init__.py (importação para schema monitoria)
- gerenciador/__init__.py      (orquestrador - se existir)
"""

# FIM DO ARQUIVO
# Novamente: vazio é proposital e correto neste caso!
