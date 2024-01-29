import os
import pandas
from static.registrar_consultar import registers

class main:
    caminho_relativo = r'\\55ASPDCARQ01\55Atende\Administrativo\06 - Gerência Contact Center\03 - Call Center Sao Luis\02 - Monitoria de Qualidade\10.BASES'
    
    def __init__(self, data_extracao) -> None:
        self.data_atual = data_extracao
        lista_de_arquivos = os.listdir(self.__class__.caminho_relativo)
        self.lista_de_arquivos = [os.path.join(self.__class__.caminho_relativo, linha) for linha in lista_de_arquivos if linha.endswith('.xls') and linha.startswith('{:02d}-{:04d}'.format(self.data_atual.month,self.data_atual.year))]
        self.conexao = registers()
        self.colunas = ''
        self.tabela = 'public.fichas_monitoria'
    
    
    def leitura(self, caminho, nome_ficha):
        leitura = pandas.read_html(caminho, keep_default_na=False, decimal=',', thousands='.')
        if len(leitura) != 1:
            pass
        
        leitura = leitura[0]
        
        colunas = leitura.columns.tolist()
        colunas = [str(coluna).lower() for coluna in colunas]
        
        leitura.columns = colunas
        
        assertividade = leitura.columns.str.contains('assertividade', case=False).tolist()
        distribuidora = leitura.columns.str.contains('distribuidora', case=False).tolist()
        protocolo = leitura.columns.str.contains('protocolo', case=False).tolist()
        
        assertividade = [contagem for contagem, linha in enumerate(assertividade) if linha]
        distribuidora = [contagem for contagem, linha in enumerate(distribuidora) if linha]
        protocolo = [contagem for contagem, linha in enumerate(protocolo) if linha]
        
        lista_de_colunas = ['matricula', 'nome_funcionario', 'data da monitoria', 'data_ligacao', 'cod_monitoria', 
                            'num_monitoria','perfil_monitoria', 'nome_monitor']
        
        if len(assertividade) != 1:
            if nome_ficha == 'CNR - COBE - REGIONAL 2022':
                assertividade = [assertividade[0]]
            elif nome_ficha == 'CNR - SCOB - REGIONAL 2022':
                assertividade = [assertividade[0]]
            else:
                print(f'Assertividade diferente que o previsto = {len(assertividade)} - {nome_ficha}')
                return leitura
            
        if len(distribuidora) > 1:
            if nome_ficha == 'FICHA DE RECLAMAÇÃO - HABILIDADE DE TRATAMENTO - NOTA RC - 2022.':
                distribuidora = [distribuidora[0]]
            else:
                print(f'Distribuidora maior que o previsto = {len(distribuidora)} - {nome_ficha}')
                return leitura
            
        for linha in assertividade:
            lista_de_colunas.append(leitura.columns[linha])
        for linha in distribuidora:
            lista_de_colunas.append(leitura.columns[linha])
            
        if len(protocolo) > 0:
            protocolo = [protocolo[0]]
            for linha in protocolo:
                lista_de_colunas.append(leitura.columns[linha])
        else:
            leitura['protocolo'] = [None] * len(leitura)
            lista_de_colunas.append('protocolo')
            
        
        
        leitura = leitura.loc[:, lista_de_colunas]
        if len(distribuidora) == 0:
            lista_nova_coluna = ['' for linha in range(leitura.shape[0])]
            leitura['distribuidora'] = lista_nova_coluna
        

        return leitura
    
    
    def tratamento_do_dataframe(self, df):
        df['data da monitoria'] = pandas.to_datetime(df['data da monitoria'], format='%d/%m/%Y %H:%M')
        df['data_ligacao'] = pandas.to_datetime(df['data_ligacao'], format='%d/%m/%Y').dt.date
        df['matricula'] = pandas.to_numeric(df['matricula'])
        df['cod_monitoria'] = pandas.to_numeric(df['cod_monitoria'])
        df['num_monitoria'] = pandas.to_numeric(df['num_monitoria'])
        df[df.columns[len(df.columns) - 3]] = pandas.to_numeric(df[df.columns[len(df.columns) - 3]])
        
        return df
    
    
    def colunas_adicionais(self, df, nome_ficha : str):
        lista_nova_coluna = [nome_ficha for linha in range(df.shape[0])]
        df['tipo_da_ficha'] = lista_nova_coluna
        
        return df
    
    
    def deixando_no_modelo_datamart(self, df):
        protocolo = df.columns[len(df.columns) - 1]
        distribuidora = df.columns[len(df.columns) - 2]
        assertividade = df.columns[len(df.columns) - 3]
        df = df.rename(columns={'nome_funcionario' : 'nome_do_funcionario', 'data da monitoria' : 'data_da_monitoria', assertividade : 'assertividade', distribuidora : 'distribuidora', protocolo: 'protocolo'})
        
        return df
    
    
    def run(self):
        self.conexao.resetando_mes(self.tabela, self.data_atual.month, self.data_atual.year)
        for caminho_ficha in self.lista_de_arquivos:
            
            caminho_relativo_ficha = os.path.dirname(caminho_ficha)
            nome_ficha = caminho_ficha.replace(caminho_relativo_ficha, '')
            nome_ficha = nome_ficha.replace('.xls', '')
            nome_ficha = nome_ficha.replace('\\', '')
            nome_ficha = nome_ficha.replace('{:02d}-{:04d} '.format(self.data_atual.month,self.data_atual.year), '')
            df = self.leitura(caminho_ficha, nome_ficha)
            if df.shape[1] != 11:
                continue
            
            df = self.tratamento_do_dataframe(df=df)
            df = self.deixando_no_modelo_datamart(df=df)
            df = self.colunas_adicionais(df=df, nome_ficha=nome_ficha)
            
            df['ano'] = self.data_atual.year
            df['mes'] = self.data_atual.month
            
            if self.colunas == '':
                self.colunas = df.columns.tolist()
                colunas_string = ','.join(self.colunas)
                
            lista_insercao = df.values.tolist()
            for linha in lista_insercao:
                self.conexao.registro_sucesso_list(self.tabela, colunas_string, linha)           