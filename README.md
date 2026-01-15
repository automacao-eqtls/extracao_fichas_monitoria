# Automação de Extração - OPTIMUS

RPA que acessa o frontend do Optimus, extrai as bases, salva em pasta local e insere no DataMart.

## ⚙️ Fluxo e Regras

* **Destino:** Tabela `public.fichas_monitorias`.
* **Verificação:** Checa a coluna `carimbo_tempo` na tabela public.hist_bases antes de executar para ver se já foi executado no dia.
* **Limite:** Máximo de **3 tentativas** de inserção por dia.

## 📝 Logs (`public.hist_bases`)

* `mensagem_erro`: Registra os logs de erros.
* `concluido`: Indica se a execução foi finalizada com sucesso.

## ▶️ Execução

```bash
./gerenciador/main.py

```

## ⚠️ Configuração

É necessário alterar os caminhos das variáveis abaixo conforme a máquina onde o robô for executado:

* `destination_folder_path`
* `caminho_relativo`