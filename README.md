# Reorganizador de Projetos

Um script Python moderno para reorganizar dados de projetos em formato CSV, transformando-os em uma estrutura tabular e exportando para Excel com formatação aprimorada.

## Visão Geral

Este projeto foi desenvolvido para reorganizar dados de projetos que estão estruturados horizontalmente em planilhas CSV, transformando-os em um formato vertical mais amigável para análise. O script utiliza métodos modernos e eficientes do pandas, eliminando loops explícitos e aproveitando operações vetorizadas para melhor performance.

## Características

- Processamento de arquivos CSV com suporte a diferentes codificações
- Identificação automática de colunas relacionadas a projetos
- Reorganização de dados para formato tabular
- Criação de tabelas de resumo (por status e por respondente)
- Formatação avançada do Excel com cores e ajustes automáticos de largura
- Processamento vetorizado eficiente usando pandas moderno

## Instalação

### Pré-requisitos

- Python 3.6 ou superior
- pip (gerenciador de pacotes Python)

### Passos para instalação

1. Clone ou baixe este repositório para sua máquina local:
```bash
git clone https://github.com/seu-usuario/reorganizador-projetos.git
cd reorganizador-projetos
```

2. Crie um ambiente virtual (recomendado):
```bash
# Criação do ambiente virtual
python -m venv venv

# Ativação no Windows
venv\Scripts\activate

# Ativação no Linux/macOS
source venv/bin/activate
```

3. Instale as dependências necessárias:
```bash
pip install -r requirements.txt
```

Ou instale diretamente:
```bash
pip install pandas openpyxl
```

## Uso

### Forma básica

```bash
python reorganizador_csv_moderno.py caminho/para/seu/arquivo.csv
```

### Com opções adicionais

```bash
python reorganizador_csv_moderno.py caminho/para/seu/arquivo.csv --saida resultado.xlsx --encoding utf-8
```

### Opções disponíveis

- `arquivo_csv`: Caminho para o arquivo CSV de entrada (obrigatório)
- `--saida` ou `-s`: Caminho para o arquivo Excel de saída (padrão: "Projetos_Reorganizados.xlsx")
- `--encoding` ou `-e`: Codificação do arquivo CSV (padrão: "cp1252")

## Como Funciona: Explicação Detalhada

### 1. Leitura do Arquivo CSV

```python
df = pd.read_csv(arquivo_csv, encoding=encoding)
```

Esta linha lê o arquivo CSV especificado usando a biblioteca pandas, com a codificação fornecida. O resultado é armazenado em um DataFrame, que é a estrutura de dados principal do pandas para manipulação de dados tabulares.

### 2. Identificação de Colunas de Projetos

```python
colunas_por_tipo = identificar_colunas_projeto(df)
```

A função `identificar_colunas_projeto` analisa o DataFrame e identifica as colunas relacionadas a projetos usando expressões regulares. O resultado é um dicionário onde as chaves são os tipos de informação (nome, status, versão, autor) e os valores são listas de nomes de colunas correspondentes.

#### Como a Identificação de Colunas Funciona

```python
# Padrões regex para identificar colunas
padroes = {
    'nome': r'^Nome do Projeto(\d*)$',
    'status': r'^Status do Projeto.Meu projeto está:(\d*)$',
    'versao': r'^Versão do Projeto(\d*)$',
    'autor': r'^Autor \(Responsável pelo Projeto\)(\d*)$',
    'continuar': r'^Deseja adicionar outro projeto \?(\d*)$'
}

# Usando compreensões de dicionário
colunas_por_tipo = {
    tipo: [col for col in df.columns if re.match(padrao, col)]
    for tipo, padrao in padroes.items()
}
```

Esta abordagem usa compreensões de dicionário e expressões regulares para identificar as colunas correspondentes a cada tipo de informação, agrupando-as de acordo com o sufixo numérico.

### 3. Extração e Reorganização dos Dados

```python
# Criar pares de colunas relacionadas
pares_colunas = []
for i in range(max(len(grupo) for grupo in colunas_por_tipo.values())):
    par = {}
    for tipo, colunas in colunas_por_tipo.items():
        if i < len(colunas):
            par[tipo] = colunas[i]
    if par:
        pares_colunas.append(par)

# Processar cada conjunto de colunas
for par in pares_colunas:
    if 'nome' not in par:
        continue
        
    # Selecionar colunas relevantes
    colunas_selecionadas = [col for col in par.values() if col in df.columns]
    colunas_selecionadas.extend(['Email', 'Nome'])
    
    # Criar DataFrame temporário e processá-lo
    temp_df = df[colunas_selecionadas].copy()
    temp_df = temp_df[temp_df[par['nome']].notna()]
    
    # Renomear e concatenar
    temp_df = temp_df.rename(columns=mapeamento_colunas)
    projetos_df = pd.concat([projetos_df, temp_df], ignore_index=True)
```

Esta parte do código:
1. Agrupa colunas relacionadas (por exemplo, "Nome do Projeto1" com "Status do Projeto1")
2. Para cada grupo, cria um DataFrame temporário com as linhas onde o nome do projeto não é nulo
3. Renomeia as colunas para o formato padrão
4. Concatena com o DataFrame principal

### 4. Criação de Resumos

```python
# Resumo por Status
pivot_status = pd.pivot_table(
    projetos_df,
    values='Nome do Projeto',
    index=['Autor'],
    columns=['Status'],
    aggfunc='count',
    fill_value=0
)

# Resumo por Respondente
resumo_respondente = (projetos_df
                      .groupby('Nome Respondente', as_index=False)
                      .size()
                      .rename(columns={'size': 'Total de Projetos'}))
```

Esta parte cria:
1. Uma tabela dinâmica (pivot table) que mostra a contagem de projetos por autor e status
2. Um resumo que conta o número total de projetos por respondente

### 5. Exportação para Excel

```python
with pd.ExcelWriter(arquivo_saida, engine='openpyxl') as writer:
    projetos_df.to_excel(writer, sheet_name='Projetos', index=False)
    pivot_status.to_excel(writer, sheet_name='Resumo por Status')
    resumo_respondente.to_excel(writer, sheet_name='Resumo por Respondente', index=False)
```

Esta parte cria um arquivo Excel com três planilhas:
- "Projetos": lista completa de todos os projetos
- "Resumo por Status": tabela dinâmica
- "Resumo por Respondente": contagem de projetos por respondente

### 6. Formatação do Excel

```python
formatar_excel(arquivo_saida)
```

A função `formatar_excel` aplica formatação visual ao arquivo Excel:
- Cabeçalhos em negrito com fundo azul claro
- Ajuste automático da largura das colunas
- Alinhamento centralizado para cabeçalhos

## Estrutura de Arquivos

- `reorganizador_csv_moderno.py`: Script principal
- `requirements.txt`: Lista de dependências
- `README.md`: Este arquivo de documentação

## Exemplos de Resultados

Após executar o script, você obterá um arquivo Excel contendo:

1. **Planilha "Projetos"**:
   - Uma tabela organizada com as colunas: Nome do Projeto, Status, Versão, Autor, Email Respondente, Nome Respondente

2. **Planilha "Resumo por Status"**:
   - Uma tabela dinâmica que cruza autores e status, mostrando a contagem de projetos

3. **Planilha "Resumo por Respondente"**:
   - Uma tabela que mostra o número total de projetos por respondente

## Solução de Problemas

### Erro de codificação

Se encontrar erros relacionados à codificação, tente especificar explicitamente:

```bash
python reorganizador_csv_moderno.py seu_arquivo.csv --encoding utf-8
```

Ou tente outras codificações comuns como 'latin1', 'iso-8859-1', etc.

### Arquivo não encontrado

Certifique-se de que o caminho para o arquivo CSV está correto. Você pode:
- Fornecer o caminho completo
- Colocar o CSV no mesmo diretório do script
- Verificar se o nome do arquivo está correto, incluindo maiúsculas/minúsculas

### Erros de formatação

Se encontrar erros durante a formatação do Excel, o script continuará funcionando, mas o arquivo Excel resultante não terá a formatação visual aprimorada.

## Contribuições

Contribuições são bem-vindas! Para contribuir:

1. Faça um fork do repositório
2. Crie uma branch para sua feature (`git checkout -b feature/nova-feature`)
3. Commit suas mudanças (`git commit -am 'Adiciona nova feature'`)
4. Push para a branch (`git push origin feature/nova-feature`)
5. Crie um Pull Request

## Licença

Este projeto está licenciado sob a licença MIT - veja o arquivo LICENSE para detalhes.