## Célula 1: Instalação de Bibliotecas
# Para instalar as bibliotecas necessárias, execute o comando abaixo no terminal do Windows (PowerShell):
#
# pip install --upgrade google-api-python-client google-auth google-auth-httplib2 google-auth-oauthlib

## Célula 2: Acesso a arquivos locais
# Substitua o acesso ao Google Drive por caminhos locais. Exemplo:
# caminho_da_pasta = r'C:\Users\Dell\Desktop\TCC\Materiais_e_Metodos\Dados\'

## Célula 3: Autenticação da Conta de Serviço
import os
from google.oauth2 import service_account
from googleapiclient.discovery import build

SERVICE_ACCOUNT_FILE = 'mba-usp-469620-7ff99acd80ca.json'  # Certifique-se de que o arquivo está na mesma pasta do script ou forneça o caminho completo
try:
    credentials = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=['https://www.googleapis.com/auth/spreadsheets.readonly'])
    print("Credenciais carregadas com sucesso!")
except FileNotFoundError:
    print(f"ERRO: O arquivo de credenciais '{SERVICE_ACCOUNT_FILE}' não foi encontrado.")
    print("Por favor, verifique se o arquivo JSON está na pasta correta e se o nome está correto.")
except Exception as e:
    print(f"Ocorreu um erro ao carregar as credenciais: {e}")

if 'credentials' in locals():
    service = build('sheets', 'v4', credentials=credentials)
    print("Serviço Google Sheets construído com sucesso!")
else:
    print("Não foi possível construir o serviço Google Sheets devido a erros de credenciais.")

# Célula 4: Leitura dos Dados da Planilha
SPREADSHEET_ID = '1sTiUrfRTddQLssaYDR38ouGwfmkSUP0rVY4ssO2dc2A'
RANGE_NAME = 'Respostas ao formulário 1!A1:AD42'

try:
    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID, range=RANGE_NAME).execute()
    values = result.get('values', [])
    if not values:
        print('Nenhum dado encontrado no range especificado.')
    else:
        print(f'Dados lidos com sucesso! Número de linhas: {len(values)}')
except Exception as e:
    print(f"Ocorreu um erro ao ler os dados da planilha: {e}")

# Célula 5: Carregar Dados em DataFrame e Exibir Head
import pandas as pd
import numpy as np
headers = values[0]
data = values[1:]

if 'values' in locals() and values:
    headers = values[0]
    data = values[1:]
    df_raw = pd.DataFrame(data, columns=headers)
    print("✅ DataFrame criado com sucesso!")
    df_raw.columns = df_raw.columns.str.strip()
    print("✅ Espaços removidos dos nomes das colunas em df_raw.")
    pd.set_option('display.max_columns', None)
    display(df_raw.head())

# Célula 6: Contagem de Respostas
print(f"Total de respostas (participantes): {df_raw.shape[0]}")
print(f"Total de perguntas (colunas): {df_raw.shape[1]}")
print("\n")

# Célula 7: Título para Análise de Frequência
print("### Análise de Frequência - Perfil da Amostra ###")
print("-" * 40)

# Célula 8: Mapeamento e Renomeação de Colunas
import pandas as pd
import numpy as np
df_raw.columns = df_raw.columns.str.strip()
column_mapping = {
    'Carimbo de data/hora': 'timestamp',
    'Qual a sua idade e nível de escolaridade? \\n(Por favor, marque a opção que melhor te representa.) [até 25 anos]': 'idade_ate_25',
    'Qual a sua idade e nível de escolaridade? \\n(Por favor, marque a opção que melhor te representa.) [26 a 35]': 'idade_26_35',
    'Qual a sua idade e nível de escolaridade? \\n(Por favor, marque a opção que melhor te representa.) [36 a 45]': 'idade_36_45',
    'Qual a sua idade e nível de escolaridade? \\n(Por favor, marque a opção que melhor te representa.) [46 ou mais]': 'idade_46_ou_mais',
    'Como você se autodeclara em relação à sua raça/cor e com qual gênero você se identifica?\\n(Por favor, marque a opção que melhor te representa.) [Preta]': 'preta',
    'Como você se autodeclara em relação à sua raça/cor e com qual gênero você se identifica?\\n(Por favor, marque a opção que melhor te representa.) [Parda]': 'parda',
    'Como você se autodeclara em relação à sua raça/cor e com qual gênero você se identifica?\\n(Por favor, marque a opção que melhor te representa.) [Indígena]': 'indigena',
    'Como você se autodeclara em relação à sua raça/cor e com qual gênero você se identifica?\\n(Por favor, marque a opção que melhor te representa.) [Amarela]': 'amarela',
    'Como você se autodeclara em relação à sua raça/cor e com qual gênero você se identifica?\\n(Por favor, marque a opção que melhor te representa.) [Branca]': 'branca',
    'Como você se autodeclara em relação à sua raça/cor e com qual gênero você se identifica?\\n(Por favor, marque a opção que melhor te representa.) [Não responder]': 'nao_respondeu_raca',
    'Em qual estado da federação você reside?': 'estado',
    'Há quanto tempo atua na área de tecnologia?': 'tempo_experiencia',
    'Qual sua especialidade e senioridade atual?\\n(Por favor, para sua principal especialidade atual, marque sua senioridade ) [Dev / Engenheira de Software]': 'dev_eng_software',
    'Qual sua especialidade e senioridade atual?\\n(Por favor, para sua principal especialidade atual, marque sua senioridade ) [DevOps]': 'devops',
    'Qual sua especialidade e senioridade atual?\\n(Por favor, para sua principal especialidade atual, marque sua senioridade ) [Dados]': 'dados',
    'Qual sua especialidade e senioridade atual?\\n(Por favor, para sua principal especialidade atual, marque sua senioridade ) [Suporte]': 'suporte',
    'Qual sua especialidade e senioridade atual?\\n(Por favor, para sua principal especialidade atual, marque sua senioridade ) [Infraestrutura]': 'infraestrutura',
    'Qual sua especialidade e senioridade atual?\\n(Por favor, para sua principal especialidade atual, marque sua senioridade ) [QA]': 'qa',
    'Qual sua especialidade e senioridade atual?\\n(Por favor, para sua principal especialidade atual, marque sua senioridade ) [SI]': 'si',
    'Qual sua especialidade e senioridade atual?\\n(Por favor, para sua principal especialidade atual, marque sua senioridade ) [UX/UI]': 'ux_ui',
    'Qual sua especialidade e senioridade atual?\\n(Por favor, para sua principal especialidade atual, marque sua senioridade ) [PO]': 'po',
    'Qual sua especialidade e senioridade atual?\\n(Por favor, para sua principal especialidade atual, marque sua senioridade ) [Negócios]': 'negocios',
    'Se você ingressou na área de tecnologia por vaga afirmativa, qual(is) era(m) o(s) critério(s) de elegibilidade que mais se destacavam? (Marque todas as opções que se aplicam)': 'criterios_vaga_afirmativa',
    'O ingresso por vaga afirmativa teve um impacto positivo na minha trajetória profissional na tecnologia.': 'impacto_vaga_afirmativa',
    'Sinto que minha presença e contribuições são valorizadas no meu ambiente de trabalho.': 'sentimento_valorizacao',
    'Minha empresa (atual ou mais recente que atuava na área de tecnologia) oferece oportunidades de desenvolvimento e crescimento profissional satisfatórias para mulheres negras.': 'oportunidades_desenvolvimento',
    'Já enfrentei barreiras ou preconceito (racial e/ou de gênero) que dificultaram minha trajetória na área. ': 'enfrentou_barreiras',
    'Quais fatores mais te motivam (ou motivariam) a permanecer e progredir em sua carreira na área de tecnologia? (Por favor, selecione até 3 fatores principais)': 'fatores_permanencia',
    'Qual a sua opinião em relação as vagas afirmativas? Se tiver e puder, deixe o relato de sua experiência e o impacto em sua trajetória.': 'opiniao_vagas_afirmativas'
}
column_mapping = {key.strip(): value for key, value in column_mapping.items()}
df = df_raw.rename(columns=column_mapping, errors='ignore')
print("✅ Colunas do DataFrame 'df' renomeadas.")
df.head(0)

# Célula 9: Exibindo o DataFrame renomeado
df.head()

# Célula 10: Função de Consolidação - Idade e Escolaridade
def consolidate_idade_escolaridade(row):
    for col in [c for c in row.index if c.startswith('idade_ate_25') or c.startswith('idade_26_35') or c.startswith('idade_36_45') or c.startswith('idade_46_ou_mais')]:
        value = row[col]
        if pd.notna(value) and str(value).strip() != '':
            idade = col
            escolaridade = value
            return pd.Series([idade, escolaridade])
    return np.nan

# Célula 11: Função de Consolidação - Gênero e Raça
def consolidate_genero_raca_pair(row):
    for col in [c for c in row.index if c.startswith('preta') or c.startswith('parda') or c.startswith('indigena') or c.startswith('amarela') or c.startswith('branca') or c.startswith('não responder')]:
        value = row[col]
        if pd.notna(value) and str(value).strip() != '':
            raca = col
            genero = value
            return pd.Series([genero, raca])
    return pd.Series([np.nan, np.nan])

# Célula 12: Função de Consolidação - Especialidade e Senioridade
def consolidate_especialidade_senioridade(row):
    for col in [c for c in row.index if c.startswith('dev_eng_software') or c.startswith('devops') or c.startswith('dados') or c.startswith('suporte') or c.startswith('infraestrutura') or c.startswith('qa') or c.startswith('si') or c.startswith('ux_ui') or c.startswith('po') or c.startswith('negocios')]:
        value = row[col]
        if pd.notna(value) and str(value).strip() != '':
            especialidade = col
            senioridade = value
            return pd.Series([especialidade, senioridade])
    return np.nan

# Célula 13 (Markdown): Agrupamento Racial
# **Agrupamento Racial: Criando a Categoria "Mulheres Negras"**
# Seu tema central são as "mulheres negras". Na pesquisa em ciências sociais no Brasil, é comum (e metodologicamente embasado pelo IBGE) agrupar as autodeclarações de "Preta" e "Parda" na categoria "Negra". Isso pode fortalecer suas análises, aumentando a representatividade do grupo focal.

# Célula 14: Função para Agrupar Raça
def agrupar_raca(raca):
    if raca in ['preta', 'parda']:
        return 'Negra'
    else:
        return raca.capitalize()

# Célula 15: Aplicação das Funções de Consolidação
df[['genero', 'raca']] = df.apply(consolidate_genero_raca_pair, axis=1)
df[['idade','escolaridade']] = df.apply(consolidate_idade_escolaridade, axis=1)
df[['senioridade', 'especialidade']] = df.apply(consolidate_especialidade_senioridade, axis=1)
df['raca_agrupada'] = df['raca'].apply(agrupar_raca)

# Célula 16: Exibindo Novas Colunas
print("✅ Novas colunas de matrizes consolidadas com sucesso!")
print(df[['genero', 'raca', 'idade', 'escolaridade', 'senioridade', 'especialidade', 'criterios_vaga_afirmativa','raca_agrupada']].head(0))
print("\n")

# Célula 17 (Markdown): Tratamento de Variáveis Ordinais
# **Tratamento de Variáveis Ordinais (Escala Likert)**
# Perguntas como impacto_vaga_afirmativa, sentimento_valorizacao, oportunidades_desenvolvimento e enfrentou_barreiras utilizam uma escala de concordância (Likert). Atualmente, elas estão como texto (object). Para análises futuras e para que os gráficos fiquem na ordem correta (de "Discordo" para "Concordo"), é uma boa prática convertê-las para um tipo de dado categórico ordenado.

# Célula 18: Conversão de Escala Likert para Categórica Ordenada
from pandas.api.types import CategoricalDtype
likert_order = ['Discordo Totalmente', 'Discordo', 'Neutro', 'Concordo', 'Concordo Totalmente']
cat_type = CategoricalDtype(categories=likert_order, ordered=True)
likert_columns = ['impacto_vaga_afirmativa', 'sentimento_valorizacao', 'oportunidades_desenvolvimento', 'enfrentou_barreiras']
for col in likert_columns:
    df[col] = df[col].astype(cat_type)
print("✅ Colunas de escala Likert convertidas para tipo categórico ordenado.")

# Célula 19: Normalização da Coluna 'estado'
df['estado_limpo'] = df['estado'].str.lower().str.strip().str.normalize('NFKD').str.encode('ascii', errors='ignore').str.decode('utf-8')
estado_mapping = {
    'sao paulo': 'São Paulo', 'sp': 'São Paulo', 'sao': 'São Paulo',
    'pernambuco ': 'Pernambuco', 'pernambuco': 'Pernambuco', 'pe': 'Pernambuco',
    'pr': 'Paraná', 'parna - pr': 'Paraná', 'parana': 'Paraná',
    'bahia': 'Bahia', 'ceara': 'Ceará', 'es': 'Espírito Santo',
    'minas gerais': 'Minas Gerais', 'mg': 'Minas Gerais',
    'rio de janeiro': 'Rio de Janeiro', 'rio de janeiro ': 'Rio de Janeiro',
    'df': 'Distrito Federal', 'rio grande do sul': 'Rio Grande do Sul',
    'sergipe': 'Sergipe', 'santa catarina': 'Santa Catarina',
    'goias': 'Goiás', 'mato grosso': 'Mato Grosso', 'alagoas': 'Alagoas',
    'paraiba': 'Paraíba', 'mato grosso do sul': 'Mato Grosso do Sul',
    'tocantins': 'Tocantins', 'roraima': 'Roraima', 'acre': 'Acre',
    'amazonas': 'Amazonas', 'amapa': 'Amapá', 'rondonia': 'Rondônia',
    'para': 'Pará', 'maranhao': 'Maranhão', 'piaui': 'Piauí',
    'rio grande do norte': 'Rio Grande do Norte'
}
df['estado_normalizado'] = df['estado_limpo'].map(estado_mapping).fillna(df['estado_limpo'].str.title())
print("### Distribuição dos estados antes da normalização ###")
print(df['estado'].value_counts().to_string())
print("\n### Distribuição dos estados depois da normalização ###")
print(df['estado_normalizado'].value_counts().to_string())
df.drop(columns=['estado'], inplace=True, errors='ignore')
df.rename(columns={'estado_normalizado': 'estado'}, inplace=True)
print("\n✅ Normalização concluída. DataFrame atualizado.")

# Célula 20: Frequência da Coluna 'estado'
print("3. Distribuição de 'estado' :")
if 'estado' in df.columns:
    fatores_permanencia_counts = df['estado'].value_counts(normalize=True).mul(100).round(1)
    print(fatores_permanencia_counts.to_string())
else:
    print("Column 'estado' not found in DataFrame.")
print("\n")

# Células 21-23: describe(), info(), isnull()
df.describe()
df.info()
df.isnull()
df.head()

# Célula 24: Salvando o DataFrame Limpo
## Atenção: Não coloque barra invertida no final de string raw!
caminho_da_pasta = r'g:\Meu Drive\MBA\TCC\Materiais_e_Metodos\Dados'  # Ajuste para o seu ambiente Google Drive
nome_do_arquivo = 'dados_limpos.csv'
os.makedirs(caminho_da_pasta, exist_ok=True)
caminho_completo = os.path.join(caminho_da_pasta, nome_do_arquivo)
df.to_csv(caminho_completo, index=False)
print(f"Arquivo '{nome_do_arquivo}' salvo com sucesso em:")
print(caminho_da_pasta)