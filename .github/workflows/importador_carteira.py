import gspread
import pandas as pd
from gspread_dataframe import set_with_dataframe
from google.oauth2.service_account import Credentials
from gspread_formatting import format_cell_range, CellFormat, Color
from datetime import datetime

# === CONFIGURAÇÕES ===
ORIGEM_ID = '1lUNIeWCddfmvJEjWJpQMtuR4oRuMsI3VImDY0xBp3Bs'
DESTINO_ID = '1gDktQhF0WIjfAX76J2yxQqEeeBsSfMUPGs5svbf9xGM'
ABA_ORIGEM = 'Carteira'
ABA_DESTINO = 'Carteira'
import json, os
cred_str = os.environ.get('GOOGLE_CREDENTIALS_JSON')
cred_dict = json.loads(cred_str)
CAMINHO_CREDENCIAIS = None

COLUNAS_ORIGEM = ['A', 'Z', 'B', 'C', 'D', 'E', 'U', 'T', 'N', 'AA', 'AB', 'CN', 'CQ', 'CR', 'CS', 'BQ', 'CE', 'V']

# === AUTENTICAÇÃO ===
escopos = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]
credenciais = Credentials.from_service_account_info(cred_dict, scopes=escopos)
cliente = gspread.authorize(credenciais)

# === ABRIR PLANILHAS ===
planilha_origem = cliente.open_by_key(ORIGEM_ID)
planilha_destino = cliente.open_by_key(DESTINO_ID)

aba_origem = planilha_origem.worksheet(ABA_ORIGEM)
aba_destino = planilha_destino.worksheet(ABA_DESTINO)

# === IDENTIFICAR ÍNDICES DAS COLUNAS ===
cabecalhos_completos = aba_origem.row_values(5)
col_indices = []
for letra in COLUNAS_ORIGEM:
    idx = gspread.utils.a1_to_rowcol(letra + '1')[1] - 1
    col_indices.append(idx)

# === OBTER DADOS A PARTIR DA LINHA 5 ===
dados_completos = aba_origem.get_all_values()
dados = dados_completos[4:]  # Linha 5 em diante

# === FILTRAR APENAS AS COLUNAS DESEJADAS ===
dados_filtrados = []
for linha in dados:
    if len(linha) > 0 and linha[0].strip():  # Verifica se coluna A está preenchida
        nova_linha = []
        for idx in col_indices:
            valor = linha[idx] if idx < len(linha) else ''
            nova_linha.append(valor)
        dados_filtrados.append(nova_linha)

# === MONTAR DATAFRAME COM CABEÇALHOS ===
cabecalhos_selecionados = [cabecalhos_completos[i] if i < len(cabecalhos_completos) else '' for i in col_indices]
df = pd.DataFrame(dados_filtrados, columns=cabecalhos_selecionados)

# === AJUSTES ESPECÍFICOS ===
# Converter coluna de Data (A) sem apagar os valores inválidos
col_data = cabecalhos_selecionados[0]
try:
    datas_convertidas = pd.to_datetime(df[col_data], dayfirst=True, errors='coerce')
    df[col_data] = datas_convertidas.fillna(df[col_data])
except Exception as e:
    print(f"Erro ao converter datas: {e}")

# Converter coluna de valor numérico (AC)
if "AC" in df.columns:
    try:
        df["AC"] = df["AC"].astype(str).str.replace(",", ".").str.extract(r'([\d.]+)').astype(float)
    except Exception as e:
        print(f"Erro ao converter coluna AC para número: {e}")

# === REMOVER ERROS PADRÃO DO EXCEL/GOOGLE SHEETS ===
df.replace(
    to_replace=['#N/A', '#DIV/0!', '#REF!', '#VALUE!', '#NAME?', '#NUM!', '#NULL!'],
    value='',
    inplace=True
)

# === LIMPAR ABA DESTINO ===
aba_destino.clear()

# === COLAR DADOS NA PLANILHA DESTINO ===
set_with_dataframe(aba_destino, df, row=1, col=1, include_index=False, resize=False)
print(f'Dados importados com sucesso! {len(df)} linhas coladas.')

# === PÓS-IMPORTAÇÃO: INSERIR LINHAS DA ABA CICLO QUE NÃO ESTÃO NA CARTEIRA ===
aba_ciclo = planilha_destino.worksheet('CICLO')
dados_ciclo = aba_ciclo.get_all_values()

coluna_E = [linha[4].strip() for linha in dados_ciclo[1:] if len(linha) > 4]
coluna_C = [linha[2].strip() if len(linha) > 2 else '' for linha in dados_ciclo[1:]]
coluna_F = [linha[5].strip() if len(linha) > 5 else '' for linha in dados_ciclo[1:]]

dados_atualizados = aba_destino.get_all_values()
coluna_A_atual = set([linha[0].strip() for linha in dados_atualizados[1:] if len(linha) > 0])

linhas_a_inserir = []
for i, valor in enumerate(coluna_E):
    if valor and valor not in coluna_A_atual:
        nova_linha = [''] * max(len(dados_atualizados[0]), 17)
        nova_linha[0] = valor         # Coluna A ← E da CICLO
        nova_linha[1] = coluna_F[i]   # Coluna B ← F da CICLO
        nova_linha[7] = coluna_C[i]   # Coluna H ← C da CICLO
        linhas_a_inserir.append(nova_linha)

if linhas_a_inserir:
    linha_inicio = len(dados_atualizados) + 1
    aba_destino.append_rows(linhas_a_inserir)
    linha_fim = linha_inicio + len(linhas_a_inserir) - 1
    intervalo = f"A{linha_inicio}:Q{linha_fim}"

    yellow_fill = CellFormat(backgroundColor=Color(1.0, 1.0, 0.6))  # Amarelo claro

    try:
        format_cell_range(aba_destino, intervalo, yellow_fill)
        print(f"{len(linhas_a_inserir)} novas linhas adicionadas com base na aba CICLO e coloridas.")
    except Exception as e:
        print(f"{len(linhas_a_inserir)} novas linhas adicionadas, mas houve erro ao aplicar formatação: {e}")
else:
    print("Nenhuma nova linha da aba CICLO foi adicionada (todas já estavam presentes).")

# === ESTAMPAR DATA E HORA NA CÉLULA T2 ===
try:
    agora = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    aba_destino.update_acell("T2", f"Atualizado em: {agora}")
    print(f"Data e hora registradas em T2: {agora}")
except Exception as e:
    print(f"Erro ao registrar data e hora na célula T2: {e}")
