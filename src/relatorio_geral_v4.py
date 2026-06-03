import pandas as pd
import numpy as np
import xlsxwriter
import datetime
import os
import sys
sys.stdout.reconfigure(encoding='utf-8', errors='replace')

import sys as _sys
_sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from processar_meta import processar_meta_dados

# ==============================================================================
# CONFIGURAÇÃO — ajuste os caminhos conforme o mês/ano do relatório
# ==============================================================================
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PASTA    = os.path.join(BASE_DIR, 'data')

meses_pt = {
    'january': 'janeiro', 'february': 'fevereiro', 'march': 'marco',
    'april': 'abril',     'may': 'maio',           'june': 'junho',
    'july': 'julho',      'august': 'agosto',      'september': 'setembro',
    'october': 'outubro', 'november': 'novembro',  'december': 'dezembro'
}
mes_dados   = meses_pt.get(datetime.datetime.now().strftime('%B').lower(), 'mes')
OUTPUT_XLSX = os.path.join(BASE_DIR, 'output', f'relatorio_geral_roaming_{mes_dados}.xlsx')

# Nomes dos arquivos de entrada em data/
# Renomeie os arquivos do mês para estes nomes antes de executar
ARQUIVO_ACORDOS  = os.path.join(PASTA, 'ACORDOS.xlsx')
ARQUIVO_OUT_TRAF = os.path.join(PASTA, 'GSM Roaming Outbound - Trafego OMR.xlsx')
ARQUIVO_IN_TRAF  = os.path.join(PASTA, 'GSM Roaming Inbound - Trafego OMR.xlsx')
ARQUIVO_RECEITA  = os.path.join(PASTA, 'Receita e Custo - Tarifas OMR.xlsx')


# ==============================================================================
# FUNÇÕES AUXILIARES
# ==============================================================================

def carregar_excel_com_header(caminho, dtype=None):
    """Lê Excel detectando a linha de cabeçalho real (pula títulos e linhas em branco)."""
    df_bruto = pd.read_excel(caminho, header=None, dtype=str)
    for i, row in df_bruto.iterrows():
        if row.dropna().shape[0] >= 3:
            linha_header = i
            break
    else:
        raise ValueError(f'Cabeçalho não encontrado em: {caminho}')
    df = pd.read_excel(caminho, header=linha_header, dtype=dtype)
    df.dropna(how='all', inplace=True)
    df = df[df.iloc[:, 0].notna()].reset_index(drop=True)
    df.columns = df.columns.str.strip()
    print(f"   '{os.path.basename(caminho)}': {len(df)} linhas (header na linha {linha_header + 1}).")
    return df


def sim_nao(val):
    """Converte valores de acordo (1, 'Sim', 'X') em 'Sim'/'Não'."""
    s = str(val).strip().lower()
    if s in ('1', 'sim', '1.0'):
        return 'Sim'
    if s in ('x', 'nan', '', 'none', 'não', 'nao'):
        return 'Não'
    return str(val).strip()


# ==============================================================================
# PARTE 1: DADOS META (Qualidade de rede)
# ==============================================================================
print('=' * 70)
print('PARTE 1 — DADOS META (Qualidade)')
print('=' * 70)
df_meta = processar_meta_dados()


# ==============================================================================
# PARTE 2: ACORDOS COMERCIAIS
# ==============================================================================
print()
print('=' * 70)
print('PARTE 2 — ACORDOS COMERCIAIS')
print('=' * 70)

df_ac_raw = pd.read_excel(ARQUIVO_ACORDOS, sheet_name=0, header=None, dtype=str)
# Linha 0 = grupo-headers (VOZ, GPRS, LTE, VoLTE, 5G NSA, etc.)
# Linha 1 = sub-headers (País, Operadora, in, out, ...)
# Linha 2+ = dados
# Seleção por posição (colunas definidas na inspeção):
#   0=País  1=Operadora  9=LTE_in  10=LTE_out  11=VoLTE_in  12=VoLTE_out
#   13=5GNSA_in  14=5GNSA_out  69=MCC+MNC  70=SoR  87=Passaporte
COLS_AC = {
    'Pais': 0, 'Operadora': 1,
    'Acordo_4G_OUT': 10, 'Acordo_VoLTE_OUT': 12, 'Acordo_5G_OUT': 14,
    'mcc_mnc_ac': 69, 'Steering': 70, 'Passaporte': 87
}

df_ac = df_ac_raw.iloc[2:].copy()   # pular as 2 linhas de header
df_ac = df_ac.rename(columns={v: k for k, v in COLS_AC.items()})
df_ac = df_ac[list(COLS_AC.keys())].copy()

# Manter apenas linhas de operadoras (col Operadora não nula)
df_ac = df_ac[df_ac['Operadora'].notna() & (df_ac['Operadora'].str.strip() != '')].copy()

# Limpar MCC+MNC (pode vir como float "41201.0")
df_ac['mcc_mnc_ac'] = (df_ac['mcc_mnc_ac']
                        .astype(str)
                        .str.replace(r'\.0$', '', regex=True)
                        .str.strip())
df_ac = df_ac[df_ac['mcc_mnc_ac'].str.match(r'^\d{5,6}$', na=False)].copy()

# Converter acordos para Sim/Não
for col in ['Acordo_4G_OUT', 'Acordo_VoLTE_OUT', 'Acordo_5G_OUT']:
    df_ac[col] = df_ac[col].apply(sim_nao)

# Limpar Steering (manter valor numérico ou texto)
df_ac['Steering'] = df_ac['Steering'].astype(str).str.strip().replace({'nan': '', 'None': ''})

# Deduplicar por MCC/MNC (manter a primeira ocorrência)
df_ac = df_ac.drop_duplicates(subset='mcc_mnc_ac', keep='first').reset_index(drop=True)
# Renomear para evitar colisão com Operadora_Nome do Tráfego
df_ac = df_ac.rename(columns={'Operadora': 'Operadora_Acordos', 'Pais': 'Pais_Acordos'})
print(f'   Acordos: {len(df_ac)} operadoras únicas.')


# ==============================================================================
# PARTE 3: TRÁFEGO OUTBOUND (Clientes OUT, Volume OUT + bridge TADIG→MCC/MNC)
# ==============================================================================
print()
print('=' * 70)
print('PARTE 3 — TRÁFEGO OUTBOUND')
print('=' * 70)

df_out_traf = carregar_excel_com_header(ARQUIVO_OUT_TRAF)

# Carregar correspondencia para validar MCC/MNC
_df_corr_ref = pd.read_csv('correspondencia.csv', sep=';', dtype=str)
_corr_mccs = set(
    _df_corr_ref['MCC/MNC']
    .str.replace('[', '', regex=False).str.replace(']', '', regex=False)
    .str.replace('"', '', regex=False).str.strip()
    .str.split(',').str[0].str.strip().dropna()
)

def construir_mcc_mnc(mcc, mnc):
    """
    Constrói MCC+MNC testando 2 e 3 dígitos no MNC.
    Alguns operadores (ex: Telcel MX 334-020) têm MNC de 3 dígitos que
    o Excel armazena como inteiro (20), perdendo o zero à esquerda.
    Testa ambas as versões e retorna a que existir na correspondencia.csv.
    """
    c2 = mcc + mnc.zfill(2)
    c3 = mcc + mnc.zfill(3)
    if c2 in _corr_mccs:
        return c2
    if c3 in _corr_mccs:
        return c3
    return c2  # fallback: manter 2 dígitos se nenhum bater

df_out_traf['Their MCC'] = df_out_traf['Their MCC'].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
df_out_traf['Their MNC'] = df_out_traf['Their MNC'].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
df_out_traf['mcc_mnc'] = df_out_traf.apply(
    lambda r: construir_mcc_mnc(r['Their MCC'], r['Their MNC']), axis=1
)

# Diagnóstico: quantos foram corrigidos de 2 → 3 dígitos
n_corrigidos = (df_out_traf['mcc_mnc'].str.len() == 6).sum()
print(f'   MCC/MNC com 3 dígitos no MNC (corrigidos): {n_corrigidos}')

# Tabela TADIG → mcc_mnc (bridge para Inbound e Receita)
tadig_bridge = (df_out_traf[['Their Original PMN (TADIG) Code', 'mcc_mnc', 'Their Operator Name']]
                .dropna(subset=['Their Original PMN (TADIG) Code', 'mcc_mnc'])
                .drop_duplicates('Their Original PMN (TADIG) Code')
                .rename(columns={'Their Original PMN (TADIG) Code': 'tadig'}))

# Converter colunas numéricas
for col in ['Number of Unique Roamers', 'Total Charged Volume (MB)']:
    df_out_traf[col] = pd.to_numeric(df_out_traf[col], errors='coerce').fillna(0)

# Agregar por mcc_mnc
df_out_agg = df_out_traf.groupby('mcc_mnc').agg(
    Clientes_OUT    = ('Number of Unique Roamers', 'sum'),
    Trafego_OUT_MB  = ('Total Charged Volume (MB)', 'sum'),
    Operadora_Nome  = ('Their Operator Name', 'first'),
).reset_index()
df_out_agg['Trafego_OUT_GB'] = df_out_agg['Trafego_OUT_MB'] / 1000
print(f'   Outbound Trafego: {len(df_out_agg)} operadoras.')


# ==============================================================================
# PARTE 4: TRÁFEGO INBOUND (Clientes IN, Volume IN)
# ==============================================================================
print()
print('=' * 70)
print('PARTE 4 — TRÁFEGO INBOUND')
print('=' * 70)

df_in_traf = carregar_excel_com_header(ARQUIVO_IN_TRAF)

for col in ['Number of Unique Roamers', 'Total Charged Volume (MB)']:
    df_in_traf[col] = pd.to_numeric(df_in_traf[col], errors='coerce').fillna(0)

# Mapear TADIG → mcc_mnc via bridge do Outbound
df_in_traf = df_in_traf.merge(
    tadig_bridge[['tadig', 'mcc_mnc']].rename(columns={'tadig': 'Their PMN (TADIG) Code'}),
    on='Their PMN (TADIG) Code', how='left'
)
sem_bridge = df_in_traf['mcc_mnc'].isna().sum()
print(f'   Inbound sem mapeamento MCC/MNC: {sem_bridge} linhas')

df_in_agg = df_in_traf.dropna(subset=['mcc_mnc']).groupby('mcc_mnc').agg(
    Clientes_IN   = ('Number of Unique Roamers', 'sum'),
    Trafego_IN_MB = ('Total Charged Volume (MB)', 'sum'),
).reset_index()
df_in_agg['Trafego_IN_GB'] = df_in_agg['Trafego_IN_MB'] / 1000
print(f'   Inbound Trafego: {len(df_in_agg)} operadoras mapeadas.')


# ==============================================================================
# PARTE 5: RECEITA E CUSTO
# ==============================================================================
print()
print('=' * 70)
print('PARTE 5 — RECEITA E CUSTO')
print('=' * 70)

df_rec = carregar_excel_com_header(ARQUIVO_RECEITA)

for col in ['Inbound Total Charges (Local Currency)', 'Outbound Total Charges (Local Currency)',
            'Inbound Total Charged Volume (MB)',       'Outbound Total Charged Volume (MB)']:
    df_rec[col] = pd.to_numeric(df_rec[col], errors='coerce').fillna(0)

# Remover linhas do próprio Brasil
mask_br = df_rec.apply(lambda c: c.astype(str).str.upper().str.strip().eq('BRAZIL')).any(axis=1)
df_rec = df_rec[~mask_br].copy()

# Mapear TADIG → mcc_mnc
df_rec = df_rec.merge(
    tadig_bridge[['tadig', 'mcc_mnc']].rename(columns={'tadig': 'Their PMN (TADIG) Code'}),
    on='Their PMN (TADIG) Code', how='left'
)

df_rec_agg = df_rec.dropna(subset=['mcc_mnc']).groupby('mcc_mnc').agg(
    Receita_IN      = ('Inbound Total Charges (Local Currency)',   'sum'),
    Custo_OUT       = ('Outbound Total Charges (Local Currency)',  'sum'),
    Vol_IN_MB       = ('Inbound Total Charged Volume (MB)',        'sum'),
    Vol_OUT_MB      = ('Outbound Total Charged Volume (MB)',       'sum'),
).reset_index()

df_rec_agg['Receita_Mega_IN'] = np.where(df_rec_agg['Vol_IN_MB']  > 0,
                                          df_rec_agg['Receita_IN']  / df_rec_agg['Vol_IN_MB'],  0)
df_rec_agg['Custo_Mega_OUT']  = np.where(df_rec_agg['Vol_OUT_MB'] > 0,
                                          df_rec_agg['Custo_OUT']   / df_rec_agg['Vol_OUT_MB'], 0)
print(f'   Receita/Custo: {len(df_rec_agg)} operadoras.')


# ==============================================================================
# PARTE 6: JOIN GERAL (META é a base — tráfego guia o ranking)
# ==============================================================================
print()
print('=' * 70)
print('PARTE 6 — JOIN GERAL')
print('=' * 70)

# Base: Outbound Tráfego (tem mais operadoras e guia o ranking de tráfego)
df_base = df_out_agg.copy()

df_base = df_base.merge(df_in_agg,   on='mcc_mnc', how='left')
df_base = df_base.merge(df_rec_agg[['mcc_mnc','Receita_IN','Receita_Mega_IN','Custo_OUT','Custo_Mega_OUT']],
                         on='mcc_mnc', how='left')
df_base = df_base.merge(df_ac.rename(columns={'mcc_mnc_ac': 'mcc_mnc'}),
                         on='mcc_mnc', how='left')
df_base = df_base.merge(
    df_meta[['mcc_mnc', 'Rede',
             'Claro_Speed_Mediana_OUT', 'TIM_Speed_Mediana_OUT', 'Vivo_Speed_Mediana_OUT',
             'Claro_RTT_OUT',
             'Claro_Speed_Mediana_IN',  'TIM_Speed_Mediana_IN',  'Vivo_Speed_Mediana_IN',
             'Claro_RTT_IN',
             'Natural_Market_Share', 'Registro_1a_Tentativa', 'Redir_Sem_Cobertura',
             'Nota_Speed_OUT', 'Nota_RTT_IN', 'Nota_NMS', 'Nota_Registro', 'Nota_Final',
             'Posicao_Qualidade']],
    on='mcc_mnc', how='left'
)

print(f'   Base após joins: {len(df_base)} operadoras.')
print(f'   Com Nota Final:  {df_base["Nota_Final"].notna().sum()}')
print(f'   Com Acordo:      {df_base["Pais_Acordos"].notna().sum()}')
print(f'   Com Inbound:     {df_base["Clientes_IN"].notna().sum()}')
print(f'   Com Receita:     {df_base["Receita_IN"].notna().sum()}')


# ==============================================================================
# PARTE 7: MÉTRICAS DERIVADAS
# ==============================================================================
print()
print('=' * 70)
print('PARTE 7 — MÉTRICAS DERIVADAS')
print('=' * 70)

# País final: prioriza o campo do Acordos, cai para o nome do OMR
df_base['Pais_Final'] = df_base['Pais_Acordos'].fillna(df_base['Operadora_Nome'].apply(
    lambda x: str(x).split(' ')[0] if pd.notna(x) else ''))

# Operadora final: prioriza Acordos, cai para nome do OMR
df_base['Operadora_Final'] = df_base['Operadora_Acordos'].fillna(df_base['Operadora_Nome'])

# Filtrar só quem tem tráfego Outbound positivo
df_base = df_base[df_base['Trafego_OUT_MB'] > 0].copy()

# Participação da operadora no País (OUT e IN)
df_base['Total_Pais_OUT'] = df_base.groupby('Pais_Final')['Trafego_OUT_MB'].transform('sum')
df_base['Total_Pais_IN']  = df_base.groupby('Pais_Final')['Trafego_IN_MB'].transform('sum')

df_base['Perc_Pais_OUT'] = np.where(df_base['Total_Pais_OUT'] > 0,
                                     df_base['Trafego_OUT_MB'] / df_base['Total_Pais_OUT'], 0)
df_base['Perc_Pais_IN']  = np.where(df_base['Total_Pais_IN']  > 0,
                                     df_base['Trafego_IN_MB']  / df_base['Total_Pais_IN'],  0)

# Ranking por tráfego Outbound
df_base = df_base.sort_values('Trafego_OUT_MB', ascending=False).reset_index(drop=True)
df_base['Posicao_Trafego'] = df_base.index + 1

# Posição qualidade (já vem do META, mas re-ranquear sobre a base de tráfego filtrada)
df_base['Posicao_Qualidade'] = (df_base['Nota_Final']
                                 .rank(method='min', ascending=False, na_option='bottom')
                                 .astype('Int64'))

print(f'   Total após filtro Tráfego > 0: {len(df_base)} operadoras')


# ==============================================================================
# PARTE 8: MONTAR DATAFRAME FINAL
# ==============================================================================
df_final = df_base.rename(columns={
    'Posicao_Qualidade':    'Posição Qualidade',
    'Nota_Final':           'Nota',
    'Posicao_Trafego':      'Posição Tráfego',
    'Pais_Final':           'País',
    'Operadora_Final':      'Operadora',
    'Trafego_OUT_GB':       'Tráfego OUT (GB)',
    'Trafego_IN_GB':        'Tráfego IN (GB)',
    'Perc_Pais_OUT':        '% Participação OUT no País',
    'Perc_Pais_IN':         '% Participação IN no País',
    'Clientes_OUT':         'Clientes OUT (mês)',
    'Clientes_IN':          'Clientes IN (mês)',
    'Passaporte':           'Passaporte',
    'Acordo_4G_OUT':        '4G OUT',
    'Acordo_VoLTE_OUT':     'VoLTE OUT',
    'Acordo_5G_OUT':        '5G NSA OUT',
    'Steering':             'Steering',
    'Receita_IN':           'Receita Mensal IN',
    'Receita_Mega_IN':      'Receita/Mega IN',
    'Custo_OUT':            'Custo Mensal OUT',
    'Custo_Mega_OUT':       'Custo/Mega OUT',
    'Claro_Speed_Mediana_OUT': 'Speed Mediana OUT Claro (Mbps)',
    'TIM_Speed_Mediana_OUT':   'Speed Mediana OUT TIM (Mbps)',
    'Vivo_Speed_Mediana_OUT':  'Speed Mediana OUT Vivo (Mbps)',
    'Claro_Speed_Mediana_IN':  'Speed Mediana IN Claro (Mbps)',
    'TIM_Speed_Mediana_IN':    'Speed Mediana IN TIM (Mbps)',
    'Vivo_Speed_Mediana_IN':   'Speed Mediana IN Vivo (Mbps)',
    'Claro_RTT_OUT':           'RTT OUT Claro (ms)',
    'Claro_RTT_IN':            'RTT IN Claro (ms)',
    'Natural_Market_Share':    'Cobertura/NMS',
    'Registro_1a_Tentativa':   'Registro 1ª Tentativa',
    'Redir_Sem_Cobertura':     'Redir. Sem Cobertura',
})

ORDEM_COLUNAS = [
    'Posição Qualidade', 'Nota', 'Posição Tráfego', 'País', 'Operadora',
    'Tráfego OUT (GB)', 'Tráfego IN (GB)',
    '% Participação OUT no País', '% Participação IN no País',
    'Clientes OUT (mês)', 'Clientes IN (mês)',
    'Passaporte', '4G OUT', 'VoLTE OUT', '5G NSA OUT', 'Steering',
    'Receita Mensal IN', 'Receita/Mega IN',
    'Custo Mensal OUT', 'Custo/Mega OUT',
    'Speed Mediana OUT Claro (Mbps)', 'Speed Mediana OUT TIM (Mbps)', 'Speed Mediana OUT Vivo (Mbps)',
    'Speed Mediana IN Claro (Mbps)',  'Speed Mediana IN TIM (Mbps)',  'Speed Mediana IN Vivo (Mbps)',
    'RTT OUT Claro (ms)', 'RTT IN Claro (ms)',
    'Cobertura/NMS', 'Registro 1ª Tentativa', 'Redir. Sem Cobertura',
]
cols_existentes = [c for c in ORDEM_COLUNAS if c in df_final.columns]
df_final = df_final[cols_existentes].copy()

# Tratar nulos em colunas de texto
for col in ['Passaporte', '4G OUT', 'VoLTE OUT', '5G NSA OUT', 'Steering']:
    if col in df_final.columns:
        df_final[col] = df_final[col].fillna('').astype(str).str.strip().replace({'nan': '', 'None': ''})


# ==============================================================================
# PARTE 9: GERAR EXCEL
# ==============================================================================
print()
print('=' * 70)
print('PARTE 9 — GERANDO EXCEL')
print('=' * 70)

writer   = pd.ExcelWriter(OUTPUT_XLSX, engine='xlsxwriter')
wb       = writer.book

# ── Formatos ─────────────────────────────────────────────────────────────────
FMT_HDR_ROXO   = wb.add_format({'bg_color': '#D86DCD', 'font_color': 'white', 'bold': True,
                                  'align': 'center', 'valign': 'vcenter', 'border': 1})
FMT_HDR_AZUL   = wb.add_format({'bg_color': '#1E6686', 'font_color': 'white', 'bold': True,
                                  'align': 'center', 'valign': 'vcenter', 'border': 1})
FMT_HDR_VERDE  = wb.add_format({'bg_color': '#397E3F', 'font_color': 'white', 'bold': True,
                                  'align': 'center', 'valign': 'vcenter', 'border': 1})
FMT_HDR_LARANJA= wb.add_format({'bg_color': '#E47C4D', 'font_color': 'white', 'bold': True,
                                  'align': 'center', 'valign': 'vcenter', 'border': 1})
FMT_HDR_INDIGO = wb.add_format({'bg_color': '#5B5EA6', 'font_color': 'white', 'bold': True,
                                  'align': 'center', 'valign': 'vcenter', 'border': 1})

FMT_INT      = wb.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter', 'num_format': '0'})
FMT_FLOAT2   = wb.add_format({'border': 1, 'align': 'right',  'valign': 'vcenter', 'num_format': '0.00'})
FMT_FLOAT4   = wb.add_format({'border': 1, 'align': 'right',  'valign': 'vcenter', 'num_format': '0.0000'})
FMT_TRAFEGO  = wb.add_format({'border': 1, 'align': 'right',  'valign': 'vcenter', 'num_format': '#,##0.00'})
FMT_CLIENTES = wb.add_format({'border': 1, 'align': 'right',  'valign': 'vcenter', 'num_format': '#,##0'})
FMT_PERCENT  = wb.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter', 'num_format': '0.00%'})
FMT_MOEDA    = wb.add_format({'border': 1, 'align': 'right',  'valign': 'vcenter', 'num_format': 'U$ #,##0.00'})
FMT_MEGACOST = wb.add_format({'border': 1, 'align': 'right',  'valign': 'vcenter', 'num_format': 'U$ 0.000000'})
FMT_TEXT     = wb.add_format({'border': 1, 'align': 'left',   'valign': 'vcenter', 'indent': 1})
FMT_CENTER   = wb.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter'})

# Formato condicional da coluna Nota
FMT_NOTA_CINZA    = wb.add_format({'bg_color': '#808080', 'font_color': '#000000', 'border': 1,
                                    'align': 'center', 'valign': 'vcenter', 'num_format': '0.0000', 'bold': True})
FMT_NOTA_AMARELO  = wb.add_format({'bg_color': '#FFF98F', 'font_color': '#000000', 'border': 1,
                                    'align': 'center', 'valign': 'vcenter', 'num_format': '0.0000', 'bold': True})
FMT_NOTA_VERDE    = wb.add_format({'bg_color': '#7AFF96', 'font_color': '#000000', 'border': 1,
                                    'align': 'center', 'valign': 'vcenter', 'num_format': '0.0000', 'bold': True})
FMT_NOTA_AZUL     = wb.add_format({'bg_color': '#66B3FF', 'font_color': '#000000', 'border': 1,
                                    'align': 'center', 'valign': 'vcenter', 'num_format': '0.0000', 'bold': True})

# Mapeamento coluna → formato
COL_FMTS = {
    'Posição Qualidade':              FMT_INT,
    'Nota':                           FMT_FLOAT4,
    'Posição Tráfego':                FMT_INT,
    'País':                           FMT_TEXT,
    'Operadora':                      FMT_TEXT,
    'Tráfego OUT (GB)':               FMT_TRAFEGO,
    'Tráfego IN (GB)':                FMT_TRAFEGO,
    '% Participação OUT no País':     FMT_PERCENT,
    '% Participação IN no País':      FMT_PERCENT,
    'Clientes OUT (mês)':             FMT_CLIENTES,
    'Clientes IN (mês)':              FMT_CLIENTES,
    'Passaporte':                     FMT_CENTER,
    '4G OUT':                         FMT_CENTER,
    'VoLTE OUT':                      FMT_CENTER,
    '5G NSA OUT':                     FMT_CENTER,
    'Steering':                       FMT_CENTER,
    'Receita Mensal IN':              FMT_MOEDA,
    'Receita/Mega IN':                FMT_MEGACOST,
    'Custo Mensal OUT':               FMT_MOEDA,
    'Custo/Mega OUT':                 FMT_MEGACOST,
    'Speed Mediana OUT Claro (Mbps)': FMT_FLOAT2,
    'Speed Mediana OUT TIM (Mbps)':   FMT_FLOAT2,
    'Speed Mediana OUT Vivo (Mbps)':  FMT_FLOAT2,
    'Speed Mediana IN Claro (Mbps)':  FMT_FLOAT2,
    'Speed Mediana IN TIM (Mbps)':    FMT_FLOAT2,
    'Speed Mediana IN Vivo (Mbps)':   FMT_FLOAT2,
    'RTT OUT Claro (ms)':             FMT_FLOAT2,
    'RTT IN Claro (ms)':              FMT_FLOAT2,
    'Cobertura/NMS':                  FMT_PERCENT,
    'Registro 1ª Tentativa':          FMT_PERCENT,
    'Redir. Sem Cobertura':           FMT_PERCENT,
}

# Mapeamento coluna → grupo (para cores do header)
# Roxo=qualidade  Azul=info geral  Verde=acordos  Laranja=financeiro  Índigo=KPIs
GRUPOS = {
    'Posição Qualidade':              'roxo',
    'Nota':                           'roxo',
    'Posição Tráfego':                'azul',
    'País':                           'azul',
    'Operadora':                      'azul',
    'Tráfego OUT (GB)':               'azul',
    'Tráfego IN (GB)':                'azul',
    '% Participação OUT no País':     'azul',
    '% Participação IN no País':      'azul',
    'Clientes OUT (mês)':             'azul',
    'Clientes IN (mês)':              'azul',
    'Passaporte':                     'verde',
    '4G OUT':                         'verde',
    'VoLTE OUT':                      'verde',
    '5G NSA OUT':                     'verde',
    'Steering':                       'verde',
    'Receita Mensal IN':              'laranja',
    'Receita/Mega IN':                'laranja',
    'Custo Mensal OUT':               'laranja',
    'Custo/Mega OUT':                 'laranja',
    'Speed Mediana OUT Claro (Mbps)': 'indigo',
    'Speed Mediana OUT TIM (Mbps)':   'indigo',
    'Speed Mediana OUT Vivo (Mbps)':  'indigo',
    'Speed Mediana IN Claro (Mbps)':  'indigo',
    'Speed Mediana IN TIM (Mbps)':    'indigo',
    'Speed Mediana IN Vivo (Mbps)':   'indigo',
    'RTT OUT Claro (ms)':             'indigo',
    'RTT IN Claro (ms)':              'indigo',
    'Cobertura/NMS':                  'indigo',
    'Registro 1ª Tentativa':          'indigo',
    'Redir. Sem Cobertura':           'indigo',
}
HDR_FMT_MAP = {
    'roxo':    FMT_HDR_ROXO,
    'azul':    FMT_HDR_AZUL,
    'verde':   FMT_HDR_VERDE,
    'laranja':  FMT_HDR_LARANJA,
    'indigo':  FMT_HDR_INDIGO,
}


def escrever_aba(writer, df, nome_aba, wb):
    """Escreve o DataFrame em uma aba com headers coloridos e formatação."""
    df.to_excel(writer, sheet_name=nome_aba, startrow=2, index=False)
    ws = writer.sheets[nome_aba]

    # Headers coloridos (linha 3 no Excel = índice 2)
    for col_num, col_name in enumerate(df.columns):
        grupo = GRUPOS.get(col_name, 'azul')
        ws.write(2, col_num, col_name, HDR_FMT_MAP[grupo])

    # Largura e formato das colunas
    LARGURAS = {
        'Posição Qualidade': 10, 'Nota': 12, 'Posição Tráfego': 10,
        'País': 28, 'Operadora': 40,
        'Tráfego OUT (GB)': 18, 'Tráfego IN (GB)': 18,
        '% Participação OUT no País': 22, '% Participação IN no País': 22,
        'Clientes OUT (mês)': 16, 'Clientes IN (mês)': 16,
        'Passaporte': 15, '4G OUT': 10, 'VoLTE OUT': 10, '5G NSA OUT': 10, 'Steering': 12,
        'Receita Mensal IN': 18, 'Receita/Mega IN': 18,
        'Custo Mensal OUT': 18, 'Custo/Mega OUT': 18,
    }
    for col_num, col_name in enumerate(df.columns):
        largura = LARGURAS.get(col_name, 22)
        fmt = COL_FMTS.get(col_name, FMT_CENTER)
        ws.set_column(col_num, col_num, largura, fmt)

    # Merge de grupos no header (linhas 1-2 = índice 0-1)
    cols = list(df.columns)
    grupos_seq = []
    cur_grupo = GRUPOS.get(cols[0], 'azul')
    ini = 0
    for i, c in enumerate(cols[1:], 1):
        g = GRUPOS.get(c, 'azul')
        if g != cur_grupo:
            grupos_seq.append((ini, i - 1, cur_grupo))
            ini, cur_grupo = i, g
    grupos_seq.append((ini, len(cols) - 1, cur_grupo))

    NOMES_GRUPO = {
        'roxo': 'RANKING QUALIDADE', 'azul': 'INFORMAÇÕES',
        'verde': 'ACORDOS COMERCIAIS', 'laranja': 'FINANCEIRO', 'indigo': 'KPIs META'
    }
    for ini, fim, grupo in grupos_seq:
        label = NOMES_GRUPO.get(grupo, '')
        fmt   = HDR_FMT_MAP[grupo]
        if ini == fim:
            ws.merge_range(0, ini, 1, fim, label, fmt)
        else:
            ws.merge_range(0, ini, 1, fim, label, fmt)

    # Formatação condicional na coluna Nota (índice 1)
    last_row = 2 + len(df)
    nota_col = 1
    ws.conditional_format(3, nota_col, last_row, nota_col,
                          {'type': 'cell', 'criteria': 'between', 'minimum': 0,   'maximum': 0.4,  'format': FMT_NOTA_CINZA})
    ws.conditional_format(3, nota_col, last_row, nota_col,
                          {'type': 'cell', 'criteria': 'between', 'minimum': 0.4, 'maximum': 0.6,  'format': FMT_NOTA_AMARELO})
    ws.conditional_format(3, nota_col, last_row, nota_col,
                          {'type': 'cell', 'criteria': 'between', 'minimum': 0.6, 'maximum': 0.8,  'format': FMT_NOTA_VERDE})
    ws.conditional_format(3, nota_col, last_row, nota_col,
                          {'type': 'cell', 'criteria': 'greater than', 'value': 0.8,               'format': FMT_NOTA_AZUL})

    ws.freeze_panes(3, 0)
    print(f'   Aba "{nome_aba}" gerada com {len(df)} linhas.')


# ── Aba 1: Ordenado por Tráfego ───────────────────────────────────────────────
df_trafego = df_final.sort_values('Posição Tráfego', ascending=True).reset_index(drop=True)
df_trafego['Nota'] = df_trafego['Nota'].fillna(0)
escrever_aba(writer, df_trafego, 'Top - Tráfego', wb)

# ── Aba 2: Ordenado por Qualidade ─────────────────────────────────────────────
df_qualidade = df_final.sort_values(['Nota', 'Posição Tráfego'],
                                     ascending=[False, True]).reset_index(drop=True)
df_qualidade['Posição Qualidade'] = df_qualidade.index + 1
df_qualidade['Nota'] = df_qualidade['Nota'].fillna(0)
escrever_aba(writer, df_qualidade, 'Top - Qualidade', wb)

# ── Aba 3: Diagnóstico — operadoras sem dados de qualidade ───────────────────
print('\n--- Gerando aba de diagnóstico ---')

# Usar df_base (que ainda tem mcc_mnc) para diagnosticar os sem nota
# Operadoras sem Nota no df_base = Nota_Final nula ou zero
sem_nota_base = df_base[df_base['Nota_Final'].isna() | (df_base['Nota_Final'] == 0)].copy()
sem_nota_base = sem_nota_base.sort_values('Trafego_OUT_MB', ascending=False).reset_index(drop=True)
sem_nota_base['Posicao_Trafego'] = sem_nota_base['Trafego_OUT_MB'].rank(
    method='first', ascending=False).astype(int)

# Carregar correspondencia para verificação
df_corr = pd.read_csv('correspondencia.csv', sep=';', dtype=str)
df_corr['mcc_mnc_corr'] = (df_corr['MCC/MNC']
                            .str.replace(r'[\[\]"]', '', regex=True)
                            .str.strip().str.split(',').str[0].str.strip())

df_diag = sem_nota_base[[
    'Posicao_Trafego', 'Pais_Final', 'Operadora_Final', 'mcc_mnc',
    'Trafego_OUT_GB', 'Clientes_OUT', 'Receita_IN', 'Custo_OUT'
]].rename(columns={
    'Posicao_Trafego': 'Posição Tráfego',
    'Pais_Final':      'País',
    'Operadora_Final': 'Operadora',
    'Trafego_OUT_GB':  'Tráfego OUT (GB)',
    'Clientes_OUT':    'Clientes OUT (mês)',
    'Receita_IN':      'Receita Mensal IN',
    'Custo_OUT':       'Custo Mensal OUT',
}).copy()

# Verificar se o MCC/MNC existe na correspondencia.csv
corr_mccs = set(df_corr['mcc_mnc_corr'].dropna())
df_diag['MCC/MNC na correspondencia?'] = sem_nota_base['mcc_mnc'].apply(
    lambda x: 'Sim' if str(x) in corr_mccs else 'Não').values

# Verificar se o MCC/MNC existe no NLR
nlr_mccs = set(df_meta['mcc_mnc'].dropna())
df_diag['No NLR?'] = sem_nota_base['mcc_mnc'].apply(
    lambda x: 'Sim' if str(x) in nlr_mccs else 'Não').values

# Guardar mcc_mnc para o diagnosticar()
df_diag['mcc_mnc'] = sem_nota_base['mcc_mnc'].values

# Diagnóstico principal
def diagnosticar(row):
    mcc = str(row.get('mcc_mnc', ''))
    if row['MCC/MNC na correspondencia?'] == 'Não':
        return 'MCC/MNC ausente na correspondencia.csv → nome da rede não foi mapeado'
    if row['No NLR?'] == 'Não':
        return 'MCC/MNC não encontrado no Network Level Report'
    return 'Operadora fora do Top 200 do ranking META (não classificada pelo Opensignal)'

df_diag['Motivo sem Nota'] = df_diag.apply(diagnosticar, axis=1)
df_diag = df_diag.sort_values('Posição Tráfego').reset_index(drop=True)

# Escrever aba manualmente (formato mais simples)
df_diag.to_excel(writer, sheet_name='Sem Dados Qualidade', startrow=1, index=False)
ws3 = writer.sheets['Sem Dados Qualidade']

FMT_HDR_DIAG = wb.add_format({'bg_color': '#C0392B', 'font_color': 'white', 'bold': True,
                               'align': 'center', 'valign': 'vcenter', 'border': 1})
FMT_MOTIVO   = wb.add_format({'border': 1, 'align': 'left', 'valign': 'vcenter',
                               'bg_color': '#FCF3CF', 'indent': 1})
FMT_SIM      = wb.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter',
                               'bg_color': '#D5F5E3'})
FMT_NAO      = wb.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter',
                               'bg_color': '#FADBD8'})

# Header colorido
for col_num, col_name in enumerate(df_diag.columns):
    ws3.write(1, col_num, col_name, FMT_HDR_DIAG)

# Larguras
larguras_diag = {
    'Posição Tráfego': 12, 'País': 22, 'Operadora': 38, 'mcc_mnc': 10,
    'Tráfego OUT (GB)': 16, 'Clientes OUT (mês)': 16,
    'Receita Mensal IN': 16, 'Custo Mensal OUT': 16,
    'MCC/MNC na correspondencia?': 22, 'No NLR?': 14,
    'Motivo sem Nota': 60,
}
for col_num, col_name in enumerate(df_diag.columns):
    largura = larguras_diag.get(col_name, 16)
    ws3.set_column(col_num, col_num, largura)

# Colorir célula de Motivo e colunas Sim/Não
for row_num, row_data in df_diag.iterrows():
    excel_row = row_num + 2  # +1 header fixo + +1 startrow
    motivo_col = list(df_diag.columns).index('Motivo sem Nota')
    sim_col    = list(df_diag.columns).index('MCC/MNC na correspondencia?')
    nlr_col    = list(df_diag.columns).index('No NLR?')
    ws3.write(excel_row, motivo_col, row_data['Motivo sem Nota'], FMT_MOTIVO)
    ws3.write(excel_row, sim_col,
              row_data['MCC/MNC na correspondencia?'],
              FMT_SIM if row_data['MCC/MNC na correspondencia?'] == 'Sim' else FMT_NAO)
    ws3.write(excel_row, nlr_col,
              row_data['No NLR?'],
              FMT_SIM if row_data['No NLR?'] == 'Sim' else FMT_NAO)

ws3.freeze_panes(2, 0)

# Resumo no topo
ws3.merge_range(0, 0, 0, len(df_diag.columns) - 1,
                f'OPERADORAS SEM DADOS DE QUALIDADE — {len(df_diag)} operadoras (de {len(df_trafego)} total)',
                wb.add_format({'bg_color': '#C0392B', 'font_color': 'white', 'bold': True,
                               'align': 'center', 'valign': 'vcenter', 'font_size': 12}))

# Contagem por motivo
print(f'   Operadoras sem nota: {len(df_diag)}')
print('   Motivos:')
for motivo, cnt in df_diag['Motivo sem Nota'].value_counts().items():
    print(f'     [{cnt:3d}] {motivo}')

writer.close()
print(f'\n✅ Relatório gerado: {OUTPUT_XLSX}')
print(f'   Operadoras: {len(df_trafego)} | Colunas: {len(df_final.columns)}')
