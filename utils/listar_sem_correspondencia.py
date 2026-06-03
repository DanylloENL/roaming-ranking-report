import pandas as pd
import sys
sys.stdout.reconfigure(encoding='utf-8')

# Outbound Trafego — tem TADIG + nome + MCC/MNC
df_out = pd.read_excel(
    '26_Maio/GSM Roaming Outbound - Usuarios Unicos  e Trafego por mes - OMR.xlsx',
    header=2
)
df_out.columns = df_out.columns.str.strip()
df_out['Their MCC'] = df_out['Their MCC'].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
df_out['Their MNC'] = df_out['Their MNC'].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)

# Carregar correspondencia antes para validar MCC/MNC
_df_corr_ref = pd.read_csv('correspondencia.csv', sep=';', dtype=str)
_corr_mccs = set(
    _df_corr_ref['MCC/MNC']
    .str.replace('[', '', regex=False).str.replace(']', '', regex=False)
    .str.replace('"', '', regex=False).str.strip()
    .str.split(',').str[0].str.strip().dropna()
)

def construir_mcc_mnc(mcc, mnc):
    c2 = mcc + mnc.zfill(2)
    c3 = mcc + mnc.zfill(3)
    if c2 in _corr_mccs:
        return c2
    if c3 in _corr_mccs:
        return c3
    return c2

df_out['mcc_mnc'] = df_out.apply(lambda r: construir_mcc_mnc(r['Their MCC'], r['Their MNC']), axis=1)

for col in ['Number of Unique Roamers', 'Total Charged Volume (MB)']:
    df_out[col] = pd.to_numeric(df_out[col], errors='coerce').fillna(0)

corr_mccs = _corr_mccs  # já carregado acima com a mesma lógica

# Filtrar os sem correspondência
df_sem = df_out[~df_out['mcc_mnc'].isin(corr_mccs)].copy()

agg = df_sem.groupby(
    ['Their Original PMN (TADIG) Code', 'Their Operator Name', 'mcc_mnc']
).agg(
    Clientes   = ('Number of Unique Roamers',    'sum'),
    Trafego_MB = ('Total Charged Volume (MB)',    'sum'),
).reset_index().sort_values('Trafego_MB', ascending=False)

agg['Trafego_GB'] = (agg['Trafego_MB'] / 1000).round(1)
agg = agg[['Their Original PMN (TADIG) Code', 'Their Operator Name',
           'mcc_mnc', 'Clientes', 'Trafego_GB']]
agg.columns = ['TADIG', 'Nome (OMR)', 'MCC/MNC', 'Clientes OUT', 'Tráfego OUT (GB)']

print(f'Total sem correspondencia.csv: {len(agg)} operadoras\n')
pd.set_option('display.float_format', '{:,.1f}'.format)
pd.set_option('display.max_colwidth', 45)
pd.set_option('display.width', 130)
print(agg.to_string(index=False))
