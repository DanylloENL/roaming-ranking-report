import pandas as pd
import sys
sys.stdout.reconfigure(encoding='utf-8')

# Operadoras a validar (TADIG → MCC/MNC construído pelo OMR)
PENDENTES = {
    'PANMS': ('71420',  'Grupo Digitales Panamá'),
    'GLP01': ('34001',  'Orange Caraïbe'),
    'UKRKS': ('25503',  'Kyivstar Ucrânia'),
    'INDAT': ('40410',  'Bharti Airtel Delhi'),
    'IDNEX': ('51011',  'XL Axiata Indonésia'),
    'CYPPT': ('28020',  'Primetel Chipre'),
    'TCADC': ('376360', 'Digicel Turks & Caicos'),
    'MACCT': ('45501',  'CTM Macau'),
    'INDA3': ('40498',  'Bharti Airtel Gujarat'),
    'INDA1': ('40492',  'Bharti Airtel Mumbai'),
    'INDJH': ('40449',  'Bharti Airtel Andhra Pradesh'),
    'MCOM2': ('21210',  'Monaco Telecom 3G'),
    'INDA6': ('40497',  'Bharti Airtel Uttar Pradesh'),
    'INDA2': ('40490',  'Bharti Airtel Maharashtra'),
    'TTODL': ('374130', 'Digicel Trinidad & Tobago'),
    'INDH1': ('40470',  'Bharti Hexacom'),
    'MKDNO': ('29403',  'ONE.VIP Macedonia'),
    'VGBDC': ('34877',  'Digicel British Virgin Islands'),
    'COLCO': ('732111', 'Colombia Movil (Tigo)'),
    'INDA7': ('40495',  'Bharti Airtel Kerala'),
    'INDJB': ('40445',  'Bharti Airtel Karnataka'),
    'INDSC': ('40440',  'Bharti Airtel Chennai'),
    'GIBGT': ('26601',  'Gibtelecom'),
    'INDBL': ('40403',  'Bharti Airtel Himachal Pradesh'),
    'FRATK': ('54720',  'Onati (Polynésia FR)'),
    'IND10': ('40554',  'Airtel UP East'),
    'IND16': ('40416',  'Airtel North East'),
    'INDMT': ('40431',  'Bharti Airtel Kolkata'),
    'RUST2': ('25020',  'Tele2 Russia'),
    'CUB01': ('36801',  'Cubacel'),
    'LIEMK': ('29505',  'Telecom Liechtenstein'),
    'OMNVF': ('42206',  'Vodafone Oman'),
    'GUYGT': ('73802',  'Guyana Telephone & Telegraph'),
    'BLZ67': ('70267',  'Belize Telemedia'),
    'DEUE2': ('26207',  'Telefónica Germany (O2)'),
    'ESPTE': ('21407',  'Telefónica Moviles Spain'),
    'MKDNO': ('29403',  'ONE.VIP Macedonia'),
    'IDNEX': ('51011',  'XL Axiata Indonésia'),
    'PHLGT': ('51502',  'Globe Telecom Philippines'),
    'FROKA': ('28802',  'HEY (Ilhas Faroé)'),
    'BLZSC': ('70269',  'SpeedNet Belize'),
    'VGBCC': ('348570', 'Caribbean Cellular (Ilhas Virgens)'),
    'CPVTM': ('62502',  'T+ Telecomunicações Cabo Verde'),
    'IND12': ('40552',  'Airtel Bihar'),
    'IND15': ('40556',  'Airtel Assam'),
    'INDA8': ('40493',  'Bharti Airtel Madhya Pradesh'),
    'INDA5': ('40496',  'Bharti Airtel Haryana'),
    'INDA4': ('40494',  'Bharti Airtel Tamilnadu'),
    'IND14': ('40551',  'Airtel West Bengal'),
}

tadig_pendentes = set(PENDENTES.keys())
mccs_pendentes  = {v[0] for v in PENDENTES.values()}

# ── 1. Buscar no NLR ─────────────────────────────────────────────────────────
print('=' * 75)
print('VALIDAÇÃO VIA NLR (Network Level Report)')
print('=' * 75)
df_nlr = pd.read_excel(
    '26_Maio/Network Level Report.xls-260529134802.xls', header=0)
df_nlr.columns = df_nlr.columns.str.strip()
df_nlr['mcc_mnc_nlr'] = (df_nlr['Network']
                          .str.extract(r'\((\d{3}-\d{2,3})\)')[0]
                          .str.replace('-', '', regex=False))
df_nlr['nome_nlr'] = df_nlr['Network'].str.replace(r'\(\d{3}-\d{2,3}\)', '', regex=True).str.strip()

encontrados_nlr = {}
for tadig, (mcc_omr, nome) in PENDENTES.items():
    # Procurar pelo MCC/MNC construído pelo OMR no NLR
    match = df_nlr[df_nlr['mcc_mnc_nlr'] == mcc_omr]
    if not match.empty:
        nome_nlr = match['Network'].iloc[0]
        encontrados_nlr[tadig] = (mcc_omr, nome_nlr, 'NLR ✓ CONFIRMA')
    else:
        # Tentar variações (ex: 25503 vs 25502 para Kyivstar)
        mcc3 = mcc_omr[:3]
        variantes = df_nlr[df_nlr['mcc_mnc_nlr'].str.startswith(mcc3)]
        if not variantes.empty:
            opcoes = variantes[['mcc_mnc_nlr','Network']].drop_duplicates('mcc_mnc_nlr')
            encontrados_nlr[tadig] = (mcc_omr,
                                       ' | '.join(f"{r['mcc_mnc_nlr']}={r['Network']}"
                                                  for _, r in opcoes.iterrows()),
                                       'NLR ⚠ MCC bate, MNC diferente')
        else:
            encontrados_nlr[tadig] = (mcc_omr, '—', 'NLR ✗ não encontrado')

for tadig, (mcc_omr, info, status) in sorted(encontrados_nlr.items(),
                                               key=lambda x: x[1][0]):
    nome = PENDENTES[tadig][1]
    print(f'  {tadig:8s} | {mcc_omr:7s} | {status:35s} | {nome}')
    if '⚠' in status:
        print(f'           → Variantes NLR: {info[:100]}')

# ── 2. Buscar no Acordos ─────────────────────────────────────────────────────
print()
print('=' * 75)
print('VALIDAÇÃO VIA ACORDOS (05. ACORDOS 2026 - Mai.xlsx)')
print('=' * 75)
df_ac = pd.read_excel('26_Maio/05. ACORDOS 2026 - Mai.xlsx',
                       sheet_name=0, header=None, dtype=str)
# col 69 = MCC+MNC, col 1 = Operadora, col 0 = País
df_ac_ops = df_ac.iloc[2:][[0, 1, 69]].copy()
df_ac_ops.columns = ['Pais', 'Operadora', 'mcc_mnc_ac']
df_ac_ops = df_ac_ops[df_ac_ops['Operadora'].notna()].copy()
df_ac_ops['mcc_mnc_ac'] = (df_ac_ops['mcc_mnc_ac']
                            .astype(str).str.replace(r'\.0$', '', regex=True).str.strip())
df_ac_ops = df_ac_ops[df_ac_ops['mcc_mnc_ac'].str.match(r'^\d{5,6}$', na=False)]

encontrados_ac = {}
for tadig, (mcc_omr, nome) in PENDENTES.items():
    match = df_ac_ops[df_ac_ops['mcc_mnc_ac'] == mcc_omr]
    if not match.empty:
        row = match.iloc[0]
        encontrados_ac[tadig] = (mcc_omr, f"{row['Pais']} / {row['Operadora']}", 'Acordos ✓')
    else:
        mcc3 = mcc_omr[:3]
        variantes = df_ac_ops[df_ac_ops['mcc_mnc_ac'].str.startswith(mcc3)]
        if not variantes.empty:
            opcoes = ' | '.join(f"{r['mcc_mnc_ac']}={r['Operadora']}"
                                for _, r in variantes.iterrows())
            encontrados_ac[tadig] = (mcc_omr, opcoes[:100], 'Acordos ⚠ MCC bate, MNC diferente')
        else:
            encontrados_ac[tadig] = (mcc_omr, '—', 'Acordos ✗ não encontrado')

for tadig, (mcc_omr, info, status) in sorted(encontrados_ac.items(),
                                               key=lambda x: x[1][0]):
    nome = PENDENTES[tadig][1]
    print(f'  {tadig:8s} | {mcc_omr:7s} | {status:35s} | {nome}')
    if '⚠' in status:
        print(f'           → Variantes Acordos: {info[:110]}')

# ── 3. Resumo consolidado ────────────────────────────────────────────────────
print()
print('=' * 75)
print('RESUMO — SITUAÇÃO DE CADA OPERADORA')
print('=' * 75)
print(f'  {"TADIG":8s} | {"MCC/MNC OMR":11s} | {"NLR":20s} | {"Acordos":20s} | Nome')
print('  ' + '-' * 100)
for tadig in sorted(PENDENTES, key=lambda t: PENDENTES[t][0]):
    mcc_omr, nome = PENDENTES[tadig]
    nlr_ok  = '✓' if 'CONFIRMA' in encontrados_nlr[tadig][2]  else ('⚠' if '⚠' in encontrados_nlr[tadig][2] else '✗')
    ac_ok   = '✓' if 'Acordos ✓' in encontrados_ac[tadig][2]  else ('⚠' if '⚠' in encontrados_ac[tadig][2]  else '✗')
    print(f'  {tadig:8s} | {mcc_omr:11s} | NLR:{nlr_ok:3s}               | Acordos:{ac_ok:3s}         | {nome}')
