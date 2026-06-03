import pandas as pd
import sys
sys.stdout.reconfigure(encoding='utf-8')

# Novas entradas: (nome para correspondencia.csv, MCC/MNC correto)
# Coluna "Carrier Options" deve ter um nome descritivo e único.
# Para os 3 com MNC errado, o nome aponta para o MCC/MNC correto (6 dígitos).
NOVAS_ENTRADAS = [
    # Monaco
    ('Monaco Telecom (MC)',                      '21210'),
    # Espanha
    ('Telefonica Moviles Spain (ES)',             '21407'),
    # Rússia
    ('Tele2 (RU)',                               '25020'),
    # Ucrânia
    ('Kyivstar (UA)',                            '25503'),
    # Alemanha
    ('Telefonica Germany O2 (DE)',               '26207'),
    # Gibraltar
    ('Gibtelecom (GI)',                          '26601'),
    # Chipre
    ('Primetel (CY)',                            '28020'),
    # Ilhas Faroé
    ('HEY (FO)',                                 '28802'),
    # Macedônia
    ('ONE.VIP (MK)',                             '29403'),
    # Liechtenstein
    ('Telecom Liechtenstein (LI)',               '29505'),
    # Guadalupe / Martinica
    ('Orange Caraibe (GP)',                      '34001'),
    # Ilhas Virgens Britânicas
    ('Caribbean Cellular (VG)',                  '348570'),
    # Digicel BVI  ← MNC CORRIGIDO: 34877 → 348770
    ('Digicel BVI (VG)',                         '348770'),
    # Cuba
    ('Cubacel (CU)',                             '36801'),
    # Trinidad & Tobago
    ('Digicel Trinidad & Tobago (TT)',           '374130'),
    # Turks & Caicos
    ('Digicel Turks & Caicos (TC)',              '376360'),
    # Índia - Airtel regional (18 PLMNs)
    ('Bharti Airtel Himachal Pradesh (IN)',      '40403'),
    ('Bharti Airtel Delhi (IN)',                 '40410'),
    ('Airtel North East (IN)',                   '40416'),
    ('Bharti Airtel Kolkata (IN)',               '40431'),
    ('Bharti Airtel Chennai (IN)',               '40440'),
    ('Bharti Airtel Karnataka (IN)',             '40445'),
    ('Bharti Airtel Andhra Pradesh (IN)',        '40449'),
    ('Bharti Hexacom (IN)',                      '40470'),
    ('Bharti Airtel Maharashtra (IN)',           '40490'),
    ('Bharti Airtel Mumbai (IN)',                '40492'),
    ('Bharti Airtel Madhya Pradesh (IN)',        '40493'),
    ('Bharti Airtel Tamilnadu (IN)',             '40494'),
    ('Bharti Airtel Kerala (IN)',                '40495'),
    ('Bharti Airtel Haryana (IN)',               '40496'),
    ('Bharti Airtel Uttar Pradesh (IN)',         '40497'),
    ('Bharti Airtel Gujarat (IN)',               '40498'),
    ('Airtel West Bengal (IN)',                  '40551'),
    ('Airtel Bihar (IN)',                        '40552'),
    ('Airtel UP East (IN)',                      '40554'),
    ('Airtel Assam (IN)',                        '40556'),
    # Omã
    ('Vodafone Oman (OM)',                       '42206'),
    # Macau
    ('CTM Macau (MO)',                           '45501'),
    # Indonésia
    ('XL Axiata (ID)',                           '51011'),
    # Filipinas
    ('Globe Telecom (PH)',                       '51502'),
    # Polinésia Francesa
    ('Onati (PF)',                               '54720'),
    # Cabo Verde
    ('T+ Telecomunicacoes (CV)',                 '62502'),
    # Belize
    ('Belize Telemedia (BZ)',                    '70267'),
    ('SpeedNet Communications (BZ)',             '70269'),
    # Panamá  ← MNC CORRIGIDO: 71420 → 714020
    ('Digitales Panama (PA)',                    '714020'),
    # Colômbia
    ('Colombia Movil Tigo (CO)',                 '732111'),
    # Guiana  ← MNC CORRIGIDO: 73802 → 738002
    ('GT&T Cellink Plus (GY)',                   '738002'),
]

# Carregar correspondencia atual
df_corr = pd.read_csv('correspondencia.csv', sep=';', dtype=str)
df_corr.columns = df_corr.columns.str.strip()

# MCC/MNCs já presentes
existentes_mccs = set(
    df_corr['MCC/MNC']
    .str.replace('[', '', regex=False).str.replace(']', '', regex=False)
    .str.replace('"', '', regex=False).str.strip()
    .str.split(',').str[0].str.strip().dropna()
)
existentes_nomes = set(df_corr['Carrier Options'].str.strip())

# Filtrar duplicatas
adicionadas = []
ignoradas   = []
for nome, mcc_mnc in NOVAS_ENTRADAS:
    if mcc_mnc in existentes_mccs:
        ignoradas.append((nome, mcc_mnc, 'MCC/MNC já existe'))
    elif nome in existentes_nomes:
        ignoradas.append((nome, mcc_mnc, 'Nome já existe'))
    else:
        adicionadas.append((nome, mcc_mnc))

print(f'Entradas a adicionar: {len(adicionadas)}')
print(f'Ignoradas (já existem): {len(ignoradas)}')
if ignoradas:
    print('Ignoradas:')
    for n, m, r in ignoradas:
        print(f'  {m:8s} | {n} — {r}')

# Adicionar ao arquivo
if adicionadas:
    novas_linhas = '\n'.join(
        f'{nome};[ "{mcc_mnc}" ]' for nome, mcc_mnc in adicionadas
    )
    with open('correspondencia.csv', 'a', encoding='utf-8') as f:
        f.write('\n' + novas_linhas)

    print(f'\nAdicionadas com sucesso:')
    for nome, mcc_mnc in adicionadas:
        print(f'  {mcc_mnc:8s} | {nome}')
else:
    print('\nNenhuma entrada nova para adicionar.')
