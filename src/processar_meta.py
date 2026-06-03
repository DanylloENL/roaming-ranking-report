import pandas as pd
import numpy as np
import os
import sys
sys.stdout.reconfigure(encoding='utf-8', errors='replace')

BASE_DIR        = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PASTA           = os.path.join(BASE_DIR, 'data')
CORRESPONDENCIA = os.path.join(BASE_DIR, 'data', 'correspondencia.csv')
OUTPUT          = os.path.join(BASE_DIR, 'output', 'resultado_meta_teste.csv')

# Os CSVs de ranking META devem estar em data/ com nomes no formato abaixo.
# Ajuste as datas conforme o mês do relatório.
ARQUIVO_OUT_100 = os.path.join(PASTA, 'Outbound_TOP 100 Ranking_Compare2G-3G-4G-5G.csv')
ARQUIVO_OUT_200 = os.path.join(PASTA, 'Outbound_TOP101_200_Compare2G-3G-4G-5G.csv')
ARQUIVO_IN_100  = os.path.join(PASTA, 'Inbound_TOP 100 Ranking_Compare2G-3G-4G-5G.csv')
ARQUIVO_IN_200  = os.path.join(PASTA, 'Inbound_TOP101_200_Compare2G-3G-4G-5G.csv')
ARQUIVO_NLR     = os.path.join(PASTA, 'Network Level Report.xls')

OPERADORES_OUT = {'Claro': 'Claro (BR)', 'Vivo': 'Vivo', 'TIM': 'TIM'}
OPERADORES_IN  = {'Claro': 'Claro (BR)', 'TIM': 'TIM',  'Vivo': 'Vivo'}


# ==============================================================================
# FUNÇÕES
# ==============================================================================

def carregar_csv_meta(caminho):
    df_raw = pd.read_csv(
        caminho, header=None, skiprows=1,
        encoding='utf-8-sig', quoting=3, on_bad_lines='skip', dtype=str
    )
    header = (df_raw.iloc[0]
              .str.strip().str.strip('"')
              .str.replace('﻿', '', regex=False).tolist())
    df_data = df_raw.iloc[1:].reset_index(drop=True)
    df_data.columns = range(len(df_data.columns))
    print(f"   '{os.path.basename(caminho)}': {len(df_data)} redes.")
    return df_data, header


def find_col(header, pattern):
    for i, h in enumerate(header):
        if pattern in h:
            return i
    return None


def limpar_numerico(serie):
    return pd.to_numeric(
        serie.str.replace(' Mbps', '', regex=False)
             .str.replace(' ms', '', regex=False)
             .str.strip().str.strip('"'),
        errors='coerce'
    )


def extrair_metricas(df_data, header, operadores):
    resultado = pd.DataFrame()
    resultado['Rede'] = (df_data.iloc[:, 0]
                         .str.strip().str.strip('"')
                         .str.replace('﻿', '', regex=False))
    for alias, nome_csv in operadores.items():
        for tech in ['4G', '5G']:
            idx_speed = find_col(header, f'{nome_csv} {tech} download speed (median)')
            if idx_speed is not None:
                resultado[f'{alias}_speed_median_{tech}'] = limpar_numerico(df_data.iloc[:, idx_speed])
                resultado[f'{alias}_samples_{tech}']      = pd.to_numeric(
                    df_data.iloc[:, idx_speed + 1].str.strip().str.strip('"'), errors='coerce').fillna(0)
            idx_rtt = find_col(header, f'{nome_csv} {tech} round-trip time (median)')
            if idx_rtt is not None:
                resultado[f'{alias}_rtt_{tech}']         = limpar_numerico(df_data.iloc[:, idx_rtt])
                resultado[f'{alias}_rtt_samples_{tech}'] = pd.to_numeric(
                    df_data.iloc[:, idx_rtt + 1].str.strip().str.strip('"'), errors='coerce').fillna(0)
    return resultado


def combinar_4g_5g(df, col_4g, n_4g, col_5g, n_5g):
    for c in [col_4g, n_4g, col_5g, n_5g]:
        if c not in df.columns:
            return pd.Series(np.nan, index=df.index)
    v4 = df[col_4g].fillna(0)
    n4 = df[n_4g].fillna(0)
    v5 = df[col_5g].fillna(0)
    n5 = df[n_5g].fillna(0)
    total = n4 + n5
    combined = np.where(total > 0, (v4 * n4 + v5 * n5) / total, np.nan)
    return pd.Series(combined, index=df.index)


def processar_nlr(caminho):
    df = pd.read_excel(caminho, header=0)
    df.columns = df.columns.str.strip()
    df['mcc_mnc'] = (df['Network']
                     .str.extract(r'\((\d{3}-\d{2,3})\)')[0]
                     .str.replace('-', '', regex=False))
    df = df[df['mcc_mnc'].notna()].copy()

    TOTAL_REG = 'Total Registration Attempts (1)'
    NMS_COL   = 'Natural Market Share at Zone Entry'
    FT_ATT    = 'First Time Registration Attempts (2)'
    FT_ALLOW  = 'First Time Registration Attempts allowed (3)'
    NO_COV    = 'Registrations due to No Coverage (7)'

    for col in [TOTAL_REG, NMS_COL, FT_ATT, FT_ALLOW, NO_COV]:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    df['NMS_weighted'] = df[NMS_COL] * df[TOTAL_REG]

    agg = df.groupby('mcc_mnc').agg(
        NMS_w       = ('NMS_weighted', 'sum'),
        TotalReg    = (TOTAL_REG,      'sum'),
        FT_Attempts = (FT_ATT,         'sum'),
        FT_Allowed  = (FT_ALLOW,       'sum'),
        NoCov       = (NO_COV,         'sum'),
    ).reset_index()

    agg['Natural_Market_Share']  = np.where(
        agg['TotalReg'] > 0, agg['NMS_w'] / agg['TotalReg'] / 100, 0)
    agg['Registro_1a_Tentativa'] = np.where(
        agg['FT_Attempts'] > 0, agg['FT_Allowed'] / agg['FT_Attempts'], 0)
    agg['Redir_Sem_Cobertura']   = np.where(
        agg['FT_Attempts'] > 0, agg['NoCov'] / agg['FT_Attempts'], 0)

    print(f"   NLR: {len(agg)} operadoras únicas processadas.")
    return agg[['mcc_mnc', 'Natural_Market_Share', 'Registro_1a_Tentativa', 'Redir_Sem_Cobertura']]


def carregar_correspondencia():
    df = pd.read_csv(CORRESPONDENCIA, sep=';', dtype=str)
    df['Carrier Options'] = df['Carrier Options'].str.strip()
    df['mcc_mnc'] = (df['MCC/MNC']
                     .str.replace(r'[\[\]"]', '', regex=True)
                     .str.strip().str.split(',').str[0].str.strip())
    df['mcc'] = df['mcc_mnc'].str.slice(0, 3)
    df['mnc'] = df['mcc_mnc'].str.slice(3)
    return df[['Carrier Options', 'mcc_mnc', 'mcc', 'mnc']]


# ==============================================================================
# PIPELINE PRINCIPAL
# ==============================================================================

def processar_meta_dados():
    """
    Executa o pipeline completo META + NLR.
    Retorna DataFrame com uma linha por operador (mcc_mnc).
    """
    print("### PROCESSANDO DADOS META ###\n")

    # ── Outbound ────────────────────────────────────────────────────────
    print("--- Outbound ---")
    df_out_raw, hdr_out = carregar_csv_meta(ARQUIVO_OUT_100)
    df_out_200, _       = carregar_csv_meta(ARQUIVO_OUT_200)
    df_out_raw = pd.concat([df_out_raw, df_out_200], ignore_index=True)
    antes = len(df_out_raw)
    df_out_raw = df_out_raw.drop_duplicates(subset=[0], keep='first').reset_index(drop=True)
    print(f"   Total: {len(df_out_raw)} únicas ({antes - len(df_out_raw)} duplicatas removidas)\n")

    # ── Inbound ─────────────────────────────────────────────────────────
    print("--- Inbound ---")
    df_in_raw, hdr_in = carregar_csv_meta(ARQUIVO_IN_100)
    df_in_200, _      = carregar_csv_meta(ARQUIVO_IN_200)
    df_in_raw = pd.concat([df_in_raw, df_in_200], ignore_index=True)
    antes = len(df_in_raw)
    df_in_raw = df_in_raw.drop_duplicates(subset=[0], keep='first').reset_index(drop=True)
    print(f"   Total: {len(df_in_raw)} únicas ({antes - len(df_in_raw)} duplicatas removidas)\n")

    # ── Extração e combinação 4G+5G ──────────────────────────────────────
    print("--- Extraindo e combinando métricas ---")
    df_out = extrair_metricas(df_out_raw, hdr_out, OPERADORES_OUT)
    df_in  = extrair_metricas(df_in_raw,  hdr_in,  OPERADORES_IN)

    for alias in OPERADORES_OUT.keys():
        df_out[f'{alias}_Speed_Mediana_OUT'] = combinar_4g_5g(
            df_out,
            f'{alias}_speed_median_4G', f'{alias}_samples_4G',
            f'{alias}_speed_median_5G', f'{alias}_samples_5G')
        df_out[f'{alias}_RTT_OUT'] = combinar_4g_5g(
            df_out,
            f'{alias}_rtt_4G', f'{alias}_rtt_samples_4G',
            f'{alias}_rtt_5G', f'{alias}_rtt_samples_5G')
        df_out[f'{alias}_Samples_Total_OUT'] = (
            df_out[f'{alias}_samples_4G'].fillna(0) + df_out[f'{alias}_samples_5G'].fillna(0))

    for alias in OPERADORES_IN.keys():
        df_in[f'{alias}_Speed_Mediana_IN'] = combinar_4g_5g(
            df_in,
            f'{alias}_speed_median_4G', f'{alias}_samples_4G',
            f'{alias}_speed_median_5G', f'{alias}_samples_5G')
        df_in[f'{alias}_RTT_IN'] = combinar_4g_5g(
            df_in,
            f'{alias}_rtt_4G', f'{alias}_rtt_samples_4G',
            f'{alias}_rtt_5G', f'{alias}_rtt_samples_5G')
        df_in[f'{alias}_Samples_Total_IN'] = (
            df_in[f'{alias}_samples_4G'].fillna(0) + df_in[f'{alias}_samples_5G'].fillna(0))

    print("   ✅ Combinação concluída.")

    # ── Join MCC/MNC ─────────────────────────────────────────────────────
    print("\n--- Adicionando MCC/MNC ---")
    df_corr = carregar_correspondencia()
    df_out = df_out.merge(df_corr, left_on='Rede', right_on='Carrier Options', how='left').drop(columns=['Carrier Options'])
    df_in  = df_in.merge(df_corr,  left_on='Rede', right_on='Carrier Options', how='left').drop(columns=['Carrier Options'])
    print(f"   Outbound sem MCC/MNC: {df_out['mcc_mnc'].isna().sum()} redes")
    print(f"   Inbound  sem MCC/MNC: {df_in['mcc_mnc'].isna().sum()} redes")

    # ── Consolidar OUT + IN ──────────────────────────────────────────────
    print("\n--- Consolidando OUT + IN ---")
    cols_out = ['Rede', 'mcc_mnc', 'mcc', 'mnc',
                'Claro_Speed_Mediana_OUT', 'TIM_Speed_Mediana_OUT', 'Vivo_Speed_Mediana_OUT',
                'Claro_RTT_OUT', 'Claro_Samples_Total_OUT']
    cols_in  = ['Rede', 'mcc_mnc',
                'Claro_Speed_Mediana_IN', 'TIM_Speed_Mediana_IN', 'Vivo_Speed_Mediana_IN',
                'Claro_RTT_IN', 'Claro_Samples_Total_IN']

    df_out_sel = df_out[[c for c in cols_out if c in df_out.columns]].copy()
    df_in_sel  = df_in[[c for c in cols_in  if c in df_in.columns]].copy()

    df_resultado = df_out_sel.merge(
        df_in_sel.rename(columns={'Rede': 'Rede_IN', 'mcc_mnc': 'mcc_mnc_IN'}),
        left_on='mcc_mnc', right_on='mcc_mnc_IN', how='left'
    ).drop(columns=['mcc_mnc_IN'])

    # ── NLR ──────────────────────────────────────────────────────────────
    print("\n--- Network Level Report ---")
    df_nlr = processar_nlr(ARQUIVO_NLR)
    df_resultado = df_resultado.merge(df_nlr, on='mcc_mnc', how='left')
    print(f"   Com NLR: {df_resultado['Natural_Market_Share'].notna().sum()} / "
          f"Sem NLR: {df_resultado['Natural_Market_Share'].isna().sum()}")

    # ── Nota Final ───────────────────────────────────────────────────────
    REF_SPEED = 20.0
    REF_RTT   = 180.0

    df_resultado['Nota_Speed_OUT'] = (df_resultado['Claro_Speed_Mediana_OUT'].fillna(0) / REF_SPEED) * 0.60
    df_resultado['Nota_RTT_IN']    = np.where(
        df_resultado['Claro_RTT_IN'].notna() & (df_resultado['Claro_RTT_IN'] > 0),
        (REF_RTT / df_resultado['Claro_RTT_IN']) * 0.30, 0)
    df_resultado['Nota_NMS']       = df_resultado['Natural_Market_Share'].fillna(0) * 0.05
    df_resultado['Nota_Registro']  = df_resultado['Registro_1a_Tentativa'].fillna(0) * 0.05
    df_resultado['Nota_Final']     = (df_resultado['Nota_Speed_OUT'] + df_resultado['Nota_RTT_IN'] +
                                      df_resultado['Nota_NMS'] + df_resultado['Nota_Registro'])

    df_resultado = df_resultado.sort_values('Nota_Final', ascending=False).reset_index(drop=True)
    df_resultado['Posicao_Qualidade'] = df_resultado.index + 1

    print(f"\n✅ META processado: {len(df_resultado)} operadoras.")
    return df_resultado


# ==============================================================================
# EXECUÇÃO STANDALONE (teste)
# ==============================================================================
if __name__ == '__main__':
    df = processar_meta_dados()
    df.to_csv(OUTPUT, sep=';', decimal=',', index=False, encoding='utf-8-sig')

    pd.set_option('display.float_format', '{:.4f}'.format)
    pd.set_option('display.max_columns', 20)
    pd.set_option('display.width', 220)

    cols_print = ['Posicao_Qualidade', 'Rede', 'mcc_mnc',
                  'Claro_Speed_Mediana_OUT', 'Claro_RTT_IN',
                  'Natural_Market_Share', 'Registro_1a_Tentativa', 'Nota_Final']
    print(f"\n{'─'*100}")
    print("TOP 20 — Nota Final")
    print(f"{'─'*100}")
    print(df[[c for c in cols_print if c in df.columns]].head(20).to_string(index=False))

    print(f"\n{'─'*50}")
    print(f"  Total operadoras:        {len(df)}")
    print(f"  Com Speed OUT:           {df['Claro_Speed_Mediana_OUT'].notna().sum()}")
    print(f"  Com RTT IN:              {df['Claro_RTT_IN'].notna().sum()}")
    print(f"  Com NLR:                 {df['Natural_Market_Share'].notna().sum()}")
    print(f"  Arquivo salvo: {OUTPUT}")
