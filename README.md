# Roaming Ranking Report

> ### Em resumo
> **Problema:** o relatório mensal de qualidade de roaming dependia de juntar à mão 5 fontes diferentes (rankings de rede, tráfego, receita e custo, acordos comerciais), cada uma com um identificador diferente.  
> **Solução:** pipeline em Python que lê as 5 fontes, padroniza a chave entre elas, calcula uma nota de qualidade por operadora e gera o Excel final já formatado.  
> **Ferramentas:** Python, pandas, NumPy, xlsxwriter.  
> **Resultado:** um único comando gera o relatório com 3 abas e o ranking de 257 operadoras, incluindo o diagnóstico das que ficaram sem dados.  
> **Como isso ajuda um cliente:** se você monta todo mês o mesmo relatório juntando várias planilhas, esse processo pode virar um script que entrega o arquivo pronto.  

Pipeline de dados em Python para geração do relatório mensal de qualidade de roaming internacional, integrando múltiplas fontes de dados operacionais e de mercado em um único Excel formatado.

---

## Visão Geral

O relatório consolida **5 fontes de dados distintas** e produz um ranking de operadoras visitadas com base em critérios de qualidade de rede e tráfego, entregue como Excel com formatação e múltiplas abas.

```
Fontes de entrada                     Pipeline                  Saída
─────────────────                     ────────                  ─────
META/Opensignal (OUT + IN)  ──┐
Network Level Report (NLR)  ──┤  processar_meta.py  ──┐
                               │                        ├──▶  relatorio_geral_roaming_{mes}.xlsx
OMR Tráfego (OUT + IN)      ──┤                        │        ├── Aba: Top - Tráfego
OMR Receita e Custo         ──┤  relatorio_geral_v4.py─┘        ├── Aba: Top - Qualidade
Acordos Comerciais          ──┘                                  └── Aba: Sem Dados Qualidade
```

---

## Fórmula de Qualidade

A nota de cada operadora (escala 0–1) combina quatro indicadores com pesos alinhados com a área de roaming:

| Indicador | Fonte | Peso | Referência | Direção |
|---|---|---|---|---|
| Velocidade Mediana Download OUT (4G+5G) | META Outbound | **60%** | 20 Mbps | Maior é melhor |
| RTT (latência) IN (4G+5G) | META Inbound | **30%** | 180 ms | Menor é melhor |
| Natural Market Share | Network Level Report | **5%** | 100% | Maior é melhor |
| Taxa de Registro em 1ª Tentativa | Network Level Report | **5%** | 100% | Maior é melhor |

```
Nota = (Speed_OUT / 20) × 0.60
     + (180 / RTT_IN)  × 0.30
     + NMS              × 0.05
     + Registro_1a      × 0.05
```

### Combinação 4G + 5G

Métricas de velocidade e latência combinam 4G e 5G por **média ponderada pelo número de amostras**, garantindo que tecnologias com mais dados medidos tenham maior influência:

```
Speed_combined = (Speed_4G × Samples_4G + Speed_5G × Samples_5G)
                 ─────────────────────────────────────────────────
                           Samples_4G + Samples_5G
```

---

## Estrutura do Projeto

```
roaming-ranking-report/
│
├── src/
│   ├── processar_meta.py        # Módulo de processamento META + NLR
│   │                            # Carrega, limpa, combina 4G/5G e calcula a Nota
│   │                            # Pode ser importado por outros scripts
│   │
│   └── relatorio_geral_v4.py   # Script principal do relatório
│                                # Une todas as fontes e gera o Excel
│
├── utils/
│   ├── validar_mcc_mnc.py       # Cruza MCC/MNCs contra NLR e Acordos
│   ├── adicionar_correspondencias.py  # Adiciona entradas na tabela de referência
│   └── listar_sem_correspondencia.py  # Diagnostica operadoras sem MCC/MNC mapeado
│
├── data/
│   ├── correspondencia.csv      # Tabela de referência: nome da rede → MCC/MNC
│   └── .gitkeep                 # (arquivos de dados mensais ficam aqui, não versionados)
│
├── output/
│   └── .gitkeep                 # Relatórios gerados ficam aqui, não versionados
│
├── .gitignore
├── requirements.txt
└── README.md
```

---

## Como Usar

### 1. Instalar dependências

```bash
pip install -r requirements.txt
```

### 2. Colocar os arquivos do mês em `data/`

Renomeie os arquivos recebidos para os nomes esperados pelo script:

| Nome esperado em `data/` | Arquivo original |
|---|---|
| `Outbound_TOP 100 Ranking_Compare2G-3G-4G-5G.csv` | Ranking META Outbound TOP 100 |
| `Outbound_TOP101_200_Compare2G-3G-4G-5G.csv` | Ranking META Outbound TOP 101-200 |
| `Inbound_TOP 100 Ranking_Compare2G-3G-4G-5G.csv` | Ranking META Inbound TOP 100 |
| `Inbound_TOP101_200_Compare2G-3G-4G-5G.csv` | Ranking META Inbound TOP 101-200 |
| `Network Level Report.xls` | Network Level Report (NLR) |
| `ACORDOS.xlsx` | Planilha de acordos comerciais |
| `GSM Roaming Outbound - Trafego OMR.xlsx` | Tráfego Outbound OMR |
| `GSM Roaming Inbound - Trafego OMR.xlsx` | Tráfego Inbound OMR |
| `Receita e Custo - Tarifas OMR.xlsx` | Receita e Custo OMR |

### 3. Gerar o relatório

```bash
python src/relatorio_geral_v4.py
```

O arquivo será salvo em `output/relatorio_geral_roaming_{mes}.xlsx`.

### 4. Testar só o processamento META (opcional)

```bash
python src/processar_meta.py
```

Gera `output/resultado_meta_teste.csv` com a nota de qualidade por operadora.

---

## Utilitários

### Diagnosticar operadoras sem MCC/MNC mapeado

```bash
python utils/listar_sem_correspondencia.py
```

Lista operadoras presentes no OMR mas ausentes em `data/correspondencia.csv`, com tráfego e TADIG para facilitar a pesquisa.

### Validar MCC/MNCs contra fontes externas

```bash
python utils/validar_mcc_mnc.py
```

Cruza os MCC/MNCs construídos a partir do OMR contra o NLR e a planilha de Acordos, identificando divergências e confirmações.

### Adicionar novas entradas na correspondência

```bash
python utils/adicionar_correspondencias.py
```

Adiciona entradas em `data/correspondencia.csv` evitando duplicatas. Edite o arquivo `NOVAS_ENTRADAS` dentro do script antes de executar.

---

## Desafios Técnicos Resolvidos

### MNC de 3 dígitos
O arquivo OMR armazena o MNC como número inteiro no Excel, perdendo zeros à esquerda. MCC `334` + MNC `20` virava `33420` em vez de `334020` (Telcel México). A solução testa as versões de 2 e 3 dígitos contra a tabela de referência e usa a que corresponder:

```python
def construir_mcc_mnc(mcc, mnc):
    c2 = mcc + mnc.zfill(2)
    c3 = mcc + mnc.zfill(3)
    if c2 in corr_mccs: return c2
    if c3 in corr_mccs: return c3
    return c2  # fallback
```

### Duplicatas nos rankings META
O mesmo operador podia aparecer no arquivo TOP 100 e no TOP 101-200 simultaneamente. A deduplicação mantém a primeira ocorrência (TOP 100) antes de qualquer join:

```python
df_raw = df_raw.drop_duplicates(subset=[0], keep='first')
```

### Múltiplas chaves de join
Cada fonte usa um identificador diferente. O pipeline constrói um MCC/MNC como chave universal e usa o arquivo OMR Outbound como bridge TADIG → MCC/MNC para as demais fontes.

---

## Saída

O Excel final contém 3 abas:

| Aba | Descrição |
|---|---|
| **Top - Tráfego** | 257 operadoras ordenadas por volume de tráfego Outbound |
| **Top - Qualidade** | As mesmas operadoras ordenadas pela Nota de qualidade |
| **Sem Dados Qualidade** | Operadoras sem nota, com diagnóstico do motivo |

Formatação por grupos de colunas com cores distintas:

- 🟣 **Roxo** — Posição e Nota de qualidade  
- 🔵 **Azul** — Informações da operadora e tráfego  
- 🟢 **Verde** — Acordos comerciais  
- 🟠 **Laranja** — Financeiro (receita e custo)  
- 🟦 **Índigo** — KPIs META (velocidade, RTT, cobertura)

---

## Tecnologias

- **Python 3.11+**
- **pandas** — manipulação de dados e joins
- **numpy** — cálculos vetorizados
- **xlsxwriter** — geração do Excel com formatação avançada
- **xlrd / openpyxl** — leitura de arquivos `.xls` e `.xlsx`

---

## Observações

- Os arquivos de dados mensais não são versionados (`.gitignore`) por conterem informações operacionais sensíveis.
- O arquivo `data/correspondencia.csv` é o único arquivo de dados versionado — é uma tabela de referência pública que mapeia nomes de operadoras para seus códigos MCC/MNC.
- O relatório cobre ~257 operadoras com tráfego, das quais ~127 possuem dados de qualidade META disponíveis.
