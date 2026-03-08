"""
dashboard.py — Dashboard executivo completo (10 seções)
Chamado por core.py e core_nutricash.py
"""
import pandas as pd
import numpy as np


# ─────────────────────────────────────────────────────────────────────
#  FUNÇÃO PRINCIPAL
# ─────────────────────────────────────────────────────────────────────

def write_dashboard(wb, df_av_raw, df_vd_raw, TODAY, cor_hdr='#00205B'):
    """
    Cria aba DASHBOARD com 10 seções executivas.
    df_av_raw / df_vd_raw: DataFrames já processados pelo core.py
    """

    ws = wb.add_worksheet('DASHBOARD')
    ws.set_zoom(85)
    ws.hide_gridlines(2)

    # ── Paleta de cores ──────────────────────────────────────────────
    CHX = cor_hdr.lstrip('#')          # hex sem #
    C_HDR   = cor_hdr                  # azul marinho
    C_GRN   = '#1E7B34'               # verde escuro
    C_RED   = '#C00000'               # vermelho
    C_YEL   = '#F5C842'               # amarelo
    C_BLU2  = '#2E75B6'              # azul médio
    C_BGLT  = '#F4F7FC'              # fundo light
    C_BGDK  = '#E8EEF8'              # fundo medium

    def f(**kw):
        return wb.add_format({**{'font_name': 'Arial', 'font_size': 10}, **kw})

    # Formatos gerais
    F_TITLE   = f(bold=True, font_size=20, font_color=C_HDR, valign='vcenter')
    F_SUB     = f(italic=True, font_size=10, font_color='#666666')
    F_SEC     = f(bold=True, font_size=12, font_color='#FFFFFF',
                  bg_color=CHX, valign='vcenter', left=2, left_color=C_YEL.lstrip('#'))
    F_SEC2    = f(bold=True, font_size=11, font_color=C_HDR,
                  bg_color='E8EEF8', valign='vcenter', bottom=2, bottom_color=CHX)

    # KPI cards
    F_KPI_VAL = f(bold=True, font_size=20, font_color=C_HDR,
                  align='center', valign='vcenter', num_format='#,##0.00',
                  bg_color='F4F7FC', border=1, border_color='D0DCF0')
    F_KPI_INT = f(bold=True, font_size=20, font_color=C_HDR,
                  align='center', valign='vcenter', num_format='#,##0',
                  bg_color='F4F7FC', border=1, border_color='D0DCF0')
    F_KPI_PCT = f(bold=True, font_size=20, font_color=C_RED,
                  align='center', valign='vcenter', num_format='0.0%',
                  bg_color='FFF4F4', border=1, border_color='F0C0C0')
    F_KPI_GRN = f(bold=True, font_size=20, font_color=C_GRN,
                  align='center', valign='vcenter', num_format='#,##0.00',
                  bg_color='F0FFF4', border=1, border_color='B0E0BC')
    F_KPI_RED = f(bold=True, font_size=20, font_color=C_RED,
                  align='center', valign='vcenter', num_format='#,##0.00',
                  bg_color='FFF4F4', border=1, border_color='F0C0C0')
    F_KPI_LBL = f(font_size=9, font_color='#777777', align='center',
                  valign='vcenter', text_wrap=True,
                  bg_color='F4F7FC', border=1, border_color='D0DCF0')
    F_KPI_LBL_R = f(font_size=9, font_color=C_RED, align='center',
                    valign='vcenter', text_wrap=True,
                    bg_color='FFF4F4', border=1, border_color='F0C0C0')
    F_KPI_LBL_G = f(font_size=9, font_color=C_GRN, align='center',
                    valign='vcenter', text_wrap=True,
                    bg_color='F0FFF4', border=1, border_color='B0E0BC')

    # Tabelas
    F_TH  = f(bold=True, bg_color=CHX, font_color='FFFFFF',
               border=1, align='center', valign='vcenter', text_wrap=True)
    F_TH2 = f(bold=True, bg_color='E8EEF8', font_color=C_HDR,
               border=1, align='center', valign='vcenter')
    F_TA  = f(bg_color='FFFFFF', border=1, border_color='E0E8F0')
    F_TB  = f(bg_color='F7FAFF', border=1, border_color='E0E8F0')
    F_MA  = f(bg_color='FFFFFF', border=1, border_color='E0E8F0',
               num_format='#,##0.00', align='right')
    F_MB  = f(bg_color='F7FAFF', border=1, border_color='E0E8F0',
               num_format='#,##0.00', align='right')
    F_PA  = f(bg_color='FFFFFF', border=1, border_color='E0E8F0',
               num_format='0.0%', align='center')
    F_PB  = f(bg_color='F7FAFF', border=1, border_color='E0E8F0',
               num_format='0.0%', align='center')
    F_IA  = f(bg_color='FFFFFF', border=1, border_color='E0E8F0',
               num_format='0', align='center')
    F_IB  = f(bg_color='F7FAFF', border=1, border_color='E0E8F0',
               num_format='0', align='center')
    F_RKA = f(bold=True, bg_color='FFFFFF', border=1, border_color='E0E8F0', align='center')
    F_RKB = f(bold=True, bg_color='F7FAFF', border=1, border_color='E0E8F0', align='center')
    F_TOT = f(bold=True, bg_color='E8EEF8', border=1, border_color='C0CCE0',
               num_format='#,##0.00', align='right')
    F_TOTL= f(bold=True, bg_color='E8EEF8', border=1, border_color='C0CCE0')
    F_TOTP= f(bold=True, bg_color='E8EEF8', border=1, border_color='C0CCE0',
               num_format='0.0%', align='center')

    # Aging
    F_AG0 = f(bold=True, bg_color='D9EAD3', font_color='1E7B34',
               border=1, align='center', valign='vcenter', num_format='#,##0.00')
    F_AG1 = f(bold=True, bg_color='FFF2CC', font_color='7F6000',
               border=1, align='center', valign='vcenter', num_format='#,##0.00')
    F_AG2 = f(bold=True, bg_color='FCE5CD', font_color='783F04',
               border=1, align='center', valign='vcenter', num_format='#,##0.00')
    F_AG3 = f(bold=True, bg_color='F4CCCC', font_color='990000',
               border=1, align='center', valign='vcenter', num_format='#,##0.00')
    F_AG4 = f(bold=True, bg_color='C00000', font_color='FFFFFF',
               border=1, align='center', valign='vcenter', num_format='#,##0.00')
    F_AGL = f(bold=True, bg_color='F0F4FA', border=1, align='center', font_size=9)
    F_AGQ = f(bg_color='F0F4FA', border=1, align='center',
               font_size=8, font_color='666666')

    # Alertas
    F_ALRT_OK  = f(bg_color='D9EAD3', font_color='1E7B34', bold=True, border=1, align='center')
    F_ALRT_WRN = f(bg_color='FFF2CC', font_color='7F6000', bold=True, border=1, align='center')
    F_ALRT_ERR = f(bg_color='F4CCCC', font_color='990000', bold=True, border=1, align='center')
    F_ALRT_VAL = f(bg_color='FFFFFF', border=1, align='right', num_format='#,##0')
    F_ALRT_LBL = f(bg_color='FFFFFF', border=1)

    MES_NOME = {1:'Jan',2:'Fev',3:'Mar',4:'Abr',5:'Mai',6:'Jun',
                7:'Jul',8:'Ago',9:'Set',10:'Out',11:'Nov',12:'Dez'}
    MES_EXT  = {1:'Janeiro',2:'Fevereiro',3:'Março',4:'Abril',5:'Maio',
                6:'Junho',7:'Julho',8:'Agosto',9:'Setembro',10:'Outubro',
                11:'Novembro',12:'Dezembro'}

    # ── Preparar dados ───────────────────────────────────────────────
    df_av = df_av_raw.copy()
    df_vd = df_vd_raw.copy()
    df_all = pd.concat([df_av, df_vd], ignore_index=True)

    for df in [df_av, df_vd, df_all]:
        for col in ['VL SALDO','VL TITULO','ATRASO']:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        for col in ['EMISSAO','VENCIMENTO']:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors='coerce')

    if 'PAGA NA DATA' in df_av.columns:
        df_av['_paga'] = df_av['PAGA NA DATA'].apply(
            lambda v: pd.NaT if str(v).strip() in ('REVISAR','','nan','None') else v)
        df_av['_paga'] = pd.to_datetime(df_av['_paga'], errors='coerce')
    else:
        df_av['_paga'] = pd.NaT

    if 'PREV PAGTO' in df_av.columns:
        df_av['_prev'] = df_av['PREV PAGTO'].apply(
            lambda v: pd.NaT if str(v).strip() in ('REVISAR','','nan','None') else v)
        df_av['_prev'] = pd.to_datetime(df_av['_prev'], errors='coerce')
    else:
        df_av['_prev'] = pd.NaT

    def clean(df, col):
        if col not in df.columns: return pd.Series(['N/D']*len(df))
        return df[col].apply(lambda v: str(v).strip()
                              if str(v).strip() not in ('','nan','None','REVISAR')
                              else 'N/D')

    df_all['_cli']  = clean(df_all,'CLIENTE')
    df_all['_exec'] = clean(df_all,'EXECUTIVO')
    df_all['_prod'] = clean(df_all,'PRODUTO')
    df_all['_grp']  = clean(df_all,'GRUPO')
    df_all['_tipo'] = clean(df_all,'TIPO')
    df_all['_uf']   = clean(df_all,'UF')
    df_all['_cid']  = clean(df_all,'CIDADE')

    df_vd['_cli']   = clean(df_vd,'CLIENTE')
    df_vd['_exec']  = clean(df_vd,'EXECUTIVO')

    # ── KPI base ─────────────────────────────────────────────────────
    total_saldo = df_all['VL SALDO'].sum()
    total_titulo= df_all['VL TITULO'].sum()
    vencido_vl  = df_vd['VL SALDO'].sum()
    pct_inadi   = vencido_vl / total_saldo if total_saldo > 0 else 0

    av_mes = df_av[df_av['VENCIMENTO'].notna() &
                   (df_av['VENCIMENTO'].dt.month == TODAY.month) &
                   (df_av['VENCIMENTO'].dt.year  == TODAY.year)]['VL SALDO'].sum()

    recebido_mes = df_av[df_av['_paga'].notna() &
                         (df_av['_paga'].dt.month == TODAY.month) &
                         (df_av['_paga'].dt.year  == TODAY.year)]['VL SALDO'].sum()

    nf_validas = df_all['NF'].apply(
        lambda v: str(v).strip() not in ('','nan','None','REVISAR','0'))
    qtd_nf  = int(nf_validas.sum())
    ticket  = total_titulo / qtd_nf if qtd_nf > 0 else 0

    diff_prazo = (df_av['_paga'] - df_av['EMISSAO']).dt.days.dropna()
    prazo_medio = diff_prazo.mean() if len(diff_prazo) > 0 else 0
    if np.isnan(prazo_medio): prazo_medio = 0

    # ── Larguras de coluna ───────────────────────────────────────────
    # A  B    C    D    E    F    G    H    I    J    K    L    M
    ws.set_column(0, 0,  2)   # A: margem
    ws.set_column(1, 1, 22)   # B
    ws.set_column(2, 2, 18)   # C
    ws.set_column(3, 3, 15)   # D
    ws.set_column(4, 4, 15)   # E
    ws.set_column(5, 5,  2)   # F: separador
    ws.set_column(6, 6, 22)   # G
    ws.set_column(7, 7, 18)   # H
    ws.set_column(8, 8, 15)   # I
    ws.set_column(9, 9, 15)   # J
    ws.set_column(10,10,  2)  # K: margem

    # ════════════════════════════════════════════════════════════════
    # LINHA 0-1: TÍTULO
    # ════════════════════════════════════════════════════════════════
    ws.set_row(0, 44)
    ws.set_row(1, 18)
    ws.merge_range('B1:J1',
        f'📊  DASHBOARD EXECUTIVO — CONTAS A RECEBER | {MES_EXT[TODAY.month]}/{TODAY.year}',
        F_TITLE)
    ws.merge_range('B2:J2',
        f'Gerado em {TODAY.strftime("%d/%m/%Y")}  •  '
        f'{len(df_all):,} títulos  •  '
        f'A Vencer: {len(df_av):,}  •  Vencidos: {len(df_vd):,}',
        F_SUB)

    # ════════════════════════════════════════════════════════════════
    # LINHAS 3-9: SEÇÃO 1 — KPIs (7 cards)
    # ════════════════════════════════════════════════════════════════
    ws.set_row(3, 22)
    ws.merge_range('B4:J4', '  1. VISÃO EXECUTIVA', F_SEC)

    # 7 KPIs: B C D E  G H I J  (F é separador)
    # Cada KPI: header(r5) + valor(r6+r7 merged) + label(r8)
    kpi_cols = [1, 2, 3, 4,  6, 7, 8]  # idx 0-based
    kpi_data = [
        ('TOTAL A RECEBER (R$)',    total_saldo,    F_KPI_VAL, F_KPI_LBL),
        ('RECEBIDO NO MÊS (R$)',    recebido_mes,   F_KPI_GRN, F_KPI_LBL_G),
        ('A VENCER NO MÊS (R$)',    av_mes,         F_KPI_VAL, F_KPI_LBL),
        ('VENCIDO (R$)',            vencido_vl,     F_KPI_RED, F_KPI_LBL_R),
        ('% INADIMPLÊNCIA',         pct_inadi,      F_KPI_PCT, F_KPI_LBL_R),
        (f'TICKET MÉDIO / NF (R$)', ticket,         F_KPI_VAL, F_KPI_LBL),
        (f'PRAZO MÉDIO RECEB (dias)',prazo_medio,   F_KPI_INT, F_KPI_LBL),
    ]
    kpi_fmts_val = [F_KPI_VAL, F_KPI_GRN, F_KPI_VAL, F_KPI_RED,
                    F_KPI_PCT, F_KPI_VAL, F_KPI_INT]

    for k, (ci, (lbl, val, fval, flbl)) in enumerate(zip(kpi_cols, kpi_data)):
        ws.set_row(4, 20)
        ws.set_row(5, 36)
        ws.set_row(6, 36)
        ws.set_row(7, 18)
        ws.write(4, ci, '', fval)
        ws.merge_range(5, ci, 6, ci, float(val) if not isinstance(val, str) else val, fval)
        ws.write(7, ci, lbl, flbl)

    # ════════════════════════════════════════════════════════════════
    # LINHAS 9-35: SEÇÃO 2 — FLUXO DE CAIXA
    # ════════════════════════════════════════════════════════════════
    ws.set_row(8, 22)
    ws.merge_range('B9:J9', '  2. FLUXO DE CAIXA — Recebimentos Previstos', F_SEC)

    # Preparar dados fluxo: usar PREV PAGTO para A VENCER + VENCIMENTO p/ vencidos
    df_flux_av = df_av[df_av['_prev'].notna()].copy()
    df_flux_av['_flux_date'] = df_flux_av['_prev']
    df_flux_vd = df_vd[df_vd['VENCIMENTO'].notna()].copy()
    df_flux_vd['_flux_date'] = df_flux_vd['VENCIMENTO']
    df_flux = pd.concat([
        df_flux_av[['_flux_date','VL SALDO']],
        df_flux_vd[['_flux_date','VL SALDO']]
    ], ignore_index=True)
    df_flux['_flux_date'] = pd.to_datetime(df_flux['_flux_date'], errors='coerce')
    df_flux = df_flux[df_flux['_flux_date'].notna()]

    # Fluxo diário (mês atual)
    daily = df_flux[
        (df_flux['_flux_date'].dt.month == TODAY.month) &
        (df_flux['_flux_date'].dt.year  == TODAY.year)
    ].groupby(df_flux['_flux_date'].dt.day)['VL SALDO'].sum().sort_index()

    # Fluxo semanal (próximas 8 semanas)
    df_flux['_wk'] = df_flux['_flux_date'].dt.to_period('W')
    weekly = df_flux.groupby('_wk')['VL SALDO'].sum().tail(12)

    # Fluxo mensal (próximos 12 meses)
    df_flux['_mo'] = df_flux['_flux_date'].dt.to_period('M')
    monthly = df_flux.groupby('_mo')['VL SALDO'].sum().tail(12)

    # Hidden data rows for charts (start at row 300)


    # ════════════════════════════════════════════════════════════════
    # LINHAS 24-36: SEÇÃO 3 — AGING
    # ════════════════════════════════════════════════════════════════
    ws.set_row(23, 22)
    ws.merge_range('B24:J24', '  3. AGING LIST — Envelhecimento da Carteira', F_SEC)
    ws.set_row(24, 22)

    df_ag = df_all.copy()
    df_ag['VENCIMENTO'] = pd.to_datetime(df_ag['VENCIMENTO'], errors='coerce')

    def aging_bucket(row):
        v = row['VENCIMENTO']
        vl= row['VL SALDO']
        if pd.isna(v): return ('SEM DATA', vl)
        dias = (TODAY - v).days
        if dias <= 0:   return ('A Vencer', vl)
        elif dias <= 30: return ('0–30 dias', vl)
        elif dias <= 60: return ('31–60 dias', vl)
        elif dias <= 90: return ('61–90 dias', vl)
        else:            return ('+90 dias', vl)

    aging_order = ['A Vencer','0–30 dias','31–60 dias','61–90 dias','+90 dias']
    aging_fmts  = [F_AG0, F_AG1, F_AG2, F_AG3, F_AG4]
    aging_colors= ['#1E7B34','#F5C842','#FFA500','#E06666','#C00000']

    aging_data  = {}
    for _, row in df_ag.iterrows():
        lbl, vl = aging_bucket(row)
        if lbl in aging_order:
            aging_data.setdefault(lbl, {'vl': 0.0, 'qtd': 0})
            aging_data[lbl]['vl']  += float(vl)
            aging_data[lbl]['qtd'] += 1

    # Header
    aging_hdrs = ['FAIXA','TOTAL (R$)','QTD TÍTULOS','% DO TOTAL','SITUAÇÃO']
    AG_R = 24
    for ci, h in enumerate(aging_hdrs):
        ws.write(AG_R, 1+ci, h, F_TH)
    ws.set_row(AG_R, 24)

    total_ag = sum(d['vl'] for d in aging_data.values())
    for ri, (lbl, fmt_ag, color) in enumerate(zip(aging_order, aging_fmts, aging_colors)):
        d   = aging_data.get(lbl, {'vl': 0.0, 'qtd': 0})
        pct = d['vl'] / total_ag if total_ag > 0 else 0
        ws.write(AG_R+1+ri, 1, lbl,    fmt_ag)
        ws.write_number(AG_R+1+ri, 2,  float(d['vl']),  fmt_ag)
        ws.write_number(AG_R+1+ri, 3,  d['qtd'],        f(bold=True,bg_color=color.lstrip('#'),font_color='FFFFFF' if ri>=3 else '000000', border=1, align='center', valign='vcenter', num_format='#,##0'))
        ws.write_number(AG_R+1+ri, 4,  pct,             f(bold=True,bg_color=color.lstrip('#'),font_color='FFFFFF' if ri>=3 else '000000', border=1, align='center', valign='vcenter', num_format='0.0%'))
        situacao = ['✅ Em dia','⚠️ Atenção','⚠️ Alerta','🔴 Crítico','🔴 Crítico'][ri]
        ws.write(AG_R+1+ri, 5, situacao, f(bg_color=color.lstrip('#'),font_color='FFFFFF' if ri>=3 else '000000', border=1, align='center'))
    # Total
    ws.write(AG_R+6, 1,  'TOTAL',        F_TOTL)
    ws.write_number(AG_R+6, 2, total_ag, F_TOT)
    ws.write(AG_R+6, 3,  '',  F_TOTL)
    ws.write(AG_R+6, 4,  '',  F_TOTP)
    ws.write(AG_R+6, 5,  '',  F_TOTL)


    # ════════════════════════════════════════════════════════════════
    # LINHAS 34-47: SEÇÃO 4+5 — CLIENTES & EXECUTIVOS (lado a lado)
    # ════════════════════════════════════════════════════════════════
    S45 = AG_R + 8
    ws.set_row(S45, 22)
    ws.merge_range(S45, 1, S45, 4,
                   '  4. RANKING DE CLIENTES — Concentração de Receita', F_SEC)
    ws.merge_range(S45, 6, S45, 9,
                   '  5. PERFORMANCE POR EXECUTIVO', F_SEC)
    ws.set_row(S45+1, 22)

    # Tabela clientes
    top_cli = df_all.groupby('_cli')['VL SALDO'].sum().sort_values(ascending=False).head(10)
    total_cli = df_all['VL SALDO'].sum()
    cli_hdrs = ['#','CLIENTE','VL SALDO (R$)','PART %']
    for ci,h in enumerate(cli_hdrs): ws.write(S45+1, 1+ci, h, F_TH)
    for i,(cli,vl) in enumerate(top_cli.items()):
        fa=F_TA if i%2==0 else F_TB; fm=F_MA if i%2==0 else F_MB
        fp=F_PA if i%2==0 else F_PB; fr=F_RKA if i%2==0 else F_RKB
        ws.write(S45+2+i, 1, i+1,          fr)
        ws.write(S45+2+i, 2, cli[:25],     fa)
        ws.write_number(S45+2+i, 3, float(vl), fm)
        ws.write_number(S45+2+i, 4, vl/total_cli if total_cli>0 else 0, fp)
    ws.write(S45+12,1,'TOTAL',F_TOTL); ws.write(S45+12,2,'',F_TOTL)
    ws.write_number(S45+12,3,float(top_cli.sum()),F_TOT)
    ws.write_number(S45+12,4,top_cli.sum()/total_cli if total_cli>0 else 0, F_TOTP)

    # Tabela executivos (faturamento, carteira, inadimplência)
    exec_tit = df_all.groupby('_exec')['VL TITULO'].sum().sort_values(ascending=False).head(10)
    exec_sal = df_all.groupby('_exec')['VL SALDO'].sum()
    exec_vd  = df_vd.groupby('_exec')['VL SALDO'].sum() if len(df_vd)>0 else pd.Series([],dtype=float)
    exec_hdrs= ['#','EXECUTIVO','FATURADO (R$)','CARTEIRA (R$)','INADI %']
    for ci,h in enumerate(exec_hdrs): ws.write(S45+1, 6+ci, h, F_TH)
    for i,(exc,tit) in enumerate(exec_tit.items()):
        sal = exec_sal.get(exc,0); vd_val = exec_vd.get(exc,0)
        pct_e = vd_val/sal if sal>0 else 0
        fa=F_TA if i%2==0 else F_TB; fm=F_MA if i%2==0 else F_MB
        fp=F_PA if i%2==0 else F_PB; fr=F_RKA if i%2==0 else F_RKB
        ws.write(S45+2+i, 6, i+1,          fr)
        ws.write(S45+2+i, 7, exc[:22],     fa)
        ws.write_number(S45+2+i, 8, float(tit),  fm)
        ws.write_number(S45+2+i, 9, float(sal),  fm)
        ws.write_number(S45+2+i,10, pct_e,        fp)
    ws.write(S45+12,6,'TOTAL',F_TOTL); ws.write(S45+12,7,'',F_TOTL)
    ws.write_number(S45+12,8,float(exec_tit.sum()),F_TOT)
    ws.write_number(S45+12,9,float(exec_sal.reindex(exec_tit.index).sum()),F_TOT)
    ws.write(S45+12,10,'',F_TOTP)


    # ════════════════════════════════════════════════════════════════
    # SEÇÃO 6+7 — PRODUTO & SEGMENTO (lado a lado)
    # ════════════════════════════════════════════════════════════════
    S67 = S45 + 14
    ws.set_row(S67, 22)
    ws.merge_range(S67, 1, S67, 4, '  6. RECEITA POR PRODUTO', F_SEC)
    ws.merge_range(S67, 6, S67, 9, '  7. RECEITA POR SEGMENTO (GRUPO & TIPO)', F_SEC)
    ws.set_row(S67+1, 22)

    # Produto
    prod_ser = df_all.groupby('_prod')['VL TITULO'].sum().sort_values(ascending=False).head(8)
    prod_hdrs= ['PRODUTO','VL TITULO (R$)','PART %']
    for ci,h in enumerate(prod_hdrs): ws.write(S67+1, 1+ci, h, F_TH)
    for i,(p,v) in enumerate(prod_ser.items()):
        fa=F_TA if i%2==0 else F_TB; fm=F_MA if i%2==0 else F_MB; fp=F_PA if i%2==0 else F_PB
        ws.write(S67+2+i, 1, p[:22],  fa)
        ws.write_number(S67+2+i, 2, float(v), fm)
        ws.write_number(S67+2+i, 3, v/total_titulo if total_titulo>0 else 0, fp)
    ws.write(S67+10,1,'TOTAL',F_TOTL); ws.write_number(S67+10,2,float(prod_ser.sum()),F_TOT); ws.write(S67+10,3,'',F_TOTP)


    # Segmento: GRUPO
    grp_ser  = df_all.groupby('_grp')['VL SALDO'].sum().sort_values(ascending=False).head(6)
    tipo_ser = df_all.groupby('_tipo')['VL SALDO'].sum().sort_values(ascending=False)
    seg_hdrs = ['GRUPO / TIPO','VL SALDO (R$)','PART %']
    for ci,h in enumerate(seg_hdrs): ws.write(S67+1, 6+ci, h, F_TH)
    # Grupos
    ws.write(S67+2, 6, '— Por Grupo —', F_TH2)
    ws.write(S67+2, 7, '', F_TH2); ws.write(S67+2, 8, '', F_TH2)
    for i,(g,v) in enumerate(grp_ser.items()):
        fa=F_TA if i%2==0 else F_TB; fm=F_MA if i%2==0 else F_MB; fp=F_PA if i%2==0 else F_PB
        ws.write(S67+3+i, 6, g[:22],  fa)
        ws.write_number(S67+3+i, 7, float(v), fm)
        ws.write_number(S67+3+i, 8, v/total_saldo if total_saldo>0 else 0, fp)
    # Tipo
    ws.write(S67+10, 6, '— Por Tipo —', F_TH2)
    ws.write(S67+10, 7, '', F_TH2); ws.write(S67+10, 8, '', F_TH2)
    for i,(t,v) in enumerate(tipo_ser.items()):
        fa=F_TA if i%2==0 else F_TB; fm=F_MA if i%2==0 else F_MB; fp=F_PA if i%2==0 else F_PB
        ws.write(S67+11+i, 6, t[:22], fa)
        ws.write_number(S67+11+i, 7, float(v), fm)
        ws.write_number(S67+11+i, 8, v/total_saldo if total_saldo>0 else 0, fp)

    # ════════════════════════════════════════════════════════════════
    # SEÇÃO 8 — REGIÃO
    # ════════════════════════════════════════════════════════════════
    S8 = S67 + 14
    ws.set_row(S8, 22)
    ws.merge_range(S8, 1, S8, 9, '  8. RECEITA POR REGIÃO (UF)', F_SEC)
    ws.set_row(S8+1, 22)

    uf_ser = df_all[df_all['_uf']!='N/D'].groupby('_uf')['VL SALDO'].sum().sort_values(ascending=False).head(12)
    reg_hdrs = ['UF','VL SALDO (R$)','PART %','']
    for ci,h in enumerate(reg_hdrs): ws.write(S8+1, 1+ci, h, F_TH)
    for ci,h in enumerate(reg_hdrs): ws.write(S8+1, 5+ci, h, F_TH)
    mid = len(uf_ser)//2 + len(uf_ser)%2
    uf_list = list(uf_ser.items())
    total_uf = uf_ser.sum()
    for i,(uf,v) in enumerate(uf_list[:mid]):
        fa=F_TA if i%2==0 else F_TB; fm=F_MA if i%2==0 else F_MB; fp=F_PA if i%2==0 else F_PB
        ws.write(S8+2+i,1,uf,fa); ws.write_number(S8+2+i,2,float(v),fm)
        ws.write_number(S8+2+i,3,v/total_uf if total_uf>0 else 0,fp); ws.write(S8+2+i,4,'',fa)
    for i,(uf,v) in enumerate(uf_list[mid:]):
        fa=F_TA if i%2==0 else F_TB; fm=F_MA if i%2==0 else F_MB; fp=F_PA if i%2==0 else F_PB
        ws.write(S8+2+i,5,uf,fa); ws.write_number(S8+2+i,6,float(v),fm)
        ws.write_number(S8+2+i,7,v/total_uf if total_uf>0 else 0,fp); ws.write(S8+2+i,8,'',fa)


    # ════════════════════════════════════════════════════════════════
    # SEÇÃO 9 — MONITORAMENTO DE ATRASOS
    # ════════════════════════════════════════════════════════════════
    S9 = S8 + mid + 3
    ws.set_row(S9, 22)
    ws.merge_range(S9, 1, S9, 4,
                   '  9. MONITORAMENTO DE ATRASOS — Top Inadimplentes', F_SEC)
    ws.merge_range(S9, 6, S9, 9,
                   '  9b. INADIMPLÊNCIA POR EXECUTIVO', F_SEC)
    ws.set_row(S9+1, 22)

    df_vd2 = df_vd.copy()
    df_vd2['ATRASO'] = pd.to_numeric(df_vd2['ATRASO'], errors='coerce').fillna(0)
    df_vd2['_cli2'] = clean(df_vd2,'CLIENTE')
    df_vd2['_exec2']= clean(df_vd2,'EXECUTIVO')

    cli_inadi = df_vd2.groupby('_cli2').agg(
        vl=('VL SALDO','sum'),
        qtd=('VL SALDO','count'),
        atraso_medio=('ATRASO','mean')
    ).sort_values('vl',ascending=False).head(10)

    mon_hdrs = ['#','CLIENTE','VL VENCIDO (R$)','TÍTULOS','ATRASO MÉDIO (dias)']
    for ci,h in enumerate(mon_hdrs): ws.write(S9+1, 1+ci, h, F_TH)
    for i,(cli,row) in enumerate(cli_inadi.iterrows()):
        fa=F_TA if i%2==0 else F_TB; fm=F_MA if i%2==0 else F_MB; fr=F_RKA if i%2==0 else F_RKB
        fi=F_IA if i%2==0 else F_IB
        ws.write(S9+2+i, 1, i+1,              fr)
        ws.write(S9+2+i, 2, cli[:25],         fa)
        ws.write_number(S9+2+i, 3, float(row['vl']),   fm)
        ws.write_number(S9+2+i, 4, int(row['qtd']),    fi)
        ws.write_number(S9+2+i, 5, float(row['atraso_medio']), fi)

    ws.write(S9+12,1,'TOTAL',F_TOTL); ws.write(S9+12,2,'',F_TOTL)
    ws.write_number(S9+12,3,float(cli_inadi['vl'].sum()),F_TOT)
    ws.write(S9+12,4,'',F_TOTL); ws.write(S9+12,5,'',F_TOTL)

    # Ranking inadimplência por executivo (lado direito)
    exec_inadi = df_vd2.groupby('_exec2').agg(
        vl=('VL SALDO','sum'),
        atraso=('ATRASO','mean')
    ).sort_values('vl',ascending=False).head(10)
    ex_hdrs=['#','EXECUTIVO','VL VENCIDO (R$)','ATRASO MÉDIO']
    for ci,h in enumerate(ex_hdrs): ws.write(S9+1, 6+ci, h, F_TH)
    for i,(exc,row) in enumerate(exec_inadi.iterrows()):
        fa=F_TA if i%2==0 else F_TB; fm=F_MA if i%2==0 else F_MB; fr=F_RKA if i%2==0 else F_RKB; fi=F_IA if i%2==0 else F_IB
        ws.write(S9+2+i, 6, i+1,         fr)
        ws.write(S9+2+i, 7, exc[:22],    fa)
        ws.write_number(S9+2+i, 8, float(row['vl']),  fm)
        ws.write_number(S9+2+i, 9, float(row['atraso']), fi)

    # ════════════════════════════════════════════════════════════════
    # SEÇÃO 10 — CONTROLE OPERACIONAL (ALERTAS)
    # ════════════════════════════════════════════════════════════════
    S10 = S9 + 14
    ws.set_row(S10, 22)
    ws.merge_range(S10, 1, S10, 9,
                   '  10. CONTROLE OPERACIONAL — Alertas Automáticos', F_SEC)
    ws.set_row(S10+1, 22)

    # Calcular alertas
    nf_sem    = df_all['NF'].apply(
        lambda v: str(v).strip() in ('','nan','None','0')).sum()
    diverge   = ((df_all['VL TITULO'] - df_all['VL SALDO']).abs() > 0.01).sum()
    sem_prev  = 0
    if 'PREV PAGTO' in df_av.columns:
        sem_prev = df_av['PREV PAGTO'].apply(
            lambda v: str(v).strip() in ('REVISAR','','nan','None')).sum()
    revisar_grp = (df_all['GRUPO'].apply(lambda v: str(v).strip()=='REVISAR')).sum() if 'GRUPO' in df_all.columns else 0
    venc_sem_atraso = df_vd2['ATRASO'].apply(lambda v: pd.isna(v) or v==0).sum()

    alertas = [
        ('📄 NFs sem número',
         nf_sem,
         F_ALRT_OK if nf_sem==0 else F_ALRT_ERR,
         '✅ OK' if nf_sem==0 else '❌ Verificar'),
        ('💰 Divergência VL TITULO vs VL SALDO',
         diverge,
         F_ALRT_OK if diverge==0 else F_ALRT_WRN,
         '✅ OK' if diverge==0 else '⚠️ Revisar'),
        ('📅 Títulos sem previsão de pagamento (A Vencer)',
         sem_prev,
         F_ALRT_OK if sem_prev==0 else F_ALRT_WRN,
         '✅ OK' if sem_prev==0 else '⚠️ Atualizar'),
        ('🔴 Clientes novos (REVISAR — sem histórico)',
         revisar_grp,
         F_ALRT_OK if revisar_grp==0 else F_ALRT_WRN,
         '✅ OK' if revisar_grp==0 else '⚠️ Cadastrar'),
        ('⏳ Títulos vencidos sem atraso preenchido',
         venc_sem_atraso,
         F_ALRT_OK if venc_sem_atraso==0 else F_ALRT_WRN,
         '✅ OK' if venc_sem_atraso==0 else '⚠️ Atualizar'),
    ]
    alrt_hdrs = ['ALERTA','QUANTIDADE','STATUS']
    for ci,h in enumerate(alrt_hdrs): ws.write(S10+1, 1+ci, h, F_TH)
    for i,(lbl,qtd,fmt_s,status) in enumerate(alertas):
        ws.write(S10+2+i, 1, lbl,    F_ALRT_LBL)
        ws.write_number(S10+2+i, 2, int(qtd), F_ALRT_VAL)
        ws.write(S10+2+i, 3, status, fmt_s)
        ws.set_row(S10+2+i, 20)
    ws.set_column(1,1,38)  # coluna B mais larga p/ alertas

    # ── Freeze e volta ao topo ────────────────────────────────────
    ws.freeze_panes(2, 0)
    ws.set_row(0, 44)

    print(f"   ✓ DASHBOARD: {len(df_all):,} títulos · "
          f"R$ {total_saldo:,.2f} total · "
          f"Inadimplência: {pct_inadi:.1%}")
