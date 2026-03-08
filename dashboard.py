"""
dashboard.py — Dashboard executivo — mês corrente
Chamado por core.py e core_nutricash.py
"""
import pandas as pd
import numpy as np


def write_dashboard(wb, df_av_raw, df_vd_raw, TODAY, cor_hdr='#00205B'):

    ws = wb.add_worksheet('DASHBOARD')
    ws.set_zoom(85)
    ws.hide_gridlines(2)

    CHX   = cor_hdr.lstrip('#')
    C_HDR = cor_hdr
    C_RED = '#C00000'
    C_YEL = '#F5C842'

    MES_EXT = {1:'Janeiro',2:'Fevereiro',3:'Março',4:'Abril',5:'Maio',
               6:'Junho',7:'Julho',8:'Agosto',9:'Setembro',10:'Outubro',
               11:'Novembro',12:'Dezembro'}
    MES_ATUAL = MES_EXT[TODAY.month]
    ANO_ATUAL = TODAY.year

    def f(**kw):
        return wb.add_format({**{'font_name':'Arial','font_size':10}, **kw})

    F_TITLE  = f(bold=True, font_size=18, font_color=C_HDR, valign='vcenter')
    F_SUB    = f(italic=True, font_size=10, font_color='#666666')
    F_SEC    = f(bold=True, font_size=12, font_color='#FFFFFF',
                 bg_color=CHX, valign='vcenter', left=2, left_color=C_YEL.lstrip('#'))
    F_TH     = f(bold=True, bg_color=CHX, font_color='FFFFFF',
                 border=1, align='center', valign='vcenter', text_wrap=True)
    F_TH2    = f(bold=True, bg_color='E8EEF8', font_color=C_HDR,
                 border=1, align='center', valign='vcenter')
    F_TA     = f(bg_color='FFFFFF', border=1, border_color='E0E8F0')
    F_TB     = f(bg_color='F7FAFF', border=1, border_color='E0E8F0')
    F_MA     = f(bg_color='FFFFFF', border=1, border_color='E0E8F0', num_format='#,##0.00', align='right')
    F_MB     = f(bg_color='F7FAFF', border=1, border_color='E0E8F0', num_format='#,##0.00', align='right')
    F_PA     = f(bg_color='FFFFFF', border=1, border_color='E0E8F0', num_format='0.0%', align='center')
    F_PB     = f(bg_color='F7FAFF', border=1, border_color='E0E8F0', num_format='0.0%', align='center')
    F_IA     = f(bg_color='FFFFFF', border=1, border_color='E0E8F0', num_format='0', align='center')
    F_IB     = f(bg_color='F7FAFF', border=1, border_color='E0E8F0', num_format='0', align='center')
    F_RKA    = f(bold=True, bg_color='FFFFFF', border=1, border_color='E0E8F0', align='center')
    F_RKB    = f(bold=True, bg_color='F7FAFF', border=1, border_color='E0E8F0', align='center')
    F_TOT    = f(bold=True, bg_color='E8EEF8', border=1, border_color='C0CCE0', num_format='#,##0.00', align='right')
    F_TOTL   = f(bold=True, bg_color='E8EEF8', border=1, border_color='C0CCE0')
    F_TOTP   = f(bold=True, bg_color='E8EEF8', border=1, border_color='C0CCE0', num_format='0.0%', align='center')
    F_KPI_VAL= f(bold=True, font_size=20, font_color=C_HDR, align='center', valign='vcenter',
                 num_format='#,##0.00', bg_color='F4F7FC', border=1, border_color='D0DCF0')
    F_KPI_INT= f(bold=True, font_size=20, font_color=C_HDR, align='center', valign='vcenter',
                 num_format='#,##0', bg_color='F4F7FC', border=1, border_color='D0DCF0')
    F_KPI_PCT= f(bold=True, font_size=20, font_color=C_RED, align='center', valign='vcenter',
                 num_format='0.0%', bg_color='FFF4F4', border=1, border_color='F0C0C0')
    F_KPI_RED= f(bold=True, font_size=20, font_color=C_RED, align='center', valign='vcenter',
                 num_format='#,##0.00', bg_color='FFF4F4', border=1, border_color='F0C0C0')
    F_KPI_LBL= f(font_size=9, font_color='#777777', align='center', valign='vcenter', text_wrap=True,
                 bg_color='F4F7FC', border=1, border_color='D0DCF0')
    F_KPI_LBL_R= f(font_size=9, font_color=C_RED, align='center', valign='vcenter', text_wrap=True,
                   bg_color='FFF4F4', border=1, border_color='F0C0C0')
    F_ALRT_OK = f(bg_color='D9EAD3', font_color='1E7B34', bold=True, border=1, align='center')
    F_ALRT_WRN= f(bg_color='FFF2CC', font_color='7F6000', bold=True, border=1, align='center')
    F_ALRT_ERR= f(bg_color='F4CCCC', font_color='990000', bold=True, border=1, align='center')
    F_ALRT_VAL= f(bg_color='FFFFFF', border=1, align='right', num_format='#,##0')
    F_ALRT_LBL= f(bg_color='FFFFFF', border=1)

    ws.set_column(0, 0,  2)
    ws.set_column(1, 1, 30)
    ws.set_column(2, 2, 20)
    ws.set_column(3, 3, 16)
    ws.set_column(4, 4, 16)
    ws.set_column(5, 5,  3)
    ws.set_column(6, 6, 30)
    ws.set_column(7, 7, 20)
    ws.set_column(8, 8, 16)
    ws.set_column(9, 9, 16)
    ws.set_column(10,10,  2)

    # ── Preparar DataFrames ──────────────────────────────────────────
    df_av  = df_av_raw.copy()
    df_vd  = df_vd_raw.copy()

    for df in [df_av, df_vd]:
        for col in ['VL SALDO','VL TITULO','IR','ISS','ATRASO']:
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
        if col not in df.columns:
            return pd.Series(['N/D']*len(df), index=df.index)
        return df[col].apply(lambda v: str(v).strip()
                             if str(v).strip() not in ('','nan','None','REVISAR') else 'N/D')

    # ── Filtros mês corrente ─────────────────────────────────────────
    def mes_mask(df, col='VENCIMENTO'):
        return (df[col].notna() &
                (df[col].dt.month == TODAY.month) &
                (df[col].dt.year  == TODAY.year))

    # A Vencer: de HOJE em diante, dentro do mês
    df_av_mes = df_av[
        df_av['VENCIMENTO'].notna() &
        (df_av['VENCIMENTO'].dt.month == TODAY.month) &
        (df_av['VENCIMENTO'].dt.year  == TODAY.year) &
        (df_av['VENCIMENTO'] >= TODAY)
    ].copy()

    # Vencido: dentro do mês corrente
    df_vd_mes = df_vd[mes_mask(df_vd)].copy()

    # Campos auxiliares
    for df in [df_av_mes, df_vd_mes]:
        df['_cli']  = clean(df, 'CLIENTE')
        df['_exec'] = clean(df, 'EXECUTIVO')
        df['_prod'] = clean(df, 'PRODUTO')
        df['_grp']  = clean(df, 'GRUPO')
        df['_tipo'] = clean(df, 'TIPO')
        df['_uf']   = clean(df, 'UF')

    df_all_mes = pd.concat([df_av_mes, df_vd_mes], ignore_index=True)
    for col in ['_cli','_exec','_prod','_grp','_tipo','_uf']:
        if col not in df_all_mes.columns:
            df_all_mes[col] = 'N/D'

    # ── KPIs ─────────────────────────────────────────────────────────
    av_mes = df_av_mes['VL SALDO'].sum()
    for col in ['IR','ISS']:
        if col in df_av_mes.columns:
            av_mes += df_av_mes[col].sum()

    vencido_mes = df_vd_mes['VL SALDO'].sum()
    for col in ['IR','ISS']:
        if col in df_vd_mes.columns:
            vencido_mes += df_vd_mes[col].sum()

    total_mes = av_mes + vencido_mes
    pct_inadi = vencido_mes / total_mes if total_mes > 0 else 0

    nf_val  = df_all_mes['NF'].apply(
        lambda v: str(v).strip() not in ('','nan','None','REVISAR','0'))
    qtd_nf  = int(nf_val.sum())
    ticket  = df_all_mes['VL TITULO'].sum() / qtd_nf if qtd_nf > 0 else 0

    diff_prazo   = (df_av['_paga'] - df_av['EMISSAO']).dt.days.dropna()
    prazo_medio  = diff_prazo.mean() if len(diff_prazo) > 0 else 0
    if np.isnan(prazo_medio): prazo_medio = 0

    # ════════════════════════════════════════════════════════════════
    # TÍTULO
    # ════════════════════════════════════════════════════════════════
    ws.set_row(0, 44); ws.set_row(1, 18)
    ws.merge_range('B1:J1',
        f'📊  DASHBOARD — CONTAS A RECEBER  |  {MES_ATUAL} / {ANO_ATUAL}', F_TITLE)
    ws.merge_range('B2:J2',
        f'Gerado em {TODAY.strftime("%d/%m/%Y")}  •  Período: {MES_ATUAL}/{ANO_ATUAL}  •  '
        f'A Vencer (mês): {len(df_av_mes):,} títulos  •  Vencidos (mês): {len(df_vd_mes):,} títulos',
        F_SUB)

    # ════════════════════════════════════════════════════════════════
    # SEÇÃO 1 — VISÃO EXECUTIVA
    # ════════════════════════════════════════════════════════════════
    ws.set_row(3, 22)
    ws.merge_range('B4:J4', f'  1. VISÃO EXECUTIVA — {MES_ATUAL}/{ANO_ATUAL}', F_SEC)

    # 5 KPIs em colunas B D F H J (cols 1 3 5 7 9)
    kpi_cols = [1, 3, 5, 7, 9]
    kpi_data = [
        (f'A VENCER — {MES_ATUAL} (R$)',           av_mes,      F_KPI_VAL, F_KPI_LBL),
        (f'VENCIDO — {MES_ATUAL} (R$)',             vencido_mes, F_KPI_RED, F_KPI_LBL_R),
        (f'% INADIMPLÊNCIA — {MES_ATUAL}',          pct_inadi,   F_KPI_PCT, F_KPI_LBL_R),
        ('TICKET MÉDIO POR NOTA FISCAL (R$)',       ticket,      F_KPI_VAL, F_KPI_LBL),
        ('PRAZO MÉDIO DE RECEBIMENTO (dias)',       prazo_medio, F_KPI_INT, F_KPI_LBL),
    ]
    ws.set_row(4, 18); ws.set_row(5, 40); ws.set_row(6, 40); ws.set_row(7, 18)
    for ci, (lbl, val, fval, flbl) in zip(kpi_cols, kpi_data):
        ws.write(4, ci, '', fval)
        ws.merge_range(5, ci, 6, ci, float(val), fval)
        ws.write(7, ci, lbl, flbl)

    # ════════════════════════════════════════════════════════════════
    # SEÇÃO 3 — RANKING DE CLIENTES
    # ════════════════════════════════════════════════════════════════
    R = 9  # começa logo após os KPIs
    ws.set_row(R, 22)
    ws.merge_range(R, 1, R, 9,
        f'  3. RANKING DE CLIENTES — Concentração de Receita — {MES_ATUAL}/{ANO_ATUAL}', F_SEC)
    R += 1; ws.set_row(R, 22)

    top_cli   = df_all_mes.groupby('_cli')['VL SALDO'].sum().sort_values(ascending=False).head(10)
    total_cli = df_all_mes['VL SALDO'].sum()

    for ci, h in enumerate(['POSIÇÃO','CLIENTE','VALOR SALDO (R$)','PARTICIPAÇÃO %']):
        ws.write(R, 1+ci, h, F_TH)
    R += 1
    for i,(cli,vl) in enumerate(top_cli.items()):
        fa=F_TA if i%2==0 else F_TB; fm=F_MA if i%2==0 else F_MB
        fp=F_PA if i%2==0 else F_PB; fr=F_RKA if i%2==0 else F_RKB
        ws.write(R+i, 1, i+1, fr)
        ws.write(R+i, 2, cli, fa)
        ws.write_number(R+i, 3, float(vl), fm)
        ws.write_number(R+i, 4, vl/total_cli if total_cli>0 else 0, fp)
    n = len(top_cli)
    ws.write(R+n, 1, 'TOTAL TOP 10', F_TOTL); ws.write(R+n, 2, '', F_TOTL)
    ws.write_number(R+n, 3, float(top_cli.sum()), F_TOT)
    ws.write_number(R+n, 4, top_cli.sum()/total_cli if total_cli>0 else 0, F_TOTP)
    R = R + n + 2

    # ════════════════════════════════════════════════════════════════
    # SEÇÃO 4 — RECEITA POR PRODUTO
    # ════════════════════════════════════════════════════════════════
    ws.set_row(R, 22)
    ws.merge_range(R, 1, R, 9,
        f'  4. RECEITA POR PRODUTO — {MES_ATUAL}/{ANO_ATUAL}', F_SEC)
    R += 1; ws.set_row(R, 22)

    prod_ser  = df_all_mes.groupby('_prod')['VL TITULO'].sum().sort_values(ascending=False).head(10)
    total_tit = df_all_mes['VL TITULO'].sum()

    for ci, h in enumerate(['PRODUTO','VALOR TÍTULO (R$)','PARTICIPAÇÃO %']):
        ws.write(R, 1+ci, h, F_TH)
    R += 1
    for i,(p,v) in enumerate(prod_ser.items()):
        fa=F_TA if i%2==0 else F_TB; fm=F_MA if i%2==0 else F_MB; fp=F_PA if i%2==0 else F_PB
        ws.write(R+i, 1, p, fa)
        ws.write_number(R+i, 2, float(v), fm)
        ws.write_number(R+i, 3, v/total_tit if total_tit>0 else 0, fp)
    n = len(prod_ser)
    ws.write(R+n, 1, 'TOTAL', F_TOTL); ws.write_number(R+n, 2, float(prod_ser.sum()), F_TOT)
    ws.write(R+n, 3, '', F_TOTP)
    R = R + n + 2

    # ════════════════════════════════════════════════════════════════
    # SEÇÃO 5 — SEGMENTO (GRUPO & TIPO) — lado a lado
    # ════════════════════════════════════════════════════════════════
    ws.set_row(R, 22)
    ws.merge_range(R, 1, R, 4,
        f'  5. RECEITA POR GRUPO — {MES_ATUAL}/{ANO_ATUAL}', F_SEC)
    ws.merge_range(R, 6, R, 9,
        f'  5b. RECEITA POR TIPO — {MES_ATUAL}/{ANO_ATUAL}', F_SEC)
    R += 1; ws.set_row(R, 22)

    grp_ser   = df_all_mes.groupby('_grp')['VL SALDO'].sum().sort_values(ascending=False).head(8)
    tipo_ser  = df_all_mes.groupby('_tipo')['VL SALDO'].sum().sort_values(ascending=False)
    total_sal = df_all_mes['VL SALDO'].sum()

    for ci, h in enumerate(['GRUPO','VALOR SALDO (R$)','PARTICIPAÇÃO %']):
        ws.write(R, 1+ci, h, F_TH)
    for ci, h in enumerate(['TIPO','VALOR SALDO (R$)','PARTICIPAÇÃO %']):
        ws.write(R, 6+ci, h, F_TH)
    R += 1

    for i,(g,v) in enumerate(grp_ser.items()):
        fa=F_TA if i%2==0 else F_TB; fm=F_MA if i%2==0 else F_MB; fp=F_PA if i%2==0 else F_PB
        ws.write(R+i,1,g,fa); ws.write_number(R+i,2,float(v),fm); ws.write_number(R+i,3,v/total_sal if total_sal>0 else 0,fp)
    n_g = len(grp_ser)
    ws.write(R+n_g,1,'TOTAL',F_TOTL); ws.write_number(R+n_g,2,float(grp_ser.sum()),F_TOT); ws.write(R+n_g,3,'',F_TOTP)

    for i,(t,v) in enumerate(tipo_ser.items()):
        fa=F_TA if i%2==0 else F_TB; fm=F_MA if i%2==0 else F_MB; fp=F_PA if i%2==0 else F_PB
        ws.write(R+i,6,t,fa); ws.write_number(R+i,7,float(v),fm); ws.write_number(R+i,8,v/total_sal if total_sal>0 else 0,fp)
    n_t = len(tipo_ser)
    ws.write(R+n_t,6,'TOTAL',F_TOTL); ws.write_number(R+n_t,7,float(tipo_ser.sum()),F_TOT); ws.write(R+n_t,8,'',F_TOTP)

    R = R + max(n_g, n_t) + 3

    # ════════════════════════════════════════════════════════════════
    # SEÇÃO 6 — RECEITA POR ESTADO (UF)
    # ════════════════════════════════════════════════════════════════
    ws.set_row(R, 22)
    ws.merge_range(R, 1, R, 9,
        f'  6. RECEITA POR ESTADO (UF) — {MES_ATUAL}/{ANO_ATUAL}', F_SEC)
    R += 1; ws.set_row(R, 22)

    uf_ser   = (df_all_mes[df_all_mes['_uf']!='N/D']
                .groupby('_uf')['VL SALDO'].sum()
                .sort_values(ascending=False).head(12))
    total_uf = uf_ser.sum()
    uf_list  = list(uf_ser.items())
    mid      = len(uf_list)//2 + len(uf_list)%2

    for ci, h in enumerate(['ESTADO (UF)','VALOR SALDO (R$)','PARTICIPAÇÃO %']):
        ws.write(R, 1+ci, h, F_TH)
    for ci, h in enumerate(['ESTADO (UF)','VALOR SALDO (R$)','PARTICIPAÇÃO %']):
        ws.write(R, 6+ci, h, F_TH)
    R += 1

    for i,(uf,v) in enumerate(uf_list[:mid]):
        fa=F_TA if i%2==0 else F_TB; fm=F_MA if i%2==0 else F_MB; fp=F_PA if i%2==0 else F_PB
        ws.write(R+i,1,uf,fa); ws.write_number(R+i,2,float(v),fm); ws.write_number(R+i,3,v/total_uf if total_uf>0 else 0,fp)
    for i,(uf,v) in enumerate(uf_list[mid:]):
        fa=F_TA if i%2==0 else F_TB; fm=F_MA if i%2==0 else F_MB; fp=F_PA if i%2==0 else F_PB
        ws.write(R+i,6,uf,fa); ws.write_number(R+i,7,float(v),fm); ws.write_number(R+i,8,v/total_uf if total_uf>0 else 0,fp)
    R = R + mid + 2

    # ════════════════════════════════════════════════════════════════
    # SEÇÃO 7 — TOP INADIMPLENTES & INADIMPLÊNCIA POR EXECUTIVO
    # ════════════════════════════════════════════════════════════════
    ws.set_row(R, 22)
    ws.merge_range(R, 1, R, 4,
        f'  7. TOP CLIENTES INADIMPLENTES — {MES_ATUAL}/{ANO_ATUAL}', F_SEC)
    ws.merge_range(R, 6, R, 9,
        f'  7b. INADIMPLÊNCIA POR EXECUTIVO — {MES_ATUAL}/{ANO_ATUAL}', F_SEC)
    R += 1; ws.set_row(R, 22)

    df_vd2 = df_vd_mes.copy()
    df_vd2['ATRASO'] = pd.to_numeric(df_vd2['ATRASO'], errors='coerce').fillna(0)
    df_vd2['_cli2']  = clean(df_vd2, 'CLIENTE')
    df_vd2['_exec2'] = clean(df_vd2, 'EXECUTIVO')

    cli_inadi  = df_vd2.groupby('_cli2').agg(
        vl=('VL SALDO','sum'), qtd=('VL SALDO','count'), atraso=('ATRASO','mean')
    ).sort_values('vl', ascending=False).head(10)
    exec_inadi = df_vd2.groupby('_exec2').agg(
        vl=('VL SALDO','sum'), atraso=('ATRASO','mean')
    ).sort_values('vl', ascending=False).head(10)

    for ci,h in enumerate(['POSIÇÃO','CLIENTE','VALOR VENCIDO (R$)','QUANTIDADE DE TÍTULOS','ATRASO MÉDIO (dias)']):
        ws.write(R, 1+ci, h, F_TH)
    for ci,h in enumerate(['POSIÇÃO','EXECUTIVO','VALOR VENCIDO (R$)','ATRASO MÉDIO (dias)']):
        ws.write(R, 6+ci, h, F_TH)
    R += 1

    n_ci = len(cli_inadi); n_ei = len(exec_inadi)
    for i,(cli,row) in enumerate(cli_inadi.iterrows()):
        fa=F_TA if i%2==0 else F_TB; fm=F_MA if i%2==0 else F_MB
        fi=F_IA if i%2==0 else F_IB; fr=F_RKA if i%2==0 else F_RKB
        ws.write(R+i,1,i+1,fr); ws.write(R+i,2,cli,fa)
        ws.write_number(R+i,3,float(row['vl']),fm)
        ws.write_number(R+i,4,int(row['qtd']),fi)
        ws.write_number(R+i,5,float(row['atraso']),fi)
    ws.write(R+n_ci,1,'TOTAL',F_TOTL); ws.write(R+n_ci,2,'',F_TOTL)
    ws.write_number(R+n_ci,3,float(cli_inadi['vl'].sum()),F_TOT)
    ws.write(R+n_ci,4,'',F_TOTL); ws.write(R+n_ci,5,'',F_TOTL)

    for i,(exc,row) in enumerate(exec_inadi.iterrows()):
        fa=F_TA if i%2==0 else F_TB; fm=F_MA if i%2==0 else F_MB
        fi=F_IA if i%2==0 else F_IB; fr=F_RKA if i%2==0 else F_RKB
        ws.write(R+i,6,i+1,fr); ws.write(R+i,7,exc,fa)
        ws.write_number(R+i,8,float(row['vl']),fm)
        ws.write_number(R+i,9,float(row['atraso']),fi)
    ws.write(R+n_ei,6,'TOTAL',F_TOTL); ws.write(R+n_ei,7,'',F_TOTL)
    ws.write_number(R+n_ei,8,float(exec_inadi['vl'].sum()),F_TOT)
    ws.write(R+n_ei,9,'',F_TOTL)

    R = R + max(n_ci, n_ei) + 3

    # ════════════════════════════════════════════════════════════════
    # SEÇÃO 8 — CONTROLE OPERACIONAL
    # ════════════════════════════════════════════════════════════════
    ws.set_row(R, 22)
    ws.merge_range(R, 1, R, 9,
        '  8. CONTROLE OPERACIONAL — Alertas Automáticos', F_SEC)
    R += 1; ws.set_row(R, 22)

    df_all_full = pd.concat([df_av_raw.copy(), df_vd_raw.copy()], ignore_index=True)
    for col in ['VL TITULO','VL SALDO']:
        df_all_full[col] = pd.to_numeric(df_all_full[col], errors='coerce').fillna(0)

    nf_sem    = df_all_full['NF'].apply(lambda v: str(v).strip() in ('','nan','None','0')).sum()
    diverge   = ((df_all_full['VL TITULO'] - df_all_full['VL SALDO']).abs() > 0.01).sum()
    sem_prev  = df_av_raw['PREV PAGTO'].apply(
        lambda v: str(v).strip() in ('REVISAR','','nan','None')).sum() \
        if 'PREV PAGTO' in df_av_raw.columns else 0
    revisar   = df_all_full['GRUPO'].apply(lambda v: str(v).strip()=='REVISAR').sum() \
                if 'GRUPO' in df_all_full.columns else 0
    df_vd_f   = df_vd_raw.copy()
    df_vd_f['ATRASO'] = pd.to_numeric(df_vd_f['ATRASO'], errors='coerce').fillna(0)
    venc_sem  = df_vd_f['ATRASO'].apply(lambda v: v==0).sum()

    alertas = [
        ('📄 Notas fiscais sem número',                          nf_sem,
         F_ALRT_OK if nf_sem==0   else F_ALRT_ERR, '✅ OK' if nf_sem==0   else '❌ Verificar'),
        ('💰 Divergência entre Valor Título e Valor Saldo',      diverge,
         F_ALRT_OK if diverge==0  else F_ALRT_WRN, '✅ OK' if diverge==0  else '⚠️ Revisar'),
        ('📅 Títulos a vencer sem previsão de pagamento',        sem_prev,
         F_ALRT_OK if sem_prev==0 else F_ALRT_WRN, '✅ OK' if sem_prev==0 else '⚠️ Atualizar'),
        ('🔴 Clientes novos sem histórico (REVISAR)',             revisar,
         F_ALRT_OK if revisar==0  else F_ALRT_WRN, '✅ OK' if revisar==0  else '⚠️ Cadastrar'),
        ('⏳ Títulos vencidos sem atraso preenchido',             venc_sem,
         F_ALRT_OK if venc_sem==0 else F_ALRT_WRN, '✅ OK' if venc_sem==0 else '⚠️ Atualizar'),
    ]
    for ci, h in enumerate(['DESCRIÇÃO DO ALERTA','QUANTIDADE','STATUS']):
        ws.write(R, 1+ci, h, F_TH)
    R += 1
    for i,(lbl,qtd,fmt_s,status) in enumerate(alertas):
        ws.set_row(R+i, 20)
        ws.write(R+i, 1, lbl, F_ALRT_LBL)
        ws.write_number(R+i, 2, int(qtd), F_ALRT_VAL)
        ws.write(R+i, 3, status, fmt_s)

    ws.freeze_panes(2, 0)

    print(f"   ✓ DASHBOARD {MES_ATUAL}/{ANO_ATUAL} · "
          f"A Vencer: R$ {av_mes:,.2f} · "
          f"Vencido: R$ {vencido_mes:,.2f} · "
          f"Inadimplência: {pct_inadi:.1%}")
