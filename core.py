import io
import unicodedata
import warnings
from datetime import datetime, timedelta

import pandas as pd
import numpy as np
import xlsxwriter
try:
    import holidays as hol_lib
except ImportError:
    hol_lib = None

warnings.filterwarnings("ignore")

# ════════════════════════════════════════════════════════════════════
#  FUNÇÕES UTILITÁRIAS
# ════════════════════════════════════════════════════════════════════

def normalize_str(val):
    """Maiúsculo + remove acentos + strip."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ''
    s = str(val).upper().strip()
    s = unicodedata.normalize('NFKD', s)
    return ''.join(c for c in s if not unicodedata.combining(c))

def to_upper(val):
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return val
    return str(val).upper().strip() if isinstance(val, str) else val

_hol_cache = {}
def br_holidays(year):
    if year not in _hol_cache:
        try:
            _hol_cache[year] = hol_lib.Brazil(years=year)
        except Exception:
            _hol_cache[year] = {}
    return _hol_cache[year]

def is_fri_or_holiday(ts):
    if ts is None or pd.isna(ts):
        return False
    try:
        dt = ts.to_pydatetime() if isinstance(ts, pd.Timestamp) else ts
        return dt.weekday() == 4 or dt.date() in br_holidays(dt.year)
    except Exception:
        return False

def next_util(ts):
    if ts is None or pd.isna(ts):
        return pd.NaT
    try:
        dt = ts.to_pydatetime() if isinstance(ts, pd.Timestamp) else ts
        d = dt + timedelta(days=1)
        for _ in range(90):
            if d.weekday() < 5 and d.date() not in br_holidays(d.year):
                return pd.Timestamp(d)
            d += timedelta(days=1)
        return pd.Timestamp(d)
    except Exception:
        return pd.NaT

def calc_prev(venc, days):
    if pd.isna(venc) or pd.isna(days):
        return pd.NaT
    try:
        dt  = venc.to_pydatetime() if isinstance(venc, pd.Timestamp) else venc
        res = pd.Timestamp(dt + timedelta(days=int(float(days))))
        return next_util(res) if is_fri_or_holiday(res) else res
    except Exception:
        return pd.NaT

def safe_rbase(val):
    """Normaliza RBASE para string de inteiro (chave do dicionário)."""
    try:
        if val is None or (isinstance(val, float) and pd.isna(val)):
            return ''
        return str(int(float(str(val).replace(',', '.').strip())))
    except Exception:
        return str(val).strip()

# ════════════════════════════════════════════════════════════════════
#  PROCESSAMENTO PRINCIPAL
# ════════════════════════════════════════════════════════════════════

def process_files(csv_bytes, xlsx_bytes):
    TODAY = pd.Timestamp(datetime.now().date())
    print(f"\n📅 Data de hoje: {TODAY.strftime('%d/%m/%Y')}")

    # ── 1. LER CRMX.CSV ─────────────────────────────────────────────
    print("\n📊 Lendo CRMX.CSV ...")
    df_raw = pd.read_csv(
        io.BytesIO(csv_bytes),
        sep=';', dtype=str, encoding='utf-8-sig',
        keep_default_na=False
    )
    df_raw.columns = [c.strip() for c in df_raw.columns]
    df_raw = df_raw[
        [c for c in df_raw.columns
         if not c.startswith('Unnamed')
         and not df_raw[c].str.strip().eq('').all()]
    ]
    N = len(df_raw)
    print(f"   ✓ {N:,} registros · {len(df_raw.columns)} colunas")

    def orig(letter):
        idx = ord(letter.upper()) - ord('A')
        if 0 <= idx < len(df_raw.columns):
            return df_raw.iloc[:, idx].copy()
        return pd.Series([''] * N, dtype=str)

    # ── 2. LER PLANILHA ANTIGA ───────────────────────────────────────
    #     Todos os campos espelhados por RBASE
    #     REVISAR somente para RBASEs que não existem na planilha antiga
    print("\n📊 Lendo CR_MAXIFROTA_2026.XLSX ...")
    rbase_to_grupo  = {}   # { '110710': 'PRIVADO BOLETO' }
    rbase_to_atraso = {}   # { '110710': 7 }
    rbase_to_paga   = {}   # { '110710': 'B' }  ← novo, por RBASE
    rbase_conhecido = set()  # todos os RBASEs vistos na planilha antiga
    old_prev        = {}   # { '377183': Timestamp } ← PREV PAGTO ainda por NF

    try:
        xf = pd.ExcelFile(io.BytesIO(xlsx_bytes))
        for sname in ['A vencer MX', 'Vencidos MX']:
            if sname not in xf.sheet_names:
                continue
            tmp = xf.parse(sname, header=None, dtype=str)
            hrow = None
            for i in range(min(5, len(tmp))):
                vals = [str(v).strip() for v in tmp.iloc[i]
                        if str(v).strip() not in ('', 'nan', 'None')]
                if 'RBASE' in vals and 'NF' in vals:
                    hrow = i
                    break
            if hrow is None:
                continue
            tmp.columns = [
                str(v).strip() if str(v).strip() not in ('', 'nan', 'None')
                else f'_col{j}'
                for j, v in enumerate(tmp.iloc[hrow])
            ]
            tmp = tmp.iloc[hrow + 1:].reset_index(drop=True)

            for _, r in tmp.iterrows():
                rb = str(r.get('RBASE', '')).strip()
                if not rb or rb in ('nan', 'None', ''):
                    continue
                rb_key = safe_rbase(rb)
                rbase_conhecido.add(rb_key)   # registrar RBASE como conhecido

                # GRUPO
                grp = str(r.get('GRUPO', '')).strip()
                if grp and grp not in ('nan', 'None', ''):
                    rbase_to_grupo[rb_key] = grp.upper()

                # ATRASO
                atr_v = str(r.get('ATRASO', '')).strip()
                if atr_v and atr_v not in ('nan', 'None', ''):
                    try:
                        rbase_to_atraso[rb_key] = abs(float(atr_v.replace(',', '.')))
                    except Exception:
                        pass

                # PAGA NA DATA por RBASE
                paga_v = str(r.get('PAGA NA DATA', '')).strip()
                if paga_v and paga_v not in ('nan', 'None', ''):
                    rbase_to_paga[rb_key] = paga_v.upper()

                # PREV PAGTO ainda por NF (data calculada é única por nota)
                nf_raw = str(r.get('NF', '')).strip()
                if nf_raw and nf_raw not in ('nan', 'None', ''):
                    try:
                        nf_key = str(int(float(nf_raw.replace(',', '.'))))
                    except Exception:
                        nf_key = nf_raw
                    for col_prev in ['PREV PAGTO', 'PREV PGTO']:
                        prev_v = str(r.get(col_prev, '')).strip()
                        if prev_v and prev_v not in ('nan', 'None', ''):
                            try:
                                old_prev[nf_key] = pd.to_datetime(
                                    prev_v, dayfirst=True, errors='coerce')
                            except Exception:
                                pass
                            break

        print(f"   ✓ {len(rbase_conhecido):,} RBASEs conhecidos na planilha antiga")
        print(f"   ✓ {len(rbase_to_grupo):,} com GRUPO · "
              f"{len(rbase_to_atraso):,} com ATRASO · "
              f"{len(rbase_to_paga):,} com PAGA NA DATA")
    except Exception as e:
        print(f"   ⚠️  Aviso ao ler planilha antiga: {e}")

    # ── 3. CONSTRUIR NOVO DATAFRAME ──────────────────────────────────
    print("\n🔄 Reorganizando colunas ...")
    out = pd.DataFrame({
        'EMPRESA'            : orig('T'),
        'EXECUTIVO'          : orig('O'),
        'PRODUTO'            : orig('G'),
        'CNPJ'               : orig('D'),
        'RBASE'              : orig('A'),
        'NF'                 : orig('S'),
        'NUM DOC'            : orig('J'),
        'CLIENTE'            : orig('B'),
        'CLIENTE 2'          : orig('C'),
        'UF'                 : orig('E'),
        'CIDADE'             : orig('Y'),
        'TIPO'               : orig('F'),
        'GRUPO'         : pd.Series([''] * N, dtype=str),
        'EMISSAO'            : pd.to_datetime(
                                   orig('H').str.strip(),
                                   errors='coerce', dayfirst=False),
        'VENCIMENTO'         : pd.to_datetime(
                                   orig('I').str.strip(),
                                   errors='coerce', dayfirst=False),
        # Valores: converter vírgula → ponto para leitura numérica
        'VL TITULO'          : pd.to_numeric(
                                   orig('L').str.replace(',', '.', regex=False),
                                   errors='coerce'),
        'VL SALDO'           : pd.to_numeric(
                                   orig('K').str.replace(',', '.', regex=False),
                                   errors='coerce'),
        'IR'                 : pd.to_numeric(
                                   orig('M').str.replace(',', '.', regex=False),
                                   errors='coerce'),
        'ISS'                : pd.to_numeric(
                                   orig('N').str.replace(',', '.', regex=False),
                                   errors='coerce'),
        'PAGA NA DATA'       : pd.Series([''] * N, dtype=str),
        # ATRASO: espelhado da planilha anterior por NF (preenchido abaixo)
        'ATRASO'             : pd.Series([np.nan] * N, dtype=float),
        'PREV PAGTO'         : pd.NaT,
        'BANCO'              : pd.Series([''] * N, dtype=str),
        'LINK NFSE'          : orig('W'),
        'COND PAGTO'         : orig('R'),
        'CODIGO DE PAGAMENTO': orig('V'),
    })
    out['PREV PAGTO'] = pd.to_datetime(out['PREV PAGTO'], errors='coerce')

    # ── 4. ESPELHAR CAMPOS DA PLANILHA ANTIGA (por RBASE) ────────────
    # Regra: REVISAR somente se o RBASE não existe na planilha antiga.
    # Se o RBASE existe mas o campo está vazio → célula fica em branco.
    print("⚙️  Espelhando campos por RBASE ...")

    def rbase_lookup(rb, mapping, default_known='', default_unknown='REVISAR'):
        key = safe_rbase(rb)
        if key in rbase_conhecido:
            return mapping.get(key, default_known)
        return default_unknown

    out['GRUPO']        = out['RBASE'].apply(lambda rb: rbase_lookup(rb, rbase_to_grupo))
    out['PAGA NA DATA'] = out['RBASE'].apply(lambda rb: rbase_lookup(rb, rbase_to_paga))
    out['ATRASO']       = out['RBASE'].apply(
        lambda rb: rbase_lookup(rb, rbase_to_atraso, default_known=np.nan))

    # PREV PAGTO: data é única por NF; REVISAR apenas se RBASE desconhecido
    def lookup_nf_key(nf_raw):
        try:
            return str(int(float(str(nf_raw).replace(',', '.').strip())))
        except Exception:
            return str(nf_raw).strip()

    def prev_lookup(row):
        if safe_rbase(row['RBASE']) not in rbase_conhecido:
            return 'REVISAR'
        return old_prev.get(lookup_nf_key(row['NF']), pd.NaT)

    out['PREV PAGTO'] = out[['RBASE', 'NF']].apply(prev_lookup, axis=1)

    print(f"   ✓ GRUPO       → {(out['GRUPO']       != 'REVISAR').sum():,} ok · {(out['GRUPO']       == 'REVISAR').sum():,} REVISAR")
    print(f"   ✓ PAGA NA DATA→ {(out['PAGA NA DATA'] != 'REVISAR').sum():,} ok · {(out['PAGA NA DATA'] == 'REVISAR').sum():,} REVISAR")
    print(f"   ✓ ATRASO      → {out['ATRASO'].apply(lambda v: v != 'REVISAR' and pd.notna(v)).sum():,} ok · {(out['ATRASO'] == 'REVISAR').sum():,} REVISAR")
    print(f"   ✓ PREV PAGTO  → {(out['PREV PAGTO'] != 'REVISAR').sum():,} ok · {(out['PREV PAGTO'] == 'REVISAR').sum():,} REVISAR")

    # ── 6. TEXTO PARA MAIÚSCULAS ─────────────────────────────────────
    STR_COLS = [
        'EMPRESA', 'EXECUTIVO', 'PRODUTO', 'CNPJ', 'CLIENTE', 'CLIENTE 2',
        'UF', 'CIDADE', 'TIPO', 'GRUPO', 'BANCO', 'COND PAGTO',
        'CODIGO DE PAGAMENTO', 'PAGA NA DATA', 'NUM DOC', 'LINK NFSE'
    ]
    for c in STR_COLS:
        if c in out.columns:
            out[c] = out[c].apply(to_upper)

    # ── 7. SEPARAR A VENCER / VENCIDOS ──────────────────────────────
    mask_venc = out['VENCIMENTO'].notna() & (out['VENCIMENTO'] < TODAY)
    df_av = out[~mask_venc].copy().reset_index(drop=True)
    df_vd = out[ mask_venc].copy().reset_index(drop=True)

    print(f"\n📋 A VENCER : {len(df_av):,} registros")
    print(f"📋 VENCIDOS : {len(df_vd):,} registros")

    # Regra 25 – remover PAGA NA DATA e UF da aba Vencidos
    df_vd = df_vd.drop(columns=['PAGA NA DATA', 'UF'], errors='ignore')

    # ── 8. GERAR EXCEL ───────────────────────────────────────────────
    print("\n📝 Gerando Excel formatado ...")
    output_buf = io.BytesIO()
    wb = xlsxwriter.Workbook(output_buf, {'in_memory': True,
                                           'default_date_format': 'dd/mm/yyyy'})

    # ── Formatos ──────────────────────────────────────────────────────
    F_HDR  = wb.add_format({
        'bold': True, 'bg_color': '#00205B', 'font_color': '#FFFFFF',
        'border': 1, 'align': 'center', 'valign': 'vcenter',
        'text_wrap': True, 'font_name': 'Arial', 'font_size': 10
    })
    F_DATE = wb.add_format({'num_format': 'dd/mm/yyyy',
                             'align': 'center', 'font_name': 'Arial', 'font_size': 10})
    # Formato numérico padrão BR: ponto como milhar, vírgula como decimal
    # O Excel renderiza conforme o locale do sistema (PT-BR)
    F_MON  = wb.add_format({'num_format': '#,##0.00',
                             'align': 'right', 'font_name': 'Arial', 'font_size': 10})
    F_NF   = wb.add_format({'num_format': '000000',
                             'align': 'center', 'font_name': 'Arial', 'font_size': 10})
    F_INT  = wb.add_format({'num_format': '0',
                             'align': 'center', 'font_name': 'Arial', 'font_size': 10})
    F_CTR  = wb.add_format({'align': 'center', 'font_name': 'Arial', 'font_size': 10})
    F_TXT  = wb.add_format({'align': 'left',   'font_name': 'Arial', 'font_size': 10})
    F_LNK  = wb.add_format({'align': 'left',   'font_name': 'Arial', 'font_size': 9,
                             'font_color': '#0563C1', 'underline': 1})
    F_REV  = wb.add_format({'bold': True, 'font_color': '#FF0000',
                             'align': 'center', 'font_name': 'Arial', 'font_size': 10})

    DATE_COLS = {'EMISSAO', 'VENCIMENTO', 'PREV PAGTO'}
    MON_COLS  = {'VL TITULO', 'VL SALDO', 'IR', 'ISS'}
    NF_COLS   = {'NF'}
    INT_COLS  = {'ATRASO', 'RBASE'}
    CTR_COLS  = {'UF', 'TIPO', 'PRODUTO', 'GRUPO', 'PAGA NA DATA',
                 'CODIGO DE PAGAMENTO', 'BANCO'}

    COL_WIDTHS = {
        'EMPRESA': 18, 'EXECUTIVO': 28, 'PRODUTO': 10, 'CNPJ': 22,
        'RBASE': 10,   'NF': 10,        'NUM DOC': 18, 'CLIENTE': 44,
        'CLIENTE 2': 44, 'UF': 6,       'CIDADE': 24,  'TIPO': 12,
        'GRUPO': 24, 'EMISSAO': 14, 'VENCIMENTO': 14,
        'VL TITULO': 16, 'VL SALDO': 16, 'IR': 12,      'ISS': 12,
        'PAGA NA DATA': 14, 'ATRASO': 10, 'PREV PAGTO': 14,
        'BANCO': 14,   'LINK NFSE': 62, 'COND PAGTO': 16,
        'CODIGO DE PAGAMENTO': 24,
    }

    def write_tab(df_tab, sname):
        ws   = wb.add_worksheet(sname)
        cols = list(df_tab.columns)
        nr   = len(df_tab)
        nc   = len(cols)

        # Cabeçalho
        ws.set_row(0, 32)
        for ci, cn in enumerate(cols):
            ws.write(0, ci, cn, F_HDR)

        # Dados
        for ri in range(nr):
            row_data = df_tab.iloc[ri]
            xrow     = ri + 1

            for ci, cn in enumerate(cols):
                val  = row_data[cn]
                miss = (
                    val is None or
                    (isinstance(val, float) and pd.isna(val)) or
                    str(val).strip() in ('nan', 'None', '')
                )

                # ── REVISAR em qualquer coluna → vermelho negrito ─
                if str(val).strip() == 'REVISAR':
                    ws.write(xrow, ci, 'REVISAR', F_REV)

                # ── Datas ─────────────────────────────────────────
                elif cn in DATE_COLS:
                    is_valid = (
                        not miss and
                        isinstance(val, (pd.Timestamp, datetime)) and
                        not (isinstance(val, pd.Timestamp) and pd.isna(val))
                    )
                    if is_valid:
                        try:
                            dt = (val.to_pydatetime()
                                  if isinstance(val, pd.Timestamp) else val)
                            ws.write_datetime(xrow, ci, dt, F_DATE)
                        except Exception:
                            ws.write(xrow, ci, '', F_DATE)
                    else:
                        ws.write(xrow, ci, '', F_DATE)

                # ── Valores monetários (padrão BR) ───────────────
                elif cn in MON_COLS:
                    if not miss:
                        try:
                            ws.write_number(xrow, ci, float(val), F_MON)
                        except Exception:
                            ws.write(xrow, ci, '', F_TXT)
                    else:
                        ws.write(xrow, ci, '', F_MON)

                # ── NF – formato 000000 ──────────────────────────
                elif cn in NF_COLS:
                    raw = str(val).strip()
                    if raw and raw not in ('nan', 'None', ''):
                        try:
                            ws.write_number(xrow, ci,
                                            int(float(raw.replace(',', '.'))),
                                            F_NF)
                        except Exception:
                            ws.write(xrow, ci, raw, F_TXT)
                    else:
                        ws.write(xrow, ci, '', F_TXT)

                # ── Inteiros (ATRASO, RBASE) ─────────────────────
                elif cn in INT_COLS:
                    if not miss:
                        try:
                            ws.write_number(xrow, ci, float(val), F_INT)
                        except Exception:
                            ws.write(xrow, ci, '', F_TXT)
                    else:
                        ws.write(xrow, ci, '', F_INT)

                # ── Link NFSE ────────────────────────────────────
                elif cn == 'LINK NFSE':
                    s = '' if miss else str(val).strip()
                    if s.startswith('http'):
                        try:
                            ws.write_url(xrow, ci, s, F_LNK, s[:255])
                        except Exception:
                            ws.write(xrow, ci, s, F_LNK)
                    else:
                        ws.write(xrow, ci, s, F_TXT)

                # ── Centralizado ─────────────────────────────────
                elif cn in CTR_COLS:
                    ws.write(xrow, ci,
                             '' if miss else str(val).strip(), F_CTR)

                # ── Texto geral ──────────────────────────────────
                else:
                    ws.write(xrow, ci,
                             '' if miss else str(val).strip(), F_TXT)

        # Larguras
        for ci, cn in enumerate(cols):
            w = COL_WIDTHS.get(cn)
            if w is None:
                try:
                    ml = int(df_tab[cn].astype(str).str.len().max() or 10)
                except Exception:
                    ml = 10
                w = min(max(len(cn), ml) + 3, 55)
            ws.set_column(ci, ci, w)

        ws.autofilter(0, 0, nr, nc - 1)
        ws.freeze_panes(1, 0)


    # ── 9. DASHBOARDS (RESUMO A VENCER / RESUMO VENCIDOS) ───────────
    print("\n📊 Gerando dashboards ...")

    def write_dashboard(ws_name, df_src, date_col, cor_hdr, cor_k1, cor_k2, tipo):
        ws = wb.add_worksheet(ws_name)
        ws.set_zoom(85)
        ws.hide_gridlines(2)

        def fmt(**kw):
            return wb.add_format({**{'font_name':'Arial','font_size':10}, **kw})

        F_TITLE = fmt(bold=True,font_size=18,font_color=cor_hdr,valign='vcenter')
        F_SUB   = fmt(italic=True,font_size=10,font_color='#666666')
        F_KN    = fmt(bold=True,font_size=22,font_color=cor_hdr,align='center',valign='vcenter',num_format='#,##0.00')
        F_KN2   = fmt(bold=True,font_size=22,font_color=cor_hdr,align='center',valign='vcenter',num_format='#,##0')
        F_KL    = fmt(font_size=9,font_color='#888888',align='center',valign='vcenter',text_wrap=True)
        F_KB    = fmt(bg_color='#F0F4FA',border=1,border_color='#D0DCF0',align='center',valign='vcenter')
        F_KB2   = fmt(bg_color='#FFF8EC',border=1,border_color='#F5C842',align='center',valign='vcenter')
        F_HDR   = fmt(bold=True,bg_color=cor_hdr,font_color='#FFFFFF',border=1,align='center',valign='vcenter',text_wrap=True)
        F_HDR2  = fmt(bold=True,bg_color='#E8EEF8',font_color=cor_hdr,border=1,align='center',valign='vcenter')
        F_RA    = fmt(bg_color='#FFFFFF',border=1,border_color='#E0E8F0')
        F_RB    = fmt(bg_color='#F7FAFF',border=1,border_color='#E0E8F0')
        F_MA    = fmt(bg_color='#FFFFFF',border=1,border_color='#E0E8F0',num_format='#,##0.00',align='right')
        F_MB    = fmt(bg_color='#F7FAFF',border=1,border_color='#E0E8F0',num_format='#,##0.00',align='right')
        F_RKA   = fmt(bold=True,bg_color='#FFFFFF',border=1,border_color='#E0E8F0',align='center')
        F_RKB   = fmt(bold=True,bg_color='#F7FAFF',border=1,border_color='#E0E8F0',align='center')
        F_TOT   = fmt(bold=True,bg_color='#E8EEF8',border=1,border_color='#C0CCE0',num_format='#,##0.00',align='right')
        F_TOTL  = fmt(bold=True,bg_color='#E8EEF8',border=1,border_color='#C0CCE0')
        F_AG1   = fmt(bold=True,bg_color='#FFF2CC',border=1,align='center',valign='vcenter',num_format='#,##0.00')
        F_AG2   = fmt(bold=True,bg_color='#FFE0B2',border=1,align='center',valign='vcenter',num_format='#,##0.00')
        F_AG3   = fmt(bold=True,bg_color='#FFCDD2',border=1,align='center',valign='vcenter',num_format='#,##0.00')
        F_AG4   = fmt(bold=True,bg_color='#B71C1C',font_color='#FFFFFF',border=1,align='center',valign='vcenter',num_format='#,##0.00')
        MES_NOME= {1:'Janeiro',2:'Fevereiro',3:'Março',4:'Abril',5:'Maio',6:'Junho',
                   7:'Julho',8:'Agosto',9:'Setembro',10:'Outubro',11:'Novembro',12:'Dezembro'}

        df = df_src.copy()
        df[date_col] = pd.to_datetime(df[date_col],errors='coerce')
        df['_vl'] = pd.to_numeric(df['VL SALDO'],errors='coerce').fillna(0)
        df['_cli']= df['CLIENTE'].apply(lambda v:str(v).strip() if str(v).strip() not in ('','nan','None') else 'N/D')
        df['_grp']= df['GRUPO'].apply(lambda v:str(v).strip() if str(v).strip() not in ('','nan','None','REVISAR') else 'SEM GRUPO')
        df['_exc']= df['EXECUTIVO'].apply(lambda v:str(v).strip() if str(v).strip() not in ('','nan','None') else 'N/D')
        df_m = df[df[date_col].notna()&(df[date_col].dt.month==TODAY.month)&(df[date_col].dt.year==TODAY.year)].copy()
        if df_m.empty: df_m = df.copy()

        total_vl = df['_vl'].sum(); qtd = len(df)
        ticket   = total_vl/qtd if qtd>0 else 0
        top_grp  = df.groupby('_grp')['_vl'].sum().idxmax() if qtd>0 else 'N/D'
        top_grp_vl=df.groupby('_grp')['_vl'].sum().max() if qtd>0 else 0

        ws.set_column('A:A',3); ws.set_column('B:C',22); ws.set_column('D:D',16)
        ws.set_column('E:E',3); ws.set_column('F:G',22); ws.set_column('H:H',16); ws.set_column('I:I',3)

        ws.set_row(0,40); ws.set_row(1,18)
        lbl = 'A VENCER' if tipo=='av' else 'VENCIDOS'
        ws.merge_range('B1:H1',f'DASHBOARD — {lbl} | {MES_NOME[TODAY.month]}/{TODAY.year}',F_TITLE)
        ws.merge_range('B2:H2',f'Gerado em {TODAY.strftime("%d/%m/%Y")}  •  {qtd:,} títulos  •  Total: R$ {total_vl:,.2f}',F_SUB)

        K = 3
        for r in range(K,K+5): ws.set_row(r,26)
        ws.merge_range(K,1,K,2,'',F_KB);ws.merge_range(K+1,1,K+2,2,total_vl,F_KN)
        ws.merge_range(K+3,1,K+3,2,'💰 Total (R$)',F_KL);ws.merge_range(K+4,1,K+4,2,'',F_KB)
        ws.write(K,3,'',F_KB);ws.merge_range(K+1,3,K+2,3,qtd,F_KN2)
        ws.write(K+3,3,'📄 Títulos',F_KL);ws.write(K+4,3,'',F_KB)
        ws.merge_range(K,5,K,6,'',F_KB2);ws.merge_range(K+1,5,K+2,6,ticket,F_KN)
        ws.merge_range(K+3,5,K+3,6,'📊 Ticket Médio',F_KL);ws.merge_range(K+4,5,K+4,6,'',F_KB2)
        ws.write(K,7,'',F_KB2)
        ws.merge_range(K+1,7,K+2,7,top_grp[:18],fmt(bold=True,font_size=12,font_color=cor_hdr,align='center',valign='vcenter',text_wrap=True))
        ws.write(K+3,7,f'🏆 Maior Grupo\nR$ {top_grp_vl:,.0f}',fmt(font_size=8,font_color='#888888',align='center',text_wrap=True))
        ws.write(K+4,7,'',F_KB2)

        GR = 42
        for i in range(15): ws.set_row(GR+i,0)
        ws.write(GR,1,'GRUPO'); ws.write(GR,2,'VALOR')
        grp_s = df.groupby('_grp')['_vl'].sum().sort_values(ascending=False).head(10)
        for i,(g,v) in enumerate(grp_s.items()):
            ws.write(GR+1+i,1,g[:30]); ws.write_number(GR+1+i,2,float(v))

        DR = GR+15
        for i in range(25): ws.set_row(DR+i,0)
        ws.write(DR,5,'DIA'); ws.write(DR,6,'VALOR')
        dia_s = df_m.groupby(df_m[date_col].dt.day)['_vl'].sum().sort_index().head(25) if not df_m.empty else pd.Series([],dtype=float)
        for i,(d,v) in enumerate(dia_s.items()):
            ws.write(DR+1+i,5,f'{int(d):02d}/{TODAY.month:02d}'); ws.write_number(DR+1+i,6,float(v))

        c1 = wb.add_chart({'type':'bar'})
        c1.add_series({'name':'Por Grupo','categories':[ws_name,GR+1,1,GR+len(grp_s),1],
                       'values':[ws_name,GR+1,2,GR+len(grp_s),2],
                       'fill':{'color':cor_k1},'border':{'color':'#FFFFFF'},
                       'data_labels':{'value':True,'num_format':'#,##0','font':{'name':'Arial','size':8}}})
        c1.set_title({'name':'Valor por Grupo (Top 10)','name_font':{'bold':True,'size':11,'color':cor_hdr,'name':'Arial'}})
        c1.set_x_axis({'num_format':'#,##0','num_font':{'name':'Arial','size':8}})
        c1.set_y_axis({'num_font':{'name':'Arial','size':8}})
        c1.set_legend({'none':True}); c1.set_chartarea({'border':{'none':True},'fill':{'color':'#FAFCFF'}})
        c1.set_size({'width':380,'height':300}); ws.insert_chart('B10',c1)

        c2 = wb.add_chart({'type':'column'})
        if not dia_s.empty:
            c2.add_series({'name':'Por Dia','categories':[ws_name,DR+1,5,DR+len(dia_s),5],
                           'values':[ws_name,DR+1,6,DR+len(dia_s),6],
                           'fill':{'color':cor_k2},'border':{'color':'#FFFFFF'},
                           'data_labels':{'value':True,'num_format':'#,##0','font':{'name':'Arial','size':7}}})
        lbl2='Previsão por Dia' if tipo=='av' else 'Vencimentos por Dia'
        c2.set_title({'name':lbl2,'name_font':{'bold':True,'size':11,'color':cor_hdr,'name':'Arial'}})
        c2.set_x_axis({'num_font':{'name':'Arial','size':8}}); c2.set_y_axis({'num_format':'#,##0','num_font':{'name':'Arial','size':8}})
        c2.set_legend({'none':True}); c2.set_chartarea({'border':{'none':True},'fill':{'color':'#FAFCFF'}})
        c2.set_size({'width':380,'height':300}); ws.insert_chart('F10',c2)

        T1 = 30; ws.set_row(T1,28)
        ws.merge_range(T1,1,T1,3,'🏅 TOP 10 CLIENTES POR VALOR',F_HDR)
        ws.write(T1+1,1,'#',F_HDR2); ws.write(T1+1,2,'CLIENTE',F_HDR2); ws.write(T1+1,3,'VALOR (R$)',F_HDR2)
        top_cli = df.groupby('_cli')['_vl'].sum().sort_values(ascending=False).head(10)
        for i,(c,v) in enumerate(top_cli.items()):
            fa=F_RA if i%2==0 else F_RB; fm=F_MA if i%2==0 else F_MB; fr=F_RKA if i%2==0 else F_RKB
            ws.write(T1+2+i,1,i+1,fr); ws.write(T1+2+i,2,c[:28],fa); ws.write_number(T1+2+i,3,float(v),fm)
        ws.write(T1+12,1,'TOTAL',F_TOTL); ws.write(T1+12,2,'',F_TOTL); ws.write_number(T1+12,3,float(top_cli.sum()),F_TOT)

        ws.merge_range(T1,5,T1,7,'👤 RANKING EXECUTIVOS',F_HDR)
        ws.write(T1+1,5,'#',F_HDR2); ws.write(T1+1,6,'EXECUTIVO',F_HDR2); ws.write(T1+1,7,'VALOR (R$)',F_HDR2)
        exec_s = df.groupby('_exc')['_vl'].sum().sort_values(ascending=False).head(10)
        for i,(e,v) in enumerate(exec_s.items()):
            fa=F_RA if i%2==0 else F_RB; fm=F_MA if i%2==0 else F_MB; fr=F_RKA if i%2==0 else F_RKB
            ws.write(T1+2+i,5,i+1,fr); ws.write(T1+2+i,6,e[:28],fa); ws.write_number(T1+2+i,7,float(v),fm)
        ws.write(T1+12,5,'TOTAL',F_TOTL); ws.write(T1+12,6,'',F_TOTL); ws.write_number(T1+12,7,float(exec_s.sum()),F_TOT)

        S2 = T1+15; ws.set_row(S2,28)
        if tipo=='vd':
            ws.merge_range(S2,1,S2,7,'⏳ AGING — Distribuição por Prazo de Atraso',F_HDR)
            bins=[0,30,60,90,180,365,99999]; lbls=['1–30 dias','31–60 dias','61–90 dias','91–180 dias','181–365 dias','> 365 dias']
            cors=[F_AG1,F_AG1,F_AG2,F_AG3,F_AG3,F_AG4]
            df2=df.copy(); df2['da']=(TODAY-df2[date_col]).dt.days.clip(lower=0)
            tots=[]; cnts=[]
            for j in range(6):
                m=(df2['da']>bins[j])&(df2['da']<=bins[j+1]); tots.append(df2.loc[m,'_vl'].sum()); cnts.append(int(m.sum()))
            for j in range(6):
                ws.write(S2+1,1+j,lbls[j],fmt(bold=True,align='center',font_size=9,bg_color='#F0F4FA',border=1))
                ws.write_number(S2+2,1+j,float(tots[j]),cors[j])
                ws.write(S2+3,1+j,f'{cnts[j]} títulos',fmt(align='center',font_size=8,font_color='#666666',border=1))
                ws.set_column(1+j,1+j,16)
            ws.set_row(S2+2,36); ws.set_row(S2+3,18)
            AGD=S2+7
            for i in range(8): ws.set_row(AGD+i,0)
            ws.write(AGD,1,'FAIXA'); ws.write(AGD,2,'VALOR')
            for j,(l,t) in enumerate(zip(lbls,tots)):
                ws.write(AGD+1+j,1,l); ws.write_number(AGD+1+j,2,float(t))
            c3=wb.add_chart({'type':'column'})
            c3.add_series({'name':'Aging','categories':[ws_name,AGD+1,1,AGD+6,1],'values':[ws_name,AGD+1,2,AGD+6,2],
                           'fill':{'color':'#C00000'},'data_labels':{'value':True,'num_format':'#,##0','font':{'name':'Arial','size':8}}})
            c3.set_title({'name':'Aging — Valor por Faixa de Atraso','name_font':{'bold':True,'size':11,'color':cor_hdr,'name':'Arial'}})
            c3.set_x_axis({'num_font':{'name':'Arial','size':8}}); c3.set_y_axis({'num_format':'#,##0','num_font':{'name':'Arial','size':8}})
            c3.set_legend({'none':True}); c3.set_chartarea({'border':{'none':True},'fill':{'color':'#FAFCFF'}})
            c3.set_size({'width':780,'height':260}); ws.insert_chart(S2+4,1,c3,{'x_offset':0,'y_offset':5})
        else:
            ws.merge_range(S2,1,S2,7,'📅 PREVISÃO POR SEMANA — Próximas 4 Semanas',F_HDR)
            df2=df[df[date_col].notna()&(df[date_col]>=TODAY)].copy()
            if not df2.empty:
                df2['sem']=((df2[date_col]-TODAY).dt.days//7).clip(upper=3)
            else:
                df2['sem']=pd.Series([],dtype=int)
            slbls=['Esta semana','Semana 2','Semana 3','Semana 4+']; tots=[]; cnts=[]
            for s in range(4):
                m=df2['sem']==s if not df2.empty else pd.Series([],dtype=bool)
                tots.append(float(df2.loc[m,'_vl'].sum()) if not df2.empty and len(m)>0 else 0.0)
                cnts.append(int(m.sum()) if not df2.empty and len(m)>0 else 0)
            sc=[fmt(bold=True,bg_color=c,border=1,align='center',valign='vcenter',num_format='#,##0.00')
                for c in ['#C8E6C9','#B3E5FC','#FFF9C4','#F8BBD0']]
            for j in range(4):
                ws.write(S2+1,1+j*2,slbls[j],fmt(bold=True,align='center',font_size=9,bg_color='#F0F4FA',border=1))
                ws.write(S2+1,2+j*2,slbls[j],fmt(bold=True,align='center',font_size=9,bg_color='#F0F4FA',border=1))
                ws.merge_range(S2+2,1+j*2,S2+2,2+j*2,tots[j],sc[j])
                ws.merge_range(S2+3,1+j*2,S2+3,2+j*2,f'{cnts[j]} títulos',fmt(align='center',font_size=8,font_color='#666666',border=1))
            ws.set_row(S2+2,36); ws.set_row(S2+3,18)
            SD=S2+7
            for i in range(6): ws.set_row(SD+i,0)
            ws.write(SD,1,'SEM'); ws.write(SD,2,'VALOR')
            for j,(l,t) in enumerate(zip(slbls,tots)):
                ws.write(SD+1+j,1,l); ws.write_number(SD+1+j,2,float(t))
            c3=wb.add_chart({'type':'column'})
            c3.add_series({'name':'Semanas','categories':[ws_name,SD+1,1,SD+4,1],'values':[ws_name,SD+1,2,SD+4,2],
                           'fill':{'color':'#2E75B6'},'data_labels':{'value':True,'num_format':'#,##0','font':{'name':'Arial','size':9}}})
            c3.set_title({'name':'Previsão — Próximas Semanas','name_font':{'bold':True,'size':11,'color':cor_hdr,'name':'Arial'}})
            c3.set_x_axis({'num_font':{'name':'Arial','size':9}}); c3.set_y_axis({'num_format':'#,##0','num_font':{'name':'Arial','size':9}})
            c3.set_legend({'none':True}); c3.set_chartarea({'border':{'none':True},'fill':{'color':'#FAFCFF'}})
            c3.set_size({'width':780,'height':260}); ws.insert_chart(S2+4,1,c3,{'x_offset':0,'y_offset':5})

        print(f"   ✓ {ws_name}: {qtd:,} títulos · R$ {total_vl:,.2f}")

    # Preparar A VENCER
    df_av_db = df_av.copy()
    df_av_db['PREV PAGTO'] = df_av_db['PREV PAGTO'].apply(
        lambda v: pd.NaT if str(v).strip() in ('REVISAR','','nan','None') else v)
    df_av_db['PREV PAGTO'] = pd.to_datetime(df_av_db['PREV PAGTO'],errors='coerce')

    write_dashboard('RESUMO A VENCER', df_av_db, 'PREV PAGTO',
                    '#00205B','#2E75B6','#70AD47','av')
    write_dashboard('RESUMO VENCIDOS', df_vd, 'VENCIMENTO',
                    '#00205B','#C00000','#FF6B35','vd')

    write_tab(df_av, 'A VENCER')
    write_tab(df_vd, 'VENCIDOS')
    wb.close()

    output_buf.seek(0)
    return output_buf.getvalue()
