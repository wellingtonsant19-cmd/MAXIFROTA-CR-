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

warnings.filterwarnings('ignore')

# ════════════════════════════════════════════════════════════════════
#  FUNÇÕES UTILITÁRIAS
# ════════════════════════════════════════════════════════════════════

def to_upper(val):
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return val
    return str(val).upper().strip() if isinstance(val, str) else val

def safe_rbase(val):
    try:
        if val is None or (isinstance(val, float) and pd.isna(val)):
            return ''
        return str(int(float(str(val).replace(',', '.').strip())))
    except Exception:
        return str(val).strip()

_hol_cache = {}
def br_holidays(year):
    if year not in _hol_cache:
        try:
            _hol_cache[year] = hol_lib.Brazil(years=year) if hol_lib else {}
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
        d  = dt + timedelta(days=1)
        for _ in range(90):
            if d.weekday() < 5 and d.date() not in br_holidays(d.year):
                return pd.Timestamp(d)
            d += timedelta(days=1)
        return pd.Timestamp(d)
    except Exception:
        return pd.NaT

# ════════════════════════════════════════════════════════════════════
#  PROCESSAMENTO PRINCIPAL — NUTRICASH
# ════════════════════════════════════════════════════════════════════

def process_files_nutricash(csv_bytes, xlsx_bytes):
    TODAY = pd.Timestamp(datetime.now().date())
    print(f"\n📅 Data de hoje: {TODAY.strftime('%d/%m/%Y')}")

    # ── 1. LER CRNC.CSV ─────────────────────────────────────────────
    # Nutricash usa encoding latin1 e não tem colunas IR / ISS
    # Mapeamento das 24 colunas (índice 0–22):
    # [0]RBASE [1]NOME CLIENTE [2]NOME [3]CNPJ [4]UF [5]TIPO [6]PRODUTO
    # [7]DT EMISSAO [8]DT VENCIMENTO [9]NUM DOC [10]VLR SALDO [11]VLR TITULO
    # [12]EXECUTIVO [13]DATA BLOQUEIO [14]ATRASO [15]COND PAGTO [16]NR NFEM
    # [17]EMPRESA [18]SISTEMA [19]PORTADOR [20]LINK NFSE [21]RBASE RAIZ [22]CIDADE
    print("\n📊 Lendo CRNC.CSV ...")
    df_raw = pd.read_csv(
        io.BytesIO(csv_bytes),
        sep=';', dtype=str, encoding='latin1',
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

    # ── 2. LER PLANILHA ANTERIOR ─────────────────────────────────────
    # Abas: 'A vencer NC' e 'Vencidos NC'
    # RBASE col → 'RBASE' (a vencer) ou 'RRBASE' (vencidos, typo na planilha)
    # PAGA NA DATA col → 'PAGA NA DATA?'
    # PREV PAGTO col → 'PREV PGTO'
    print("\n📊 Lendo CR_NUTRICASH_2026.XLSX ...")
    rbase_to_grupo  = {}
    rbase_to_atraso = {}
    rbase_to_paga   = {}
    rbase_conhecido = set()
    old_prev        = {}

    try:
        xf = pd.ExcelFile(io.BytesIO(xlsx_bytes))
        for sname in ['A vencer NC', 'Vencidos NC']:
            if sname not in xf.sheet_names:
                continue
            tmp = xf.parse(sname, header=None, dtype=str)
            hrow = None
            for i in range(min(6, len(tmp))):
                vals = [str(v).strip() for v in tmp.iloc[i]
                        if str(v).strip() not in ('', 'nan', 'None')]
                if any(x in vals for x in ['RBASE', 'RRBASE']) and 'NF' in vals:
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

            # Normalizar nome da coluna RBASE (Vencidos NC tem 'RRBASE')
            if 'RRBASE' in tmp.columns and 'RBASE' not in tmp.columns:
                tmp = tmp.rename(columns={'RRBASE': 'RBASE'})

            for _, r in tmp.iterrows():
                rb = str(r.get('RBASE', '')).strip()
                if not rb or rb in ('nan', 'None', ''):
                    continue
                rk = safe_rbase(rb)
                rbase_conhecido.add(rk)

                # GRUPO
                grp = str(r.get('GRUPO', '')).strip()
                if grp and grp not in ('nan', 'None', ''):
                    rbase_to_grupo[rk] = grp.upper()

                # ATRASO
                atr_v = str(r.get('ATRASO', '')).strip()
                if atr_v and atr_v not in ('nan', 'None', ''):
                    try:
                        rbase_to_atraso[rk] = abs(float(atr_v.replace(',', '.')))
                    except Exception:
                        pass

                # PAGA NA DATA (coluna se chama 'PAGA NA DATA?' na Nutricash)
                for paga_col in ['PAGA NA DATA?', 'PAGA NA DATA']:
                    paga_v = str(r.get(paga_col, '')).strip()
                    if paga_v and paga_v not in ('nan', 'None', ''):
                        rbase_to_paga[rk] = paga_v.upper()
                        break

                # PREV PGTO por NF
                nf_raw = str(r.get('NF', '')).strip()
                if nf_raw and nf_raw not in ('nan', 'None', ''):
                    try:
                        nk = str(int(float(nf_raw.replace(',', '.'))))
                    except Exception:
                        nk = nf_raw
                    for col_prev in ['PREV PGTO', 'PREV PAGTO']:
                        prev_v = str(r.get(col_prev, '')).strip()
                        if prev_v and prev_v not in ('nan', 'None', '', 'S/PREV'):
                            try:
                                old_prev[nk] = pd.to_datetime(
                                    prev_v, dayfirst=True, errors='coerce')
                            except Exception:
                                pass
                            break

        print(f"   ✓ {len(rbase_conhecido):,} RBASEs conhecidos")
        print(f"   ✓ {len(rbase_to_grupo):,} com GRUPO · "
              f"{len(rbase_to_atraso):,} com ATRASO · "
              f"{len(rbase_to_paga):,} com PAGA NA DATA")
    except Exception as e:
        print(f"   ⚠️  Aviso ao ler planilha anterior: {e}")

    # ── 3. CONSTRUIR DATAFRAME ───────────────────────────────────────
    # Mapeamento Nutricash (24 colunas, sem IR/ISS):
    # A=RBASE, B=CLIENTE, C=CLIENTE2, D=CNPJ, E=UF, F=TIPO, G=PRODUTO
    # H=EMISSAO, I=VENCIMENTO, J=NUM DOC, K=VL SALDO, L=VL TITULO
    # M=EXECUTIVO, O=ATRASO, P=COND PAGTO, Q=NF, R=EMPRESA
    # T=PORTADOR, U=LINK NFSE, W=CIDADE
    print("\n🔄 Reorganizando colunas ...")
    out = pd.DataFrame({
        'EMPRESA'            : orig('R'),   # idx 17
        'EXECUTIVO'          : orig('M'),   # idx 12
        'PRODUTO'            : orig('G'),   # idx  6
        'CNPJ'               : orig('D'),   # idx  3
        'RBASE'              : orig('A'),   # idx  0
        'NF'                 : orig('Q'),   # idx 16
        'NUM DOC'            : orig('J'),   # idx  9
        'CLIENTE'            : orig('B'),   # idx  1
        'CLIENTE 2'          : orig('C'),   # idx  2
        'UF'                 : orig('E'),   # idx  4
        'CIDADE'             : orig('W'),   # idx 22
        'TIPO'               : orig('F'),   # idx  5
        'GRUPO'              : pd.Series([''] * N, dtype=str),
        'EMISSAO'            : pd.to_datetime(
                                   orig('H').str.strip(),
                                   errors='coerce', dayfirst=False),
        'VENCIMENTO'         : pd.to_datetime(
                                   orig('I').str.strip(),
                                   errors='coerce', dayfirst=False),
        'VL TITULO'          : pd.to_numeric(
                                   orig('L').str.replace(',', '.', regex=False),
                                   errors='coerce'),
        'VL SALDO'           : pd.to_numeric(
                                   orig('K').str.replace(',', '.', regex=False),
                                   errors='coerce'),
        'PAGA NA DATA'       : pd.Series([''] * N, dtype=str),
        'ATRASO'             : pd.Series([np.nan] * N, dtype=float),
        'PREV PAGTO'         : pd.NaT,
        'BANCO'              : pd.Series([''] * N, dtype=str),
        'LINK NFSE'          : orig('U'),   # idx 20
        'COND PAGTO'         : orig('P'),   # idx 15
        'CODIGO DE PAGAMENTO': orig('T'),   # idx 19
    })
    out['PREV PAGTO'] = pd.to_datetime(out['PREV PAGTO'], errors='coerce')

    # ── 4. ESPELHAR CAMPOS POR RBASE ─────────────────────────────────
    # REVISAR somente se o RBASE não existe na planilha anterior
    print("⚙️  Espelhando campos por RBASE ...")

    def rbase_lookup(rb, mapping, default_known='', default_unknown='REVISAR'):
        key = safe_rbase(rb)
        if key in rbase_conhecido:
            return mapping.get(key, default_known)
        return default_unknown

    def lookup_nf_key(nf_raw):
        try:
            return str(int(float(str(nf_raw).replace(',', '.').strip())))
        except Exception:
            return str(nf_raw).strip()

    def prev_lookup(row):
        if safe_rbase(row['RBASE']) not in rbase_conhecido:
            return 'REVISAR'
        return old_prev.get(lookup_nf_key(row['NF']), pd.NaT)

    out['GRUPO']        = out['RBASE'].apply(lambda rb: rbase_lookup(rb, rbase_to_grupo))
    out['PAGA NA DATA'] = out['RBASE'].apply(lambda rb: rbase_lookup(rb, rbase_to_paga))
    out['ATRASO']       = out['RBASE'].apply(
        lambda rb: rbase_lookup(rb, rbase_to_atraso, default_known=np.nan))
    out['PREV PAGTO']   = out[['RBASE', 'NF']].apply(prev_lookup, axis=1)

    print(f"   ✓ GRUPO       → {(out['GRUPO']       != 'REVISAR').sum():,} ok · {(out['GRUPO']       == 'REVISAR').sum():,} REVISAR")
    print(f"   ✓ PAGA NA DATA→ {(out['PAGA NA DATA'] != 'REVISAR').sum():,} ok · {(out['PAGA NA DATA'] == 'REVISAR').sum():,} REVISAR")
    print(f"   ✓ ATRASO      → {out['ATRASO'].apply(lambda v: v != 'REVISAR' and pd.notna(v)).sum():,} ok · {(out['ATRASO'] == 'REVISAR').sum():,} REVISAR")

    # ── 5. TEXTO PARA MAIÚSCULAS ─────────────────────────────────────
    STR_COLS = [
        'EMPRESA', 'EXECUTIVO', 'PRODUTO', 'CNPJ', 'CLIENTE', 'CLIENTE 2',
        'UF', 'CIDADE', 'TIPO', 'GRUPO', 'BANCO', 'COND PAGTO',
        'CODIGO DE PAGAMENTO', 'PAGA NA DATA', 'NUM DOC', 'LINK NFSE'
    ]
    for c in STR_COLS:
        if c in out.columns:
            out[c] = out[c].apply(to_upper)

    # ── 6. SEPARAR A VENCER / VENCIDOS ──────────────────────────────
    mask_venc = out['VENCIMENTO'].notna() & (out['VENCIMENTO'] < TODAY)
    df_av = out[~mask_venc].copy().reset_index(drop=True)
    df_vd = out[ mask_venc].copy().reset_index(drop=True)
    df_vd = df_vd.drop(columns=['PAGA NA DATA', 'UF'], errors='ignore')

    print(f"\n📋 A VENCER : {len(df_av):,} registros")
    print(f"📋 VENCIDOS : {len(df_vd):,} registros")

    # ── 7. GERAR EXCEL ───────────────────────────────────────────────
    print("\n📝 Gerando Excel formatado ...")
    output_buf = io.BytesIO()
    wb = xlsxwriter.Workbook(output_buf, {'in_memory': True,
                                           'default_date_format': 'dd/mm/yyyy'})

    F_HDR  = wb.add_format({'bold':True,'bg_color':'#1C2B5A','font_color':'#FFFFFF',
                             'border':1,'align':'center','valign':'vcenter',
                             'text_wrap':True,'font_name':'Arial','font_size':10})
    F_DATE = wb.add_format({'num_format':'dd/mm/yyyy','align':'center',
                             'font_name':'Arial','font_size':10})
    F_MON  = wb.add_format({'num_format':'#,##0.00','align':'right',
                             'font_name':'Arial','font_size':10})
    F_NF   = wb.add_format({'num_format':'000000','align':'center',
                             'font_name':'Arial','font_size':10})
    F_INT  = wb.add_format({'num_format':'0','align':'center',
                             'font_name':'Arial','font_size':10})
    F_CTR  = wb.add_format({'align':'center','font_name':'Arial','font_size':10})
    F_TXT  = wb.add_format({'align':'left','font_name':'Arial','font_size':10})
    F_LNK  = wb.add_format({'align':'left','font_name':'Arial','font_size':9,
                             'font_color':'#0563C1','underline':1})
    F_REV  = wb.add_format({'bold':True,'font_color':'#FF0000','align':'center',
                             'font_name':'Arial','font_size':10})

    DATE_COLS = {'EMISSAO', 'VENCIMENTO', 'PREV PAGTO'}
    MON_COLS  = {'VL TITULO', 'VL SALDO'}          # sem IR/ISS
    NF_COLS   = {'NF'}
    INT_COLS  = {'ATRASO', 'RBASE'}
    CTR_COLS  = {'UF', 'TIPO', 'PRODUTO', 'GRUPO', 'PAGA NA DATA',
                 'CODIGO DE PAGAMENTO', 'BANCO'}
    COL_WIDTHS = {
        'EMPRESA':18,'EXECUTIVO':28,'PRODUTO':10,'CNPJ':22,'RBASE':10,
        'NF':10,'NUM DOC':18,'CLIENTE':44,'CLIENTE 2':44,'UF':6,
        'CIDADE':24,'TIPO':12,'GRUPO':24,'EMISSAO':14,'VENCIMENTO':14,
        'VL TITULO':16,'VL SALDO':16,'PAGA NA DATA':14,'ATRASO':10,
        'PREV PAGTO':14,'BANCO':14,'LINK NFSE':62,'COND PAGTO':16,
        'CODIGO DE PAGAMENTO':24,
    }

    def write_tab(df_tab, sname):
        ws   = wb.add_worksheet(sname)
        cols = list(df_tab.columns)
        nr, nc = len(df_tab), len(cols)
        ws.set_row(0, 32)
        for ci, cn in enumerate(cols):
            ws.write(0, ci, cn, F_HDR)

        for ri in range(nr):
            xrow = ri + 1
            for ci, cn in enumerate(cols):
                val  = df_tab.iloc[ri][cn]
                miss = (val is None or
                        (isinstance(val, float) and pd.isna(val)) or
                        str(val).strip() in ('nan', 'None', ''))

                if str(val).strip() == 'REVISAR':
                    ws.write(xrow, ci, 'REVISAR', F_REV)
                elif cn in DATE_COLS:
                    isv = (not miss and
                           isinstance(val, (pd.Timestamp, datetime)) and
                           not (isinstance(val, pd.Timestamp) and pd.isna(val)))
                    if isv:
                        try:
                            dt = val.to_pydatetime() if isinstance(val, pd.Timestamp) else val
                            ws.write_datetime(xrow, ci, dt, F_DATE)
                        except Exception:
                            ws.write(xrow, ci, '', F_DATE)
                    else:
                        ws.write(xrow, ci, '', F_DATE)
                elif cn in MON_COLS:
                    if not miss:
                        try: ws.write_number(xrow, ci, float(val), F_MON)
                        except: ws.write(xrow, ci, '', F_TXT)
                    else:
                        ws.write(xrow, ci, '', F_MON)
                elif cn in NF_COLS:
                    raw = str(val).strip()
                    if raw and raw not in ('nan', 'None', ''):
                        try: ws.write_number(xrow, ci, int(float(raw.replace(',', '.'))), F_NF)
                        except: ws.write(xrow, ci, raw, F_TXT)
                    else:
                        ws.write(xrow, ci, '', F_TXT)
                elif cn in INT_COLS:
                    if not miss:
                        try: ws.write_number(xrow, ci, float(val), F_INT)
                        except: ws.write(xrow, ci, '', F_TXT)
                    else:
                        ws.write(xrow, ci, '', F_INT)
                elif cn == 'LINK NFSE':
                    s = '' if miss else str(val).strip()
                    if s.startswith('http'):
                        try: ws.write_url(xrow, ci, s, F_LNK, s[:255])
                        except: ws.write(xrow, ci, s, F_LNK)
                    else:
                        ws.write(xrow, ci, s, F_TXT)
                elif cn in CTR_COLS:
                    ws.write(xrow, ci, '' if miss else str(val).strip(), F_CTR)
                else:
                    ws.write(xrow, ci, '' if miss else str(val).strip(), F_TXT)

        for ci, cn in enumerate(cols):
            w = COL_WIDTHS.get(cn)
            if w is None:
                try: ml = int(df_tab[cn].astype(str).str.len().max() or 10)
                except: ml = 10
                w = min(max(len(cn), ml) + 3, 55)
            ws.set_column(ci, ci, w)
        ws.autofilter(0, 0, nr, nc - 1)
        ws.freeze_panes(1, 0)

    # ── 9. DASHBOARD EXECUTIVO COMPLETO ────────────────────────────
    print("\n📊 Gerando dashboard executivo (10 seções) ...")
    from dashboard import write_dashboard as _dash
    _dash(wb, df_av, df_vd, TODAY, cor_hdr='#1C2B5A')

    write_tab(df_av, 'A VENCER')
    write_tab(df_vd, 'VENCIDOS')
    wb.close()

    output_buf.seek(0)
    return output_buf.getvalue()
