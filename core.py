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
        # Ajusta se cair em sexta, sábado, domingo ou feriado
        if res.weekday() >= 4 or res.date() in br_holidays(res.year):
            return next_util(res)
        return res
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

    # PREV PAGTO dos vencidos = VENCIMENTO + ATRASO dias → próximo dia útil
    # (sobrescreve o valor espelhado da planilha anterior)
    def calc_prev_vd(row):
        venc  = row.get('VENCIMENTO')
        atr   = row.get('ATRASO')
        if pd.isna(venc) or str(atr).strip() in ('', 'nan', 'None', 'REVISAR'):
            return pd.NaT
        return calc_prev(venc, atr)
    df_vd['PREV PAGTO'] = df_vd.apply(calc_prev_vd, axis=1)

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


    # ── 9. DASHBOARD EXECUTIVO COMPLETO ────────────────────────────
    print("\n📊 Gerando dashboard executivo (10 seções) ...")
    from dashboard import write_dashboard as _dash
    _dash(wb, df_av, df_vd, TODAY, cor_hdr='#00205B')

    write_tab(df_av, 'A VENCER')
    write_tab(df_vd, 'VENCIDOS')
    wb.close()

    output_buf.seek(0)
    return output_buf.getvalue()
