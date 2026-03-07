import streamlit as st
import io
from datetime import datetime

# ── Configuração da página ───────────────────────────────────────────
st.set_page_config(
    page_title="Processador MAXIFROTA CR",
    page_icon="🚗",
    layout="centered",
)

# ── CSS personalizado ────────────────────────────────────────────────
st.markdown("""
<style>
    /* Fundo e fonte geral */
    html, body, [class*="css"] { font-family: Arial, sans-serif; }

    /* Cabeçalho principal */
    .header-box {
        background: linear-gradient(135deg, #00205B 0%, #1a3a7a 100%);
        border-radius: 12px;
        padding: 28px 32px;
        margin-bottom: 28px;
        text-align: center;
    }
    .header-box h1 { color: white; font-size: 1.9em; margin: 0 0 6px 0; }
    .header-box p  { color: #a8c4e0; font-size: 0.95em; margin: 0; }

    /* Cards de upload */
    .upload-card {
        background: #f8fafd;
        border: 2px dashed #2E75B6;
        border-radius: 10px;
        padding: 20px;
        margin-bottom: 16px;
        text-align: center;
    }
    .upload-card h4 { color: #00205B; margin: 0 0 8px 0; font-size: 1em; }
    .upload-card p  { color: #666; font-size: 0.85em; margin: 0; }

    /* Botão de processar */
    div.stButton > button {
        background: linear-gradient(135deg, #00205B, #2E75B6);
        color: white;
        font-size: 1.1em;
        font-weight: bold;
        border: none;
        border-radius: 8px;
        padding: 14px 40px;
        width: 100%;
        cursor: pointer;
        transition: opacity 0.2s;
    }
    div.stButton > button:hover { opacity: 0.88; }

    /* Caixa de sucesso */
    .success-box {
        background: #e8f5e9;
        border-left: 5px solid #2e7d32;
        border-radius: 8px;
        padding: 18px 22px;
        margin-top: 20px;
    }
    .success-box h3 { color: #1b5e20; margin: 0 0 6px 0; }
    .success-box p  { color: #2e7d32; margin: 0; font-size: 0.92em; }

    /* Caixa de erro */
    .error-box {
        background: #ffeaea;
        border-left: 5px solid #c62828;
        border-radius: 8px;
        padding: 18px 22px;
        margin-top: 20px;
    }

    /* Instruções */
    .step {
        display: flex;
        align-items: flex-start;
        gap: 14px;
        margin-bottom: 12px;
        background: #f0f4fa;
        border-radius: 8px;
        padding: 12px 16px;
    }
    .step-num {
        background: #00205B;
        color: white;
        border-radius: 50%;
        width: 28px; height: 28px;
        display: flex; align-items: center; justify-content: center;
        font-weight: bold; font-size: 0.9em; flex-shrink: 0;
    }
    .step-txt { color: #333; font-size: 0.92em; line-height: 1.5; }

    /* Rodapé */
    .footer { text-align: center; color: #aaa; font-size: 0.8em; margin-top: 40px; }
</style>
""", unsafe_allow_html=True)

# ── Cabeçalho ────────────────────────────────────────────────────────
st.markdown("""
<div class="header-box">
    <h1>🚗 PROCESSADOR MAXIFROTA CR</h1>
    <p>Faça upload dos arquivos, clique em processar e baixe o Excel pronto.</p>
</div>
""", unsafe_allow_html=True)

# ── Instruções ───────────────────────────────────────────────────────
with st.expander("📖 Como usar — clique para ver", expanded=False):
    st.markdown("""
    <div class="step">
        <div class="step-num">1</div>
        <div class="step-txt">
            Selecione o arquivo <strong>CRMX.CSV</strong> no campo abaixo.
            É o arquivo de contas a receber exportado do sistema.
        </div>
    </div>
    <div class="step">
        <div class="step-num">2</div>
        <div class="step-txt">
            Selecione o arquivo <strong>CR_MAXIFROTA_2026.XLSX</strong>
            — a planilha do mês anterior com grupos e histórico dos clientes.
        </div>
    </div>
    <div class="step">
        <div class="step-num">3</div>
        <div class="step-txt">
            Clique em <strong>⚡ PROCESSAR PLANILHA</strong> e aguarde alguns
            segundos. O arquivo aparecerá automaticamente para download.
        </div>
    </div>
    <div class="step">
        <div class="step-num">4</div>
        <div class="step-txt">
            O Excel gerado já vem com 4 abas:<br>
            <strong>RESUMO A VENCER · RESUMO VENCIDOS · A VENCER · VENCIDOS</strong><br>
            Células em <span style="color:red;font-weight:bold">REVISAR</span>
            indicam clientes novos sem histórico.
        </div>
    </div>
    """, unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ── Uploads ──────────────────────────────────────────────────────────
col1, col2 = st.columns(2)

with col1:
    st.markdown("""
    <div class="upload-card">
        <h4>📄 Arquivo 1</h4>
        <p>CRMX.CSV — dados brutos do sistema</p>
    </div>
    """, unsafe_allow_html=True)
    csv_file = st.file_uploader(
        "Selecione o CRMX.CSV",
        type=["csv"],
        key="csv",
        label_visibility="collapsed",
    )
    if csv_file:
        st.success(f"✅ {csv_file.name}  ({csv_file.size // 1024} KB)")

with col2:
    st.markdown("""
    <div class="upload-card">
        <h4>📊 Arquivo 2</h4>
        <p>CR_MAXIFROTA_2026.XLSX — planilha anterior</p>
    </div>
    """, unsafe_allow_html=True)
    xlsx_file = st.file_uploader(
        "Selecione o CR_MAXIFROTA_2026.XLSX",
        type=["xlsx", "xls"],
        key="xlsx",
        label_visibility="collapsed",
    )
    if xlsx_file:
        st.success(f"✅ {xlsx_file.name}  ({xlsx_file.size // 1024} KB)")

st.markdown("<br>", unsafe_allow_html=True)

# ── Botão de processar ───────────────────────────────────────────────
btn_disabled = not (csv_file and xlsx_file)

if btn_disabled:
    st.info("⬆️  Faça upload dos dois arquivos para habilitar o processamento.")

processar = st.button(
    "⚡  PROCESSAR PLANILHA",
    disabled=btn_disabled,
    use_container_width=True,
)

# ── Processamento ────────────────────────────────────────────────────
if processar and csv_file and xlsx_file:
    progress = st.progress(0, text="Iniciando...")

    try:
        # Importar a função de processamento
        from core import process_files
        import sys, io as _io

        # Capturar prints para exibir no log
        log_buffer = _io.StringIO()
        import contextlib

        progress.progress(10, text="📂 Lendo os arquivos...")
        csv_bytes  = csv_file.read()
        xlsx_bytes = xlsx_file.read()

        progress.progress(25, text="📊 Lendo CRMX.CSV e planilha anterior...")

        with contextlib.redirect_stdout(log_buffer):
            result_bytes = process_files(csv_bytes, xlsx_bytes)

        progress.progress(90, text="📝 Finalizando Excel...")

        ts       = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_name = f"CR_MAXIFROTA_PROCESSADO_{ts}.xlsx"

        progress.progress(100, text="✅ Concluído!")

        # ── Resultado ────────────────────────────────────────────────
        st.markdown("""
        <div class="success-box">
            <h3>🎉 Planilha gerada com sucesso!</h3>
            <p>Clique no botão abaixo para baixar o arquivo Excel.</p>
        </div>
        """, unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        st.download_button(
            label="📥  BAIXAR EXCEL PROCESSADO",
            data=result_bytes,
            file_name=out_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

        # Log detalhado (collapsível)
        log_text = log_buffer.getvalue()
        if log_text.strip():
            with st.expander("📋 Log do processamento", expanded=False):
                st.code(log_text, language=None)

    except Exception as e:
        import traceback
        progress.empty()
        st.markdown(f"""
        <div class="error-box">
            <strong>❌ Erro durante o processamento</strong><br><br>
            {str(e)}
        </div>
        """, unsafe_allow_html=True)
        with st.expander("🔍 Detalhes do erro"):
            st.code(traceback.format_exc())

# ── Rodapé ───────────────────────────────────────────────────────────
st.markdown("""
<div class="footer">
    MAXIFROTA CR · Processador de Planilhas · 2026
</div>
""", unsafe_allow_html=True)
