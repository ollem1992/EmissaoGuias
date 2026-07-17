import streamlit as st
import pandas as pd
import pdfplumber
import io
import re
from pypdf import PdfReader, PdfWriter

# ============================================================
# 🎨 CONFIGURAÇÃO DA PÁGINA E TEMA DARK
# ============================================================
st.set_page_config(page_title="Conferência de GNREs", page_icon="📋", layout="centered")

st.markdown("""
    <style>
    .stApp { background-color: #0d0e12 !important; color: #ffffff !important; }
    h1, h2, h3, h4, p, span, label, div { color: #ffffff !important; }
    .stButton>button { background-color: #1e2029 !important; color: #00ffcc !important; border: 1px solid #333333 !important; }
    .stButton>button:hover { background-color: #2a2d3a !important; border: 1px solid #00ffcc !important; }
    section[data-testid="stFileUploadDropzone"] { background-color: #1e2029 !important; border: 2px dashed #00ffcc !important; border-radius: 8px; padding: 20px; }
    section[data-testid="stFileUploadDropzone"]:hover { background-color: #2a2d3a !important; border: 2px dashed #00ffcc !important; }
    .stTextInput>div>div>input { background-color: #1e2029 !important; color: #ffffff !important; border: 1px solid #444444 !important; }
    .stSpinner > div { border-top-color: #00ffcc !important; }
    </style>
""", unsafe_allow_html=True)

# ============================================================
# 🔐 AUTENTICAÇÃO SIMPLES
# ============================================================
if "autenticado" not in st.session_state:
    st.session_state.autenticado = False

if not st.session_state.autenticado:
    st.title("🔒 Acesso Restrito — Fiscal")
    st.write("Identifique-se para liberar o sistema de conferência.")
    
    usuario = st.text_input("Usuário")
    senha = st.text_input("Senha", type="password")
    
    if st.button("ENTRAR NO SISTEMA", use_container_width=True):
        if usuario == "fiscal" and senha == "fiscal":
            st.session_state.autenticado = True
            st.rerun()
        else:
            st.error("❌ Usuário ou senha incorretos.")
    st.stop()

# ============================================================
# ⚙️ CONFIGURAÇÕES FISCAIS FIXAS
# ============================================================
COLUNA_NOTA   = 'Nº NOTA'
COLUNA_UF     = 'UF'
COLUNA_VALOR1 = 'VALOR 1'
COLUNA_VALOR2 = 'VALOR 2'
COLUNA_JUROS  = 'JUROS'
SHEET_NAME    = 'Resumo'

BANCO_BRASIL  = {'AC'}
BRADESCO      = {'MS', 'PI'}
ITAU_ARQUIVO  = {'AL','AM','AP','BA','CE','DF','ES','GO','MA','MG','MT','PA','PB','PE','PR','RJ','RN','RO','RR','RS','SC','SE','SP','TO'}
ENTREGA_FISICA = {'AC', 'ES', 'MS', 'PI', 'SP'}

def classificar_banco(uf):
    uf = str(uf).strip().upper()
    if uf in BANCO_BRASIL: return 'Banco do Brasil'
    if uf in BRADESCO: return 'Bradesco'
    if uf in ITAU_ARQUIVO: return 'Itaú (arquivo)'
    return 'Não mapeado'

def limpar_valor_pdf(texto_valor):
    v = texto_valor.strip().replace('R$', '').replace(' ', '').replace('\xa0', '')
    if ',' in v and '.' in v: v = v.replace('.', '').replace(',', '.')
    elif ',' in v: v = v.replace(',', '.')
    return float(v)

def limpar_valor_excel(v):
    if pd.isna(v): return 0.0
    if isinstance(v, (int, float)): return float(v)
    s = str(v).strip().replace('R$','').replace('R','').replace('\xa0','').replace(' ','')
    if not s or s == '': return 0.0
    if ',' in s and '.' in s: s = s.replace('.','').replace(',','.')
    elif ',' in s: s = s.replace(',','.')
    try: return float(s)
    except: return 0.0

# ============================================================
# 🖥️ INTERFACE PRINCIPAL
# ============================================================
st.title("📋 Conferência de GNREs (Fiscal)")

if st.button("🚪 Encerrar Sessão (Sair)"):
    st.session_state.autenticado = False
    st.rerun()

st.write("---")

pdf_upload = st.file_uploader("1. Selecione o PDF Bruto de Guias (.pdf)", type=["pdf"])
planilha_upload = st.file_uploader("2. Selecione a Planilha Excel (.xlsx)", type=["xlsx"])

st.write("---")

# ESTRATÉGIA 3: O BOTÃO DE MODO ANALISTA / LOTE ATRASADO
lote_atrasado = st.checkbox("⚠️ Lote Atrasado / Feriado (Ativa o 'Modo Analista' para descobrir juros automaticamente)", value=False)

if "processo_concluido" not in st.session_state:
    st.session_state.processo_concluido = False

if st.button("🚀 INICIAR CONFERÊNCIA E SEPARAÇÃO", use_container_width=True):
    if not planilha_upload or not pdf_upload:
        st.error("❌ Erro: Carregue a planilha e o PDF antes de rodar.")
    else:
        with st.spinner("⚡ Lendo PDF e ativando robô de conferência..."):
            
            pdf_bytes = pdf_upload.read()
            excel_bytes = planilha_upload.read()
            
            # --- EXTRAÇÃO DO PDF ---
            resultados_pdf = []
            with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
                for i, page in enumerate(pdf.pages):
                    texto = page.extract_text()
                    if not texto: continue

                    uf, doc, valor = '??', '', None

                    # Padrão Tradicional
                    m_val_gnre = re.search(r'Total\s+a\s+Recolher.*?R\$\s*([\d.,]+)', texto, re.IGNORECASE | re.DOTALL)
                    if m_val_gnre:
                        try: valor = limpar_valor_pdf(m_val_gnre.group(1))
                        except ValueError: pass
                        m_uf = re.search(r'Guia Nacional de Recolhimento.*?\n([A-Z]{2})\s+\d{5,}', texto, re.DOTALL)
                        if m_uf: uf = m_uf.group(1)
                        m_doc = re.search(r'N[ºo°]\s+Documento\s+de\s+Origem\n[^\n]+?(\d{5,})\s*$', texto, re.MULTILINE)
                        if m_doc: doc = m_doc.group(1).lstrip('0')
                    
                    # Padrão SP (DARE-SP)
                    elif re.search(r'DARE-SP', texto, re.IGNORECASE) or re.search(r'S[ãa]o\s*Paulo', texto, re.IGNORECASE):
                        uf = 'SP'
                        
                        # Para o DARE, a extração de texto pode misturar as colunas.
                        # A forma segura é caçar todos os "R$ X,XX" da página e pegar o maior (que já embute os juros/multa).
                        valores_encontrados = re.findall(r'R\$\s*([\d.,]+)', texto)
                        valores_float = []
                        for v in valores_encontrados:
                            try: valores_float.append(limpar_valor_pdf(v))
                            except: pass
                        if valores_float: 
                            valor = max(valores_float)
                            
                        m_doc_dare = re.search(r'NFe?\s*[nN]?[ºo°]?\s*[:]?\s*(\d+)', texto, re.IGNORECASE)
                        if m_doc_dare: doc = m_doc_dare.group(1)
                    
                    # Padrão ES (DUA)
                    elif re.search(r'Esp[íi]rito\s*Santo', texto, re.IGNORECASE) or re.search(r'Documento\s*[UÚuú]nico', texto, re.IGNORECASE):
                        uf = 'ES'
                        m_val_dua = re.search(r'(?:Total|Receita)[\s\n]*R?\$?[\s\n]*([\d.,]{3,})', texto, re.IGNORECASE | re.DOTALL)
                        if m_val_dua:
                            try: valor = limpar_valor_pdf(m_val_dua.group(1))
                            except ValueError: pass
                        if valor is None:
                            valores_encontrados = re.findall(r'R\$\s*([\d.,]{3,})', texto)
                            valores_float = []
                            for v in valores_encontrados:
                                try: valores_float.append(limpar_valor_pdf(v))
                                except: pass
                            if valores_float: valor = max(valores_float)
                        m_doc_dua = re.search(r'(?:NFe|documento)[^\d]*(\d{5,})', texto, re.IGNORECASE)
                        if m_doc_dua: doc = m_doc_dua.group(1)

                    if valor is not None:
                        resultados_pdf.append({'Página': i + 1, 'UF': uf, 'Nº Nota': doc, 'Total a Recolher (R$)': valor})

            if not resultados_pdf:
                st.error("❌ Nenhum valor encontrado no PDF.")
                st.stop()

            # --- LEITURA DO EXCEL ---
            try:
                df_excel = pd.read_excel(io.BytesIO(excel_bytes), sheet_name=SHEET_NAME)
            except Exception as e:
                st.error(f"❌ Erro ao ler a aba '{SHEET_NAME}'. Verifique a planilha.")
                st.stop()

            obrigatorias = [COLUNA_NOTA, COLUNA_UF, COLUNA_VALOR1]
            faltando = [c for c in obrigatorias if c not in df_excel.columns]
            if faltando:
                st.error(f"❌ Colunas não encontradas na planilha: {faltando}")
                st.stop()

            tem_valor2 = COLUNA_VALOR2 in df_excel.columns
            tem_juros  = COLUNA_JUROS in df_excel.columns

            df_excel['_v1'] = df_excel[COLUNA_VALOR1].apply(limpar_valor_excel)
            df_excel['_v2'] = df_excel[COLUNA_VALOR2].apply(limpar_valor_excel) if tem_valor2 else 0.0
            df_excel['_jr'] = df_excel[COLUNA_JUROS].apply(limpar_valor_excel) if tem_juros else 0.0
            df_excel['_valor_total'] = df_excel['_v1'] + df_excel['_v2'] + df_excel['_jr']

            df_excel['_nota_valida'] = df_excel[COLUNA_NOTA].notna() & (df_excel[COLUNA_NOTA] != 0)
            df_excel['_valor_valido'] = df_excel['_valor_total'] > 0
            df = df_excel[df_excel['_nota_valida'] & df_excel['_valor_valido']].copy()

            df['_uf']     = df[COLUNA_UF].astype(str).str.strip().str.upper()
            df['_nota_str'] = df[COLUNA_NOTA].astype(str).str.strip().str.replace(r'\.0$', '', regex=True).str.lstrip('0')
            
            # --- O "BEM BOLADO" (ESTRATÉGIAS 1, 2 E 3 INTEGRADAS) ---
            guias_disponiveis = []
            for i in resultados_pdf:
                nota_pdf = str(i['Nº Nota']).strip().lstrip('0') if i['Nº Nota'] else ''
                guias_disponiveis.append({
                    'uf': i['UF'], 'nota': nota_pdf, 'valor': i['Total a Recolher (R$)'], 
                    'pagina': i['Página'] - 1, 'usada': False
                })

            paginas_itau_arquivo = []
            paginas_impressao = []
            relatorio_juros = []

            def buscar_guia_inteligente(nota_alvo, uf_alvo, valor_alvo, permite_atraso):
                for g in guias_disponiveis:
                    if not g['usada'] and g['nota'] == nota_alvo and abs(g['valor'] - valor_alvo) <= 0.02:
                        g['usada'] = True; return g['pagina'], 0.0
                
                for g in guias_disponiveis:
                    if not g['usada'] and g['nota'] == nota_alvo and g['valor'] > valor_alvo + 0.02:
                        juros = g['valor'] - valor_alvo
                        g['usada'] = True; return g['pagina'], juros
                
                for g in guias_disponiveis:
                    if not g['usada'] and g['uf'] == uf_alvo and abs(g['valor'] - valor_alvo) <= 0.02:
                        g['usada'] = True; return g['pagina'], 0.0
                
                if permite_atraso:
                    for g in guias_disponiveis:
                        if not g['usada'] and g['uf'] == uf_alvo and (valor_alvo - 0.02) <= g['valor'] <= (valor_alvo * 1.20):
                            juros = g['valor'] - valor_alvo
                            g['usada'] = True; return g['pagina'], juros
                
                return None, 0.0

            total_excel_base = 0.0
            total_pdf_pago = 0.0

            for idx, row in df.iterrows():
                uf, v_total, v1, v2, jr = row['_uf'], row['_valor_total'], row['_v1'], row['_v2'], row['_jr']
                nota_excel = row['_nota_str']
                total_excel_base += v_total
                
                pag_total, juros_total = buscar_guia_inteligente(nota_excel, uf, v_total, lote_atrasado)
                if pag_total is not None:
                    if uf in ENTREGA_FISICA: paginas_impressao.append(pag_total)
                    elif uf in ITAU_ARQUIVO: paginas_itau_arquivo.append(pag_total)
                    
                    total_pdf_pago += (v_total + juros_total)
                    if juros_total > 0.02:
                        relatorio_juros.append({"Nota": nota_excel, "UF": uf, "Valor Base": v_total, "Total Pago": v_total + juros_total, "Juros/Multa SEFAZ": juros_total})
                    continue
                
                if v2 > 0:
                    pag_v1, juros1 = buscar_guia_inteligente(nota_excel, uf, v1 + jr, lote_atrasado)
                    pag_v2, juros2 = buscar_guia_inteligente(nota_excel, uf, v2, lote_atrasado)
                    
                    if pag_v1 is not None:
                        if uf in ENTREGA_FISICA: paginas_impressao.append(pag_v1)
                        elif uf in ITAU_ARQUIVO: paginas_itau_arquivo.append(pag_v1)
                        total_pdf_pago += (v1 + jr + juros1)
                        if juros1 > 0.02: relatorio_juros.append({"Nota": f"{nota_excel} (Guia 1)", "UF": uf, "Valor Base": v1+jr, "Total Pago": v1+jr+juros1, "Juros/Multa SEFAZ": juros1})
                            
                    if pag_v2 is not None:
                        if uf in ENTREGA_FISICA: paginas_impressao.append(pag_v2)
                        elif uf in ITAU_ARQUIVO: paginas_itau_arquivo.append(pag_v2)
                        total_pdf_pago += (v2 + juros2)
                        if juros2 > 0.02: relatorio_juros.append({"Nota": f"{nota_excel} (Guia 2)", "UF": uf, "Valor Base": v2, "Total Pago": v2+juros2, "Juros/Multa SEFAZ": juros2})

            # --- GERAÇÃO DOS ARQUIVOS NA MEMÓRIA ---
            reader = PdfReader(io.BytesIO(pdf_bytes))
            
            pdf_itau_bytes = io.BytesIO()
            if paginas_itau_arquivo:
                writer_itau = PdfWriter()
                for pag in paginas_itau_arquivo: writer_itau.add_page(reader.pages[pag])
                writer_itau.write(pdf_itau_bytes)
                
            pdf_imp_bytes = io.BytesIO()
            if paginas_impressao:
                writer_imp = PdfWriter()
                for pag in paginas_impressao: writer_imp.add_page(reader.pages[pag])
                writer_imp.write(pdf_imp_bytes)

            st.session_state.pdf_itau = pdf_itau_bytes.getvalue() if paginas_itau_arquivo else None
            st.session_state.pdf_imp = pdf_imp_bytes.getvalue() if paginas_impressao else None
            st.session_state.qtd_itau = len(paginas_itau_arquivo)
            st.session_state.qtd_imp = len(paginas_impressao)
            
            st.session_state.total_excel_base = total_excel_base
            st.session_state.total_pdf_pago = total_pdf_pago
            st.session_state.relatorio_juros = relatorio_juros
            
            st.session_state.processo_concluido = True


# ============================================================
# 3. ÁREA DE EXIBIÇÃO / RELATÓRIOS
# ============================================================
if st.session_state.processo_concluido:
    st.write("---")
    st.subheader("📊 Resumo Financeiro da Conferência")
    col1, col2, col3 = st.columns(3)
    
    col1.metric("Total Base (Planilha)", f"R$ {st.session_state.total_excel_base:,.2f}")
    col2.metric("Total Final Pago (PDF)", f"R$ {st.session_state.total_pdf_pago:,.2f}")
    
    diferenca = st.session_state.total_pdf_pago - st.session_state.total_excel_base
    if abs(diferenca) < 0.02:
        col3.metric("Juros/Acréscimos", "R$ 0,00", delta_color="off")
        st.success("✅ Perfeito! Todos os valores do PDF bateram exatamente com a planilha base.")
    else:
        col3.metric("Juros/Acréscimos", f"+ R$ {diferenca:,.2f}", delta_color="inverse")
    
    if len(st.session_state.relatorio_juros) > 0:
        st.warning("⚠️ **ATENÇÃO:** O sistema identificou guias com acréscimos/juros cobrados pela SEFAZ:")
        df_juros_report = pd.DataFrame(st.session_state.relatorio_juros)
        
        st.dataframe(df_juros_report.style.format({
            "Valor Base": "R$ {:.2f}",
            "Total Pago": "R$ {:.2f}",
            "Juros/Multa SEFAZ": "R$ {:.2f}"
        }), use_container_width=True)

    st.write("---")
    st.subheader("🗂️ Download dos Lotes Separados")
    
    col_btn1, col_btn2 = st.columns(2)
    
    if st.session_state.qtd_itau > 0:
        col_btn1.download_button(
            label=f"📥 BAIXAR GUIAS ITAÚ ({st.session_state.qtd_itau} pág)", 
            data=st.session_state.pdf_itau, 
            file_name="guias_itau_arquivo.pdf", 
            mime="application/pdf", 
            use_container_width=True
        )
    else:
        col_btn1.info("Nenhuma guia Itaú Arquivo neste lote.")

    if st.session_state.qtd_imp > 0:
        col_btn2.download_button(
            label=f"📥 BAIXAR GUIAS FÍSICAS ({st.session_state.qtd_imp} pág)", 
            data=st.session_state.pdf_imp, 
            file_name="guias_impressao_fisica.pdf", 
            mime="application/pdf", 
            use_container_width=True
        )
    else:
        col_btn2.info("Nenhuma guia de Impressão Física neste lote.")
