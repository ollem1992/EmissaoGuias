import streamlit as st
import pandas as pd
import pdfplumber
import io
import re
from pypdf import PdfReader, PdfWriter

# ============================================================
# 🎨 CONFIGURAÇÃO DA PÁGINA E TEMA DARK
# ============================================================
st.set_page_config(page_title="Conferência de GNREs", page_icon="📋", layout="wide")

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
            paginas_sem_leitura = []
            with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
                for i, page in enumerate(pdf.pages):
                    texto = page.extract_text()
                    if not texto:
                        paginas_sem_leitura.append(i + 1)
                        continue

                    uf, doc, valor = '??', '', None

                    # 1. Padrão Tradicional (GNRE)
                    m_val_gnre = re.search(r'Total\s+a\s+Recolher[^\d\n]*R?\$?\s*([\d.,]+)', texto, re.IGNORECASE)
                    if m_val_gnre:
                        try: valor = limpar_valor_pdf(m_val_gnre.group(1))
                        except ValueError: pass
                        
                        m_uf = re.search(r'Guia Nacional de Recolhimento.*?\n\s*([A-Z]{2})\b', texto, re.DOTALL)
                        if not m_uf:
                            m_uf = re.search(r'UF\s*Favorecida[^\n]*\n\s*([A-Z]{2})\b', texto, re.IGNORECASE)
                        if not m_uf:
                            m_uf = re.search(r'([A-Z]{2})\s+\d+', texto)
                        if m_uf: uf = m_uf.group(1)
                        
                        # Aceita nota com qualquer quantidade de dígitos (ex: \d+ em vez de fixo \d{5,})
                        m_doc = re.search(r'N[ºo°]\s*Documento\s*de\s*Origem[^\n]*\n[^\n\d]*?(\d+)', texto, re.IGNORECASE)
                        if not m_doc:
                            m_doc = re.search(r'Documento\s*de\s*Origem[^\d]*(\d+)', texto, re.IGNORECASE)
                        if m_doc: doc = m_doc.group(1).lstrip('0')
                    
                    # 2. Padrão SP (DARE-SP)
                    elif re.search(r'DARE-SP', texto, re.IGNORECASE) or re.search(r'S[ãa]o\s*Paulo', texto, re.IGNORECASE):
                        uf = 'SP'
                        # Busca prioritária por rótulo específico de valor para não capturar Base de Cálculo
                        m_val_dare = re.search(r'(?:09\s*-\s*Valor\s*Total|Total\s*a\s*Recolher|Valor\s*Total)[\s:]*R?\$?\s*([\d.,]+)', texto, re.IGNORECASE)
                        if m_val_dare:
                            try: valor = limpar_valor_pdf(m_val_dare.group(1))
                            except: valor = None
                        
                        if valor is None:
                            m_total_gen = re.search(r'\bTotal[\s:]*R?\$?\s*([\d.,]+)', texto, re.IGNORECASE)
                            if m_total_gen:
                                try: valor = limpar_valor_pdf(m_total_gen.group(1))
                                except: pass
                                
                        if valor is None:
                            valores_encontrados = re.findall(r'R\$\s*([\d.,]+)', texto)
                            valores_float = []
                            for v in valores_encontrados:
                                try:
                                    vf = limpar_valor_pdf(v)
                                    if vf > 0: valores_float.append(vf)
                                except: pass
                            if valores_float: 
                                valor = max(valores_float)
                            
                        m_doc_dare = re.search(r'NFe?\s*[nN]?[ºo°]?\s*[:]?\s*(\d+)', texto, re.IGNORECASE)
                        if not m_doc_dare:
                            m_doc_dare = re.search(r'(?:Documento\s*de\s*Origem|N[ºo°]\s*Doc\.?\s*Origem)[^\d]*(\d+)', texto, re.IGNORECASE)
                        if m_doc_dare: doc = m_doc_dare.group(1).lstrip('0')
                    
                    # 3. Padrão ES (DUA)
                    elif re.search(r'Esp[íi]rito\s*Santo', texto, re.IGNORECASE) or re.search(r'Documento\s*[UÚuú]nico', texto, re.IGNORECASE):
                        uf = 'ES'
                        m_val_dua = re.search(r'(?:Total\s*a\s*Recolher|Valor\s*Total|Total|Receita)[\s\n:]*R?\$?[\s\n]*([\d.,]{3,})', texto, re.IGNORECASE)
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
                        m_doc_dua = re.search(r'(?:NFe?|documento)[^\d]*(\d+)', texto, re.IGNORECASE)
                        if m_doc_dua: doc = m_doc_dua.group(1).lstrip('0')

                    # 4. Fallback genérico para outros modelos de guia
                    if valor is None:
                        m_gen_val = re.search(r'(?:Total\s*a\s*Pagar|Total\s*a\s*Recolher|Valor\s*Total)[\s:]*R?\$?\s*([\d.,]+)', texto, re.IGNORECASE)
                        if m_gen_val:
                            try: valor = limpar_valor_pdf(m_gen_val.group(1))
                            except: pass
                        m_gen_doc = re.search(r'N[ºo°]?\s*(?:da\s*)?Nota[^\d]*(\d+)', texto, re.IGNORECASE)
                        if m_gen_doc: doc = m_gen_doc.group(1).lstrip('0')

                    if valor is not None:
                        resultados_pdf.append({'Página': i + 1, 'UF': uf, 'Nº Nota': doc, 'Total a Recolher (R$)': valor})
                    else:
                        paginas_sem_leitura.append(i + 1)

            if not resultados_pdf:
                st.error("❌ Nenhum valor de guia foi identificado no PDF.")
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
            
            # --- O "BEM BOLADO" E CLASSIFICAÇÃO DE BANCOS ---
            guias_disponiveis = []
            for i in resultados_pdf:
                nota_pdf = str(i['Nº Nota']).strip().lstrip('0') if i['Nº Nota'] else ''
                guias_disponiveis.append({
                    'uf': i['UF'], 
                    'nota': nota_pdf, 
                    'valor': i['Total a Recolher (R$)'], 
                    'pagina': i['Página'] - 1, 
                    'pagina_num': i['Página'],
                    'usada': False
                })

            # Listas de páginas por banco / modalidade
            paginas_itau_arquivo = []
            paginas_itau_fisico  = []
            paginas_bradesco     = []
            paginas_bb           = []
            paginas_todas_fisicas = []
            paginas_outros       = []

            relatorio_juros = []
            notas_pendentes = []
            
            # Memória da tabela de resumo dos bancos
            resumo_bancos = {
                "Itau Arquivo": {"Qtd": 0, "Valor": 0.0},
                "Restantes Itaú": {"Qtd": 0, "Valor": 0.0},
                "Bradesco": {"Qtd": 0, "Valor": 0.0},
                "BB": {"Qtd": 0, "Valor": 0.0},
            }

            def classificar_resumo(uf_alvo):
                if uf_alvo in BANCO_BRASIL: return "BB"
                if uf_alvo in BRADESCO: return "Bradesco"
                if uf_alvo in ITAU_ARQUIVO and uf_alvo in ENTREGA_FISICA: return "Restantes Itaú"
                if uf_alvo in ITAU_ARQUIVO and uf_alvo not in ENTREGA_FISICA: return "Itau Arquivo"
                return "Outros"

            def destinar_pagina(uf_alvo, pag_idx, valor_pago):
                cat = classificar_resumo(uf_alvo)
                if cat == "BB":
                    paginas_bb.append(pag_idx)
                    paginas_todas_fisicas.append(pag_idx)
                    resumo_bancos["BB"]["Qtd"] += 1
                    resumo_bancos["BB"]["Valor"] += valor_pago
                elif cat == "Bradesco":
                    paginas_bradesco.append(pag_idx)
                    paginas_todas_fisicas.append(pag_idx)
                    resumo_bancos["Bradesco"]["Qtd"] += 1
                    resumo_bancos["Bradesco"]["Valor"] += valor_pago
                elif cat == "Restantes Itaú":
                    paginas_itau_fisico.append(pag_idx)
                    paginas_todas_fisicas.append(pag_idx)
                    resumo_bancos["Restantes Itaú"]["Qtd"] += 1
                    resumo_bancos["Restantes Itaú"]["Valor"] += valor_pago
                elif cat == "Itau Arquivo":
                    paginas_itau_arquivo.append(pag_idx)
                    resumo_bancos["Itau Arquivo"]["Qtd"] += 1
                    resumo_bancos["Itau Arquivo"]["Valor"] += valor_pago
                else:
                    paginas_outros.append(pag_idx)
                    paginas_todas_fisicas.append(pag_idx)

            def buscar_guia_inteligente(nota_alvo, uf_alvo, valor_alvo, permite_atraso):
                # 1. Nota exata + Valor exato (tolerância de R$ 0,02)
                for g in guias_disponiveis:
                    if not g['usada'] and g['nota'] and g['nota'] == nota_alvo and abs(g['valor'] - valor_alvo) <= 0.02:
                        g['usada'] = True
                        return g['pagina'], 0.0
                
                # 2. Nota exata + Valor com acréscimo/juros SEFAZ
                for g in guias_disponiveis:
                    if not g['usada'] and g['nota'] and g['nota'] == nota_alvo and g['valor'] > valor_alvo + 0.02:
                        juros = g['valor'] - valor_alvo
                        g['usada'] = True
                        return g['pagina'], juros
                
                # 3. Fallback: se a nota não foi lida do PDF mas a UF e o valor batem exatamente
                for g in guias_disponiveis:
                    if not g['usada'] and (not g['nota'] or g['nota'] == nota_alvo) and g['uf'] == uf_alvo and abs(g['valor'] - valor_alvo) <= 0.02:
                        g['usada'] = True
                        return g['pagina'], 0.0
                
                # 4. Modo analista (lote atrasado com acréscimo de até 20% na UF)
                if permite_atraso:
                    for g in guias_disponiveis:
                        if not g['usada'] and (not g['nota'] or g['nota'] == nota_alvo) and g['uf'] == uf_alvo and (valor_alvo - 0.02) <= g['valor'] <= (valor_alvo * 1.20):
                            juros = g['valor'] - valor_alvo
                            g['usada'] = True
                            return g['pagina'], juros
                
                return None, 0.0

            total_excel_base = 0.0
            total_pdf_pago = 0.0

            for idx, row in df.iterrows():
                uf, v_total, v1, v2, jr = row['_uf'], row['_valor_total'], row['_v1'], row['_v2'], row['_jr']
                nota_excel = row['_nota_str']
                total_excel_base += v_total
                
                # Tentativa 1: Guia com valor total consolidado
                pag_total, juros_total = buscar_guia_inteligente(nota_excel, uf, v_total, lote_atrasado)
                if pag_total is not None:
                    val_pago = v_total + juros_total
                    destinar_pagina(uf, pag_total, val_pago)
                    total_pdf_pago += val_pago
                    if juros_total > 0.02:
                        relatorio_juros.append({
                            "Nota": nota_excel, "UF": uf, "Valor Base": v_total, 
                            "Total Pago": val_pago, "Juros/Multa SEFAZ": juros_total
                        })
                    continue
                
                # Tentativa 2: Guias desmembradas (V1 + V2)
                if v2 > 0:
                    pag_v1, juros1 = buscar_guia_inteligente(nota_excel, uf, v1 + jr, lote_atrasado)
                    pag_v2, juros2 = buscar_guia_inteligente(nota_excel, uf, v2, lote_atrasado)
                    
                    if pag_v1 is not None:
                        val_pago1 = v1 + jr + juros1
                        destinar_pagina(uf, pag_v1, val_pago1)
                        total_pdf_pago += val_pago1
                        if juros1 > 0.02:
                            relatorio_juros.append({
                                "Nota": f"{nota_excel} (Guia 1)", "UF": uf, 
                                "Valor Base": v1 + jr, "Total Pago": val_pago1, "Juros/Multa SEFAZ": juros1
                            })
                    else:
                        notas_pendentes.append({
                            "Nota": f"{nota_excel} (Guia 1)", "UF": uf, 
                            "Valor Esperado": v1 + jr, "Motivo": "Guia 1 não localizada no PDF"
                        })
                            
                    if pag_v2 is not None:
                        val_pago2 = v2 + juros2
                        destinar_pagina(uf, pag_v2, val_pago2)
                        total_pdf_pago += val_pago2
                        if juros2 > 0.02:
                            relatorio_juros.append({
                                "Nota": f"{nota_excel} (Guia 2)", "UF": uf, 
                                "Valor Base": v2, "Total Pago": val_pago2, "Juros/Multa SEFAZ": juros2
                            })
                    else:
                        notas_pendentes.append({
                            "Nota": f"{nota_excel} (Guia 2)", "UF": uf, 
                            "Valor Esperado": v2, "Motivo": "Guia 2 não localizada no PDF"
                        })
                else:
                    notas_pendentes.append({
                        "Nota": nota_excel, "UF": uf, 
                        "Valor Esperado": v_total, "Motivo": "Guia não localizada no PDF"
                    })

            # --- IDENTIFICAÇÃO DE GUIAS ÓRFÃS (SOBRARAM NO PDF) ---
            guias_orfas = []
            for g in guias_disponiveis:
                if not g['usada']:
                    guias_orfas.append({
                        "Página": g['pagina_num'],
                        "UF no PDF": g['uf'],
                        "Nota no PDF": g['nota'] if g['nota'] else "(Não detectada)",
                        "Valor Guia": g['valor']
                    })

            for p_num in paginas_sem_leitura:
                guias_orfas.append({
                    "Página": p_num,
                    "UF no PDF": "(Não identificada)",
                    "Nota no PDF": "(Texto não extraível)",
                    "Valor Guia": 0.0
                })

            guias_orfas.sort(key=lambda x: x['Página'])

            # --- GERAÇÃO DOS ARQUIVOS PDF NA MEMÓRIA ---
            reader = PdfReader(io.BytesIO(pdf_bytes))

            def gerar_pdf_bytes(lista_paginas):
                if not lista_paginas:
                    return None
                writer = PdfWriter()
                for pag in lista_paginas:
                    writer.add_page(reader.pages[pag])
                buf = io.BytesIO()
                writer.write(buf)
                return buf.getvalue()

            st.session_state.pdf_itau_arquivo = gerar_pdf_bytes(paginas_itau_arquivo)
            st.session_state.pdf_itau_fisico  = gerar_pdf_bytes(paginas_itau_fisico)
            st.session_state.pdf_bradesco     = gerar_pdf_bytes(paginas_bradesco)
            st.session_state.pdf_bb           = gerar_pdf_bytes(paginas_bb)
            st.session_state.pdf_todas_fisicas = gerar_pdf_bytes(paginas_todas_fisicas)

            st.session_state.qtd_itau_arquivo = len(paginas_itau_arquivo)
            st.session_state.qtd_itau_fisico  = len(paginas_itau_fisico)
            st.session_state.qtd_bradesco     = len(paginas_bradesco)
            st.session_state.qtd_bb           = len(paginas_bb)
            st.session_state.qtd_todas_fisicas = len(paginas_todas_fisicas)
            
            st.session_state.total_excel_base = total_excel_base
            st.session_state.total_pdf_pago   = total_pdf_pago
            st.session_state.relatorio_juros  = relatorio_juros
            st.session_state.resumo_bancos    = resumo_bancos
            st.session_state.notas_pendentes  = notas_pendentes
            st.session_state.guias_orfas      = guias_orfas
            
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
        st.success("✅ Perfeito! Todos os valores do PDF bateram com a planilha base.")
    else:
        col3.metric("Juros/Acréscimos", f"+ R$ {diferenca:,.2f}", delta_color="inverse")
    
    # --------------------------------------------------------
    # TABELINHA DE DETALHAMENTO DOS BANCOS
    # --------------------------------------------------------
    st.write("---")
    st.subheader("📑 Detalhamento por Banco (Físicas e Arquivo)")
    
    df_bancos = pd.DataFrame([
        {"Banco": "Itau Arquivo", "Qtd": st.session_state.resumo_bancos["Itau Arquivo"]["Qtd"], "Valor": st.session_state.resumo_bancos["Itau Arquivo"]["Valor"]},
        {"Banco": "Restantes Itaú (Físicas)", "Qtd": st.session_state.resumo_bancos["Restantes Itaú"]["Qtd"], "Valor": st.session_state.resumo_bancos["Restantes Itaú"]["Valor"]},
        {"Banco": "Bradesco", "Qtd": st.session_state.resumo_bancos["Bradesco"]["Qtd"], "Valor": st.session_state.resumo_bancos["Bradesco"]["Valor"]},
        {"Banco": "Banco do Brasil", "Qtd": st.session_state.resumo_bancos["BB"]["Qtd"], "Valor": st.session_state.resumo_bancos["BB"]["Valor"]}
    ])
    
    st.dataframe(df_bancos.style.format({
        "Valor": "R$ {:.2f}"
    }), use_container_width=True, hide_index=True)

    # --------------------------------------------------------
    # PAINEL DE AUDITORIA E CONFORMIDADE (NOTAS PENDENTES / ÓRFÃS / JUROS)
    # --------------------------------------------------------
    st.write("---")
    st.subheader("🛡️ Painel de Auditoria e Conformidade")

    total_pendentes = len(st.session_state.notas_pendentes)
    total_orfas     = len(st.session_state.guias_orfas)
    total_juros     = len(st.session_state.relatorio_juros)

    if total_pendentes == 0 and total_orfas == 0:
        st.success("🎉 **Lote 100% Conciliado!** Todas as notas da planilha foram localizadas e nenhuma guia sobrou no PDF.")
    
    if total_pendentes > 0:
        st.error(f"❌ **{total_pendentes} Nota(s) da Planilha NÃO encontrada(s) no PDF:**")
        df_pendentes = pd.DataFrame(st.session_state.notas_pendentes)
        st.dataframe(df_pendentes.style.format({
            "Valor Esperado": "R$ {:.2f}"
        }), use_container_width=True, hide_index=True)

    if total_orfas > 0:
        st.warning(f"⚠️ **{total_orfas} Guia(s) no PDF NÃO associada(s) à Planilha (Guias Órfãs):**")
        df_orfas = pd.DataFrame(st.session_state.guias_orfas)
        st.dataframe(df_orfas.style.format({
            "Valor Guia": "R$ {:.2f}"
        }), use_container_width=True, hide_index=True)

    if total_juros > 0:
        st.info(f"ℹ️ **{total_juros} Guia(s) identificada(s) com acréscimos/juros da SEFAZ:**")
        df_juros_report = pd.DataFrame(st.session_state.relatorio_juros)
        st.dataframe(df_juros_report.style.format({
            "Valor Base": "R$ {:.2f}",
            "Total Pago": "R$ {:.2f}",
            "Juros/Multa SEFAZ": "R$ {:.2f}"
        }), use_container_width=True, hide_index=True)

    # --------------------------------------------------------
    # DOWNLOAD DOS LOTES SEPARADOS POR BANCO
    # --------------------------------------------------------
    st.write("---")
    st.subheader("🗂️ Download dos Lotes Separados por Banco")
    
    col_b1, col_b2, col_b3, col_b4 = st.columns(4)
    
    # 1. Itaú Arquivo
    if st.session_state.qtd_itau_arquivo > 0:
        col_b1.download_button(
            label=f"📥 ITAÚ ARQUIVO\n({st.session_state.qtd_itau_arquivo} pág)", 
            data=st.session_state.pdf_itau_arquivo, 
            file_name="guias_itau_arquivo.pdf", 
            mime="application/pdf", 
            use_container_width=True
        )
    else:
        col_b1.info("Sem guias Itaú Arquivo.")

    # 2. Itaú Físico
    if st.session_state.qtd_itau_fisico > 0:
        col_b2.download_button(
            label=f"📥 ITAÚ FÍSICO\n({st.session_state.qtd_itau_fisico} pág)", 
            data=st.session_state.pdf_itau_fisico, 
            file_name="guias_itau_fisica.pdf", 
            mime="application/pdf", 
            use_container_width=True
        )
    else:
        col_b2.info("Sem guias Itaú Físico.")

    # 3. Bradesco
    if st.session_state.qtd_bradesco > 0:
        col_b3.download_button(
            label=f"📥 BRADESCO\n({st.session_state.qtd_bradesco} pág)", 
            data=st.session_state.pdf_bradesco, 
            file_name="guias_bradesco.pdf", 
            mime="application/pdf", 
            use_container_width=True
        )
    else:
        col_b3.info("Sem guias Bradesco.")

    # 4. Banco do Brasil
    if st.session_state.qtd_bb > 0:
        col_b4.download_button(
            label=f"📥 BANCO DO BRASIL\n({st.session_state.qtd_bb} pág)", 
            data=st.session_state.pdf_bb, 
            file_name="guias_banco_brasil.pdf", 
            mime="application/pdf", 
            use_container_width=True
        )
    else:
        col_b4.info("Sem guias BB.")

    # Opção Unificada de Impressão Física
    if st.session_state.qtd_todas_fisicas > 0:
        st.write("")
        st.download_button(
            label=f"🖨️ BAIXAR TODAS AS GUIAS FÍSICAS UNIFICADAS ({st.session_state.qtd_todas_fisicas} pág)", 
            data=st.session_state.pdf_todas_fisicas, 
            file_name="guias_todas_impressao_fisica.pdf", 
            mime="application/pdf", 
            use_container_width=True
        )
