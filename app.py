import streamlit as st
import pandas as pd
import pdfplumber
import re
from fpdf import FPDF, XPos, YPos
import io
import os

# ==========================================
# CONFIGURAÇÃO INICIAL
# ==========================================
st.set_page_config(
    page_title="Conciliador: Acervo Bibliográfico",
    page_icon="📚",
    layout="wide"
)

# Oculta marcas do Streamlit e a barra lateral
hide_streamlit_style = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            header {visibility: hidden;}
            [data-testid="stSidebar"] {display: none;}
            [data-testid="collapsedControl"] {display: none;}
            </style>
            """
st.markdown(hide_streamlit_style, unsafe_allow_html=True)

# Botão para retornar à tela inicial solto no topo da tela
st.page_link("Menu_principal.py", label="⬅️ Voltar ao Menu Inicial")

# ==========================================
# FUNÇÕES E CLASSES (BASTIDORES)
# ==========================================
def formatar_real(valor):
    sinal = "-" if valor < -0.001 else ""
    return f"{sinal}{abs(valor):,.2f}".replace(',', '_').replace('.', ',').replace('_', '.')

def limpar_valor_flex(v):
    # Pega apenas números, pontos e vírgulas
    v = re.sub(r'[^\d\.,]', '', str(v))
    if not v: return 0.0
    
    # Tratamento para casos onde o centavo vem separado por ponto ex: 3.074.625.29
    if len(v) >= 3 and v[-3] in ['.', ',']:
        inteiro = v[:-3].replace('.', '').replace(',', '')
        decimal = v[-2:]
        return float(f"{inteiro}.{decimal}")
    else:
        # Se não tiver casas decimais claras, limpa tudo e converte
        return float(v.replace('.', '').replace(',', '.'))

def extrair_valor_pdf(pdf_bytes, texto_busca, is_dep=False):
    texto_completo = ""
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for page in pdf.pages:
                texto_completo += page.extract_text() + "\n"
                
        for line in texto_completo.split('\n'):
            line = line.strip().replace('"', '') # Remove aspas caso a leitura traga
            
            # Procura exatamente a linha que começa com o Mês (ex: Janeiro) ou Mês/Ano (ex: 01/2026)
            if line.upper().startswith(texto_busca.upper()):
                # Pega todos os blocos numéricos da linha
                matches = re.findall(r'[\d\.,]+', line)
                
                # Só aceita se a linha tiver números suficientes (garante que é a linha da tabela e não um título solto)
                if len(matches) >= 2:
                    # O saldo final/acumulado é sempre o último bloco numérico extraído da linha
                    return limpar_valor_flex(matches[-1])
    except Exception:
        pass
    return 0.0

class PDF_Report(FPDF):
    def header(self):
        self.set_font('helvetica', 'B', 12)
        self.cell(0, 10, 'Relatório de Conferência: Acervo Bibliográfico x Pergamum', align='C', new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        self.ln(5)
    def footer(self):
        self.set_y(-15)
        self.set_font('helvetica', 'I', 8)
        self.cell(0, 10, f'Página {self.page_no()}', align='C')

# ==========================================
# INTERFACE DO USUÁRIO
# ==========================================
st.title("📚 Conciliador: Acervo Bibliográfico")

with st.expander("📘 GUIA DE USO (Clique para abrir)", expanded=False):
    st.markdown("📌 **Orientações de Uso**")
    st.markdown("""
    1. Selecione o **Mês** e o **Ano** exatos que deseja conciliar.
    2. Anexe a **Planilha Excel (Conf. RMB)** e todos os **arquivos PDF (Pergamum)** de uma só vez.
    3. **Atenção aos Nomes dos PDFs:** O arquivo do Acervo deve ter o número da UG (ex: `153289.pdf`) e o da Depreciação deve ter um 'd' no final (ex: `153289d.pdf`).
    4. Clique em "Iniciar Conciliação" e aguarde o relatório.
    """)

# Seleção de Data
col_mes, col_ano = st.columns(2)
meses = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
with col_mes:
    mes_selecionado = st.selectbox("Selecione o Mês:", meses)
with col_ano:
    ano_selecionado = st.number_input("Digite o Ano:", min_value=2000, max_value=2100, value=2026, step=1)

# Constrói o texto exato que o sistema vai procurar no PDF de acordo com a seleção
mes_num = f"{meses.index(mes_selecionado) + 1:02d}"
texto_busca_acervo = mes_selecionado           # Ex: "Janeiro"
texto_busca_dep = f"{mes_num}/{ano_selecionado}" # Ex: "01/2026"

# Área de Upload Unificada
uploaded_files = st.file_uploader(
    "📂 Arraste a Planilha do Tesouro e os PDFs do Pergamum para esta área", 
    accept_multiple_files=True,
    type=['pdf', 'xlsx', 'xls', 'csv']
)

# ==========================================
# EXECUÇÃO DO SISTEMA
# ==========================================
if st.button("🚀 Iniciar Conciliação", use_container_width=True, type="primary"):
    
    if not uploaded_files:
        st.warning("⚠️ Por favor, insira seus arquivos para que possamos realizar a conciliação.")
    else:
        progresso = st.progress(0)
        status_text = st.empty()
        
        # Classificação dos arquivos
        pdfs = {f.name.lower(): f for f in uploaded_files if f.name.lower().endswith('.pdf')}
        excel_files = [f for f in uploaded_files if f.name.lower().endswith(('.xlsx', '.xls', '.csv'))]
        
        if not excel_files:
            st.error("❌ A planilha base em Excel não foi encontrada no upload.")
            st.stop()
            
        planilha_mestre = excel_files[0]
        dados_ug = {}
        logs = []

        status_text.text("Lendo os dados da Planilha Base...")
        try:
            # Tenta ler como Excel, se falhar tenta como CSV
            planilha_mestre.seek(0)
            if planilha_mestre.name.lower().endswith('.csv'):
                df = pd.read_csv(planilha_mestre)
            else:
                df = pd.read_excel(planilha_mestre, header=None)
            
            # Varrer a planilha à procura de códigos de UG (números de 6 dígitos) na primeira coluna
            for idx, row in df.iterrows():
                val0 = str(row[0]).strip()
                if val0.isdigit() and len(val0) >= 5:
                    ug = val0
                    nome = str(row[1]).strip()
                    saldo_acervo = limpar_valor_flex(row[2]) if len(row) > 2 else 0.0
                    saldo_dep = limpar_valor_flex(row[3]) if len(row) > 3 else 0.0
                    
                    dados_ug[ug] = {
                        'nome': nome,
                        'ex_acervo': saldo_acervo,
                        'ex_dep': abs(saldo_dep), # Pega valor absoluto pois no excel costuma vir negativo
                        'pdf_acervo': 0.0,
                        'pdf_dep': 0.0,
                        'achou_pdf_acervo': False,
                        'achou_pdf_dep': False
                    }
        except Exception as e:
            st.error(f"❌ Erro ao ler a estrutura da planilha: {e}")
            st.stop()

        status_text.text("Processando e cruzando os documentos PDF...")
        total_ugs = len(dados_ug)
        if total_ugs == 0:
            st.warning("⚠️ Nenhuma Unidade Gestora (UG) foi encontrada na primeira coluna da planilha.")
            st.stop()

        # Extração dos valores dos PDFs
        for i, (ug, info) in enumerate(dados_ug.items()):
            nome_pdf_acervo = f"{ug}.pdf"
            nome_pdf_dep = f"{ug}d.pdf"
            
            # Tenta achar o PDF Normal (Acervo)
            if nome_pdf_acervo in pdfs:
                info['achou_pdf_acervo'] = True
                pdfs[nome_pdf_acervo].seek(0)
                info['pdf_acervo'] = extrair_valor_pdf(pdfs[nome_pdf_acervo].read(), texto_busca_acervo, is_dep=False)
            else:
                logs.append(f"⚠️ UG {ug}: Faltou o PDF do Acervo ({nome_pdf_acervo}).")

            # Tenta achar o PDF de Depreciação (com 'd')
            if nome_pdf_dep in pdfs:
                info['achou_pdf_dep'] = True
                pdfs[nome_pdf_dep].seek(0)
                info['pdf_dep'] = extrair_valor_pdf(pdfs[nome_pdf_dep].read(), texto_busca_dep, is_dep=True)
            else:
                logs.append(f"⚠️ UG {ug}: Faltou o PDF de Depreciação ({nome_pdf_dep}).")
                
            progresso.progress((i + 1) / total_ugs)

        # ==========================================
        # GERAÇÃO DO RELATÓRIO E EXIBIÇÃO
        # ==========================================
        pdf_out = PDF_Report()
        pdf_out.add_page()
        
        st.markdown("---")
        st.subheader("🔍 Resultados da Análise")
        
        total_ex_acervo = total_ex_dep = total_pdf_acervo = total_pdf_dep = 0.0

        for ug, info in dados_ug.items():
            dif_acervo = info['pdf_acervo'] - info['ex_acervo']
            dif_dep = info['pdf_dep'] - info['ex_dep']
            
            total_ex_acervo += info['ex_acervo']
            total_pdf_acervo += info['pdf_acervo']
            total_ex_dep += info['ex_dep']
            total_pdf_dep += info['pdf_dep']
            
            tem_erro = abs(dif_acervo) > 0.05 or abs(dif_dep) > 0.05
            
            # Geração do bloco no PDF
            pdf_out.set_font("helvetica", 'B', 10)
            pdf_out.set_fill_color(240, 240, 240)
            pdf_out.cell(0, 8, text=f"Unidade Gestora: {ug} - {info['nome'][:60]}", border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT, fill=True)
            
            pdf_out.set_font("helvetica", 'B', 8)
            pdf_out.set_fill_color(220, 230, 241)
            pdf_out.cell(46, 7, "Conta", 1, fill=True)
            pdf_out.cell(48, 7, "Saldo PDF (Pergamum)", 1, fill=True, align='C')
            pdf_out.cell(48, 7, "Saldo Excel (SIAFI)", 1, fill=True, align='C')
            pdf_out.cell(48, 7, "Diferença", 1, fill=True, align='C', new_x=XPos.LMARGIN, new_y=YPos.NEXT)
            
            pdf_out.set_font("helvetica", '', 8)
            
            # Linha 1: Acervo
            pdf_out.cell(46, 7, "Acervo Bibliográfico", 1)
            pdf_out.cell(48, 7, f"R$ {formatar_real(info['pdf_acervo'])}", 1, align='R')
            pdf_out.cell(48, 7, f"R$ {formatar_real(info['ex_acervo'])}", 1, align='R')
            if abs(dif_acervo) > 0.05: pdf_out.set_text_color(200, 0, 0)
            pdf_out.cell(48, 7, f"R$ {formatar_real(dif_acervo)}", 1, align='R', new_x=XPos.LMARGIN, new_y=YPos.NEXT)
            pdf_out.set_text_color(0, 0, 0)
            
            # Linha 2: Depreciação
            pdf_out.cell(46, 7, "Depreciação Acumulada", 1)
            pdf_out.cell(48, 7, f"R$ {formatar_real(info['pdf_dep'])}", 1, align='R')
            pdf_out.cell(48, 7, f"R$ {formatar_real(info['ex_dep'])}", 1, align='R')
            if abs(dif_dep) > 0.05: pdf_out.set_text_color(200, 0, 0)
            pdf_out.cell(48, 7, f"R$ {formatar_real(dif_dep)}", 1, align='R', new_x=XPos.LMARGIN, new_y=YPos.NEXT)
            pdf_out.set_text_color(0, 0, 0)
            
            pdf_out.ln(5)
            
            # Exibição na Tela (Expander se tiver erro)
            if tem_erro:
                with st.expander(f"⚠️ UG {ug}: Divergências Encontradas", expanded=True):
                    df_view = pd.DataFrame([
                        {"Conta": "Acervo Bibliográfico", "PDF": info['pdf_acervo'], "Excel": info['ex_acervo'], "Diferença": dif_acervo},
                        {"Conta": "Depreciação Acumulada", "PDF": info['pdf_dep'], "Excel": info['ex_dep'], "Diferença": dif_dep}
                    ])
                    # Estilização básica para a tela
                    st.dataframe(df_view.style.format({"PDF": "R$ {:,.2f}", "Excel": "R$ {:,.2f}", "Diferença": "R$ {:,.2f}"}))

        # Totais Finais
        dif_total_acervo = total_pdf_acervo - total_ex_acervo
        dif_total_dep = total_pdf_dep - total_ex_dep
        
        st.markdown("### Resumo Geral da Conciliação")
        c1, c2, c3 = st.columns(3)
        c1.metric("Diferença Total (Acervo)", f"R$ {dif_total_acervo:,.2f}", delta_color="inverse" if abs(dif_total_acervo) > 0.05 else "normal")
        c2.metric("Diferença Total (Depreciação)", f"R$ {dif_total_dep:,.2f}", delta_color="inverse" if abs(dif_total_dep) > 0.05 else "normal")
        
        status_text.success("Conciliação concluída com sucesso!")
        progresso.empty()
        
        if logs:
            with st.expander("⚠️ Avisos de Ficheiros Ausentes", expanded=False):
                for log in logs: st.write(log)
        
        # Download do PDF
        try:
            pdf_bytes = bytes(pdf_out.output())
            st.download_button(
                label="📄 BAIXAR RELATÓRIO DE CONCILIAÇÃO (.PDF)", 
                data=pdf_bytes, 
                file_name=f"RELATORIO_ACERVO_BIBLIOGRAFICO_{mes_selecionado}_{ano_selecionado}.pdf", 
                mime="application/pdf", 
                type="primary", 
                use_container_width=True
            )
        except Exception as e:
            st.error(f"Erro ao gerar o download: {e}")
