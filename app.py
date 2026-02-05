import streamlit as st
import pdfplumber
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="Conversor Voalle Pro", page_icon="📊")

st.title("📊 Conversor de Relatórios Voalle")
st.info("Padronização automática de PDF e Excel com extração detalhada de títulos e solicitações.")

def limpar(txt):
    return " ".join(str(txt).split()).replace('|', '').strip()

def extrair_logica_voalle(texto_c1, texto_c2, texto_c4):
    """Função central atualizada para novas colunas"""
    # Dados básicos
    cliente = re.split(r"Contrato", texto_c1, flags=re.I)[0].strip()
    contrato = re.search(r"n[°º²:#\s.]*(\d+)", texto_c1)
    data = re.search(r"(\d{2}/\d{2}/\d{4})", texto_c1)
    
    # Tipo de Contrato e Cobrança
    tipo_contrato = re.search(r"Tipo de Contrato:\s*(.*?)(?=Tipo de Cobrança|$)", texto_c2, re.I)
    tipo_cobranca = re.search(r"Tipo de Cobrança:\s*(.*)", texto_c2, re.I)
    
    # Local e Vendedor
    local = re.search(r"Local:\s*(.*?)(?=Tipo de|$)", texto_c2)
    vendedor = re.search(r"Vendedor:\s*(.*)", texto_c2)

    # Solicitações (Formatando para Total:X, Em aberto:Y, Em atraso:Z)
    s_total = re.search(r"Total:\s*(\d+)", texto_c4)
    s_aberto = re.search(r"Em Aberto:\s*(\d+)", texto_c4)
    s_atraso = re.search(r"Em Atraso:\s*(\d+)", texto_c4)
    
    solicitacoes = f"Total:{s_total.group(1) if s_total else 0}, Em aberto:{s_aberto.group(1) if s_aberto else 0}, Em atraso:{s_atraso.group(1) if s_atraso else 0}"

    # Financeiro (Novas 3 colunas extraídas da Coluna D/texto_c4)
    t_aberto = re.search(r"Títulos em Aberto:\s*R\$\s*([\d.,]+)", texto_c4, re.I)
    t_atraso_qtd = re.search(r"Títulos em Atraso:\s*R\$\s*([\d.,]+)", texto_c4, re.I)
    v_total_atraso = re.search(r"Total em Atraso:\s*R\$\s*([\d.,]+)", texto_c4, re.I)

    return {
        "Cliente": cliente, 
        "Contrato": contrato.group(1) if contrato else "",
        "Data Ativação": data.group(1) if data else "",
        "Tipo de Contrato": tipo_contrato.group(1).strip() if tipo_contrato else "",
        "Tipo de Cobrança": tipo_cobranca.group(1).strip() if tipo_cobranca else "",
        "Local": local.group(1).strip() if local else "",
        "Vendedor": vendedor.group(1).strip() if vendedor else "",
        "Solicitações": solicitacoes,
        "Títulos em Aberto": t_aberto.group(1) if t_aberto else "0,00",
        "Títulos em Atraso": t_atraso_qtd.group(1) if t_atraso_qtd else "0,00",
        "Total em Atraso": v_total_atraso.group(1) if v_total_atraso else "0,00"
    }

def extrair_dados(uploaded_files):
    todos_dados = []
    
    for f in uploaded_files:
        extensao = f.name.split('.')[-1].lower()
        f.seek(0)
        
        if extensao == "pdf":
            with pdfplumber.open(f) as pdf:
                for page in pdf.pages:
                    table = page.extract_table()
                    if not table: continue
                    buffer = None
                    for row in table:
                        if not row[0] or "Cliente" in str(row[0]): continue
                        
                        # Lógica de buffer para linhas quebradas no PDF
                        if "Contrato" not in str(row[0]) and buffer is None:
                            buffer = row
                            continue
                        elif buffer:
                            row = [f"{limpar(buffer[i])} {limpar(row[i])}" for i in range(len(row))]
                            buffer = None
                        
                        res = extrair_logica_voalle(limpar(row[0]), limpar(row[1]), limpar(row[3]) if len(row) > 3 else "")
                        todos_dados.append(res)

        elif extensao in ["xlsx", "xls", "csv"]:
            if extensao == "csv":
                df_temp = pd.read_csv(f, sep=None, engine='python', encoding='utf-8-sig')
            else:
                df_temp = pd.read_excel(f)
            
            for _, row in df_temp.iterrows():
                c1 = str(row.iloc[0]) if len(row) > 0 else ""
                c2 = str(row.iloc[1]) if len(row) > 1 else ""
                c4 = str(row.iloc[3]) if len(row) > 3 else ""
                
                res = extrair_logica_voalle(limpar(c1), limpar(c2), limpar(c4))
                todos_dados.append(res)
            
    return pd.DataFrame(todos_dados)

files = st.file_uploader("Arquivos (PDF/Excel)", type=["pdf", "xlsx", "xls", "csv"], accept_multiple_files=True)

if files:
    if st.button("🚀 Processar e Padronizar"):
        df_final = extrair_dados(files)
        
        if not df_final.empty:
            st.success(f"Pronto! {len(df_final)} registros processados.")
            st.dataframe(df_final)
            
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_final.to_excel(writer, index=False)
            
            st.download_button(
                label="📥 Baixar Excel Padronizado", 
                data=output.getvalue(), 
                file_name="Relatorio_Voalle_Unificado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
