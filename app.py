import streamlit as st
import pdfplumber
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="Conversor Voalle Pro", page_icon="📊")

st.title("📊 Conversor de Relatórios Voalle")
st.info("Padronização automática de PDF e Excel.")

def limpar(txt):
    return " ".join(str(txt).split()).replace('|', '').strip()

def extrair_logica_voalle(texto_c1, texto_c2, texto_c4):
    """Função central para garantir que PDF e Excel sigam a mesma regra"""
    cliente = re.split(r"Contrato", texto_c1, flags=re.I)[0].strip()
    contrato = re.search(r"n[°º²:#\s.]*(\d+)", texto_c1)
    data = re.search(r"(\d{2}/\d{2}/\d{4})", texto_c1)
    local = re.search(r"Local:\s*(.*?)(?=Tipo de|$)", texto_c2)
    vendedor = re.search(r"Vendedor:\s*(.*)", texto_c2)
    valor = re.search(r"Total em Atraso:\s*R\$\s*([\d.,]+)", texto_c4, re.I)

    return {
        "Cliente": cliente, 
        "Contrato": contrato.group(1) if contrato else "",
        "Data Ativação": data.group(1) if data else "",
        "Local": local.group(1).strip() if local else "",
        "Vendedor": vendedor.group(1).strip() if vendedor else "",
        "Total em Atraso": valor.group(1) if valor else "0,00"
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
                        if "Contrato" not in str(row[0]) and buffer is None:
                            buffer = row
                            continue
                        elif buffer:
                            row = [f"{limpar(buffer[i])} {limpar(row[i])}" for i in range(len(row))]
                            buffer = None
                        
                        # Usa a lógica padronizada
                        res = extrair_logica_voalle(limpar(row[0]), limpar(row[1]), limpar(row[3]) if len(row) > 3 else "")
                        todos_dados.append(res)

        elif extensao in ["xlsx", "xls", "csv"]:
            if extensao == "csv":
                df_temp = pd.read_csv(f, sep=None, engine='python', encoding='utf-8-sig')
            else:
                df_temp = pd.read_excel(f)
            
            # Se o Excel vier com as colunas 'bagunçadas' como no PDF, tratamos linha a linha
            for _, row in df_temp.iterrows():
                # Tentamos identificar se os dados estão concentrados nas primeiras colunas
                c1 = str(row.iloc[0]) if len(row) > 0 else ""
                c2 = str(row.iloc[1]) if len(row) > 1 else ""
                c4 = str(row.iloc[3]) if len(row) > 3 else ""
                
                # Aplica a mesma limpeza e regex do PDF
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
            # Trocado para openpyxl para evitar o erro ModuleNotFoundError
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_final.to_excel(writer, index=False)
            
            st.download_button(
                label="📥 Baixar Excel Padronizado", 
                data=output.getvalue(), 
                file_name="Relatorio_Voalle_Unificado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
