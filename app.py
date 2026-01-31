import streamlit as st
import pdfplumber
import pandas as pd
import re
from io import BytesIO

# Configuração visual da página
st.set_page_config(page_title="Conversor Voalle Multi", page_icon="📊")

st.title("📊 Conversor de Relatórios Voalle")
st.info("Arraste PDFs, arquivos Excel (.xlsx) ou CSV para converter.")

def limpar(txt):
    return " ".join(str(txt).split()).replace('|', '').strip()

def processar_pdf(f):
    dados_pdf = []
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

                c1, c2, c4 = limpar(row[0]), limpar(row[1]), limpar(row[3]) if len(row) > 3 else ""

                cliente = re.split(r"Contrato", c1, flags=re.I)[0].strip()
                contrato = re.search(r"n[°º²:#\s.]*(\d+)", c1)
                data = re.search(r"(\d{2}/\d{2}/\d{4})", c1)
                local = re.search(r"Local:\s*(.*?)(?=Tipo de|$)", c2)
                vendedor = re.search(r"Vendedor:\s*(.*)", c2)
                valor = re.search(r"Total em Atraso:\s*R\$\s*([\d.,]+)", c4, re.I)

                dados_pdf.append({
                    "Cliente": cliente, 
                    "Contrato": contrato.group(1) if contrato else "",
                    "Data Ativação": data.group(1) if data else "",
                    "Local": local.group(1).strip() if local else "",
                    "Vendedor": vendedor.group(1).strip() if vendedor else "",
                    "Total em Atraso": valor.group(1) if valor else "0,00"
                })
    return dados_pdf

def extrair_dados(uploaded_files):
    todos_dados = []
    
    for f in uploaded_files:
        extensao = f.name.split('.')[-1].lower()
        f.seek(0) # Reset essencial para evitar EmptyDataError
        
        try:
            if extensao == "pdf":
                todos_dados.extend(processar_pdf(f))
            
            elif extensao == "csv":
                # sep=None faz o pandas detectar automaticamente se é , ou ;
                df_temp = pd.read_csv(f, sep=None, engine='python', encoding='utf-8-sig')
                todos_dados.extend(df_temp.to_dict('records'))
            
            elif extensao in ["xlsx", "xls"]:
                df_temp = pd.read_excel(f)
                todos_dados.extend(df_temp.to_dict('records'))
        except Exception as e:
            st.error(f"Erro ao processar o arquivo {f.name}: {e}")
            
    return pd.DataFrame(todos_dados)

# Interface de Upload
files = st.file_uploader(
    "Suba seus arquivos", 
    type=["pdf", "xlsx", "xls", "csv"], 
    accept_multiple_files=True
)

if files:
    if st.button("🚀 Processar Arquivos"):
        df_final = extrair_dados(files)
        
        if not df_final.empty:
            st.success(f"Pronto! {len(df_final)} registros encontrados.")
            st.dataframe(df_final)
            
            # Exportação para Excel
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_final.to_excel(writer, index=False)
            
            st.download_button(
                label="📥 Baixar Excel Consolidado", 
                data=output.getvalue(), 
                file_name="Relatorio_Consolidado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("Nenhum dado pôde ser extraído. Verifique o formato dos arquivos.")
