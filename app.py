import streamlit as st
import pdfplumber
import pandas as pd
import re
from io import BytesIO

# Configuração visual da página
st.set_page_config(page_title="Conversor Voalle Multi", page_icon="📊")

st.title("📊 Conversor de Relatórios Voalle")
st.info("Arraste PDFs ou arquivos Excel/CSV para extrair os dados.")

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
        
        if extensao == "pdf":
            todos_dados.extend(processar_pdf(f))
        
        elif extensao in ["xlsx", "xls", "csv"]:
            if extensao == "csv":
                # Tenta ler com vírgula, se falhar tenta ponto e vírgula
                try:
                    df_temp = pd.read_csv(f, sep=",")
                except:
                    df_temp = pd.read_csv(f, sep=";")
            else:
                df_temp = pd.read_excel(f)
            
            # Converte o DataFrame do Excel para o formato de dicionário da nossa lista
            # Isso garante que mesmo que o Excel tenha colunas extras, pegamos só o que importa
            dict_dados = df_temp.to_dict('records')
            todos_dados.extend(dict_dados)
            
    return pd.DataFrame(todos_dados)

# Interface de Upload atualizada para aceitar mais formatos
files = st.file_uploader(
    "Suba os arquivos aqui (PDF, Excel ou CSV)", 
    type=["pdf", "xlsx", "xls", "csv"], 
    accept_multiple_files=True
)

if files:
    if st.button("Processar Arquivos"):
        df_final = extrair_dados(files)
        
        if not df_final.empty:
            st.success(f"Sucesso! {len(df_final)} registros processados.")
            st.dataframe(df_final)
            
            # Gerador do botão de download
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_final.to_excel(writer, index=False)
            
            st.download_button(
                label="📥 Baixar Excel Consolidado", 
                data=output.getvalue(), 
                file_name="Relatorio_Voalle_Consolidado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("Nenhum dado encontrado nos arquivos enviados.")
