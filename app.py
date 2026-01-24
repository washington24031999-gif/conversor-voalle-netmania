import streamlit as st
import pdfplumber
import pandas as pd
import re
from io import BytesIO

# Configuração visual da página
st.set_page_config(page_title="Conversor Voalle", page_icon="📊")

st.title("📊 Conversor de Relatórios Voalle")
st.info("Arraste um ou mais PDFs para extrair os dados automaticamente.")

def limpar(txt):
    return " ".join(str(txt).split()).replace('|', '').strip()

def extrair_dados(pdf_files):
    dados = []
    for f in pdf_files:
        with pdfplumber.open(f) as pdf:
            for page in pdf.pages:
                table = page.extract_table()
                if not table: continue
                
                buffer = None
                for row in table:
                    if not row[0] or "Cliente" in str(row[0]): continue
                    
                    # Junta linhas quebras (ex: A3 Topografia)
                    if "Contrato" not in str(row[0]) and buffer is None:
                        buffer = row
                        continue
                    elif buffer:
                        row = [f"{limpar(buffer[i])} {limpar(row[i])}" for i in range(len(row))]
                        buffer = None

                    c1, c2, c4 = limpar(row[0]), limpar(row[1]), limpar(row[3]) if len(row) > 3 else ""

                    # Extração com lógica de separação total
                    cliente = re.split(r"Contrato", c1, flags=re.I)[0].strip()
                    contrato = re.search(r"n[°º²:#\s.]*(\d+)", c1)
                    data = re.search(r"(\d{2}/\d{2}/\d{4})", c1)
                    
                    local = re.search(r"Local:\s*(.*?)(?=Tipo de|$)", c2)
                    t_contrato = re.search(r"Tipo de Contrato:\s*(.*?)(?=Tipo de|$)", c2)
                    t_cobranca = re.search(r"Tipo de Cobrança:\s*(.*?)(?=Vendedor|$)", c2)
                    vendedor = re.search(r"Vendedor:\s*(.*)", c2)
                    valor = re.search(r"Total em Atraso:\s*R\$\s*([\d.,]+)", c4, re.I)

                    dados.append({
                        "Cliente": cliente, "Contrato": contrato.group(1) if contrato else "",
                        "Data Ativação": data.group(1) if data else "",
                        "Local": local.group(1).strip() if local else "",
                        "Vendedor": vendedor.group(1).strip() if vendedor else "",
                        "Total em Atraso": valor.group(1) if valor else "0,00"
                    })
    return pd.DataFrame(dados)

# Interface de Upload
files = st.file_uploader("Suba os arquivos aqui", type="pdf", accept_multiple_files=True)

if files:
    if st.button("Processar PDFs"):
        df = extrair_dados(files)
        st.dataframe(df) # Mostra a tabela na tela
        
        # Gerador do botão de download
        output = BytesIO()
        df.to_excel(output, index=False)
        st.download_button(label="📥 Descarregar Excel", data=output.getvalue(), file_name="Dados_Extraidos.xlsx")