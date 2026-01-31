def extrair_dados(uploaded_files):
    todos_dados = []
    
    for f in uploaded_files:
        extensao = f.name.split('.')[-1].lower()
        
        if extensao == "pdf":
            todos_dados.extend(processar_pdf(f))
        
        elif extensao in ["xlsx", "xls", "csv"]:
            if extensao == "csv":
                try:
                    # Garante que a leitura comece do início do arquivo
                    f.seek(0) 
                    df_temp = pd.read_csv(f, sep=",")
                    # Se o pandas ler mas não encontrar colunas reais, força o erro para tentar o próximo separador
                    if len(df_temp.columns) <= 1:
                        raise Exception("Provável separador errado")
                except:
                    f.seek(0) # Volta ao início novamente antes da segunda tentativa
                    df_temp = pd.read_csv(f, sep=";")
            else:
                f.seek(0)
                df_temp = pd.read_excel(f)
            
            # Limpeza básica: remove linhas totalmente vazias que o Excel costuma gerar
            df_temp = df_temp.dropna(how='all')
            
            dict_dados = df_temp.to_dict('records')
            todos_dados.extend(dict_dados)
            
    return pd.DataFrame(todos_dados)
