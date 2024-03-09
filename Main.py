import pandas as pd
import os
import glob

# Caminho para ler os caminhos
folder_path = "SRC\\Data\\RAW"

# Listar todos os arquivos do excel
excel_files = glob.glob(os.path.join(folder_path , "*.xlsx"))

if not excel_files:
    print("Nenhum arquivo .xlsx encontrado.")
else:
    #dts = Data frame = tabela na memória para guardar os conteúdos dos arquivos .xlsx
        #df_temp = Data frame temporário
    dfs = []
    for excel_file in (excel_files):
        try:

            #ler o arquivo de excel
            df_temp = pd.read_excel(excel_file)

            #extrair o nome do arquivo por meio de busca por diretorio
            file_name = os.path.basename(excel_file)

            df_temp['filename'] = file_name
            #criamos uma nova coluna chamada location
            if "brasil" in file_name.lower():
                df_temp["location"] = "br"
            elif "france" in file_name.lower(): 
                df_temp["location"] = "fr"
            elif "italian" in file_name.lower(): 
                df_temp["location"] = "it"
             
                
            #criar uma nova coluna chamada campanha
            df_temp["campaign"] = df_temp['utm_link'].str.extract(r'utm_campaign=(.*)')    

            #guarda os dados dentro de um data frame comum    
            dfs.append(df_temp)    
            print(df_temp)

        except Exception as e:
            print(f"Erro ao ler o arquivo {excel_file} : {e}")

if dfs:
        
        #concatena todas as tabelas salvas no dfs em um única tabela.
        result = pd.concat(dfs, ignore_index=True)

        #caminho de saída do arquivo final após tratamento de dados
        output_file = os.path.join('SRC', 'Data', 'Ready', 'Clean.xlsx')

        writer = pd.ExcelWriter(output_file, engine='xlsxwriter')

        #leva os dados do resultado a serem escritos no motor de excel configurado
        result.to_excel(writer, index=False)

        #salva o arquivo do excel
        writer._save()
else:
    print('Nenhum dado para ser salvo')