
        
import os
import tabula
import pandas as pd
import re
os.environ["JAVA_HOME"] = r"C:\Program Files\Java\jdk-22"

def convert_pdfs_to_excel(input_folder, output_folder):
    # Verifica se a pasta de saída existe e cria se não existir
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # Lista todos os arquivos PDF na pasta de entrada
    pdf_files = [file for file in os.listdir(input_folder) if file.endswith('.pdf')]
    arquivos_ordenados = sorted(pdf_files, key=lambda x: int(x.split("_Parte")[1].split(".")[0]))
    # Obtem o total de arquivos PDF
    total_files = len(pdf_files)

    # Variável para armazenar o nome do primeiro arquivo processado
    first_file_name = None
    last_saved_file_name = None

    # Array para armazenar os valores de excel_data.iloc[1, 0]
    refs = []
    descricoes = []

    # Loop através de todos os arquivos na pasta de entrada
    for i, file_name in enumerate(arquivos_ordenados):

        # Gera o caminho completo para o arquivo PDF
        pdf_path = os.path.join(input_folder, file_name)

        # Verifica se o índice do arquivo é ímpar
        # Processa o arquivo e lê o Excel
        # Gera o nome do arquivo Excel de saída (substitui a extensão .pdf por .xlsx)
        excel_file_name = os.path.splitext(file_name)[0] + '.xlsx'
        excel_path = os.path.join(output_folder, excel_file_name)

        # Lê as tabelas do arquivo PDF
        tables = tabula.read_pdf(pdf_path, pages='all')

        # Escreve as tabelas em um arquivo Excel
        with pd.ExcelWriter(excel_path) as writer:
            for j, table in enumerate(tables):
                table.to_excel(writer, sheet_name=f'Sheet{j+1}', index=False)

        # Lê as informações relevantes do arquivo Excel
        excel_data = pd.read_excel(excel_path, sheet_name='Sheet1', header=None)
        peca_pattern = re.compile(r'Peça:', re.IGNORECASE)
        
        desc_atual = excel_data.iloc[0, 1] if excel_data.shape[1] > 1 and not pd.isna(excel_data.iloc[0, 1]) else ""
  
        # Verifica se "Tema:" existe em qualquer lugar da tabela
        tema_found = False
        for row in excel_data.iterrows():
            for cell in row[1]:
                if isinstance(cell, str) and 'Tema:' in cell:
                    tema_found = True
                    break
            if tema_found:
                break

        if tema_found:
                print(excel_data.iloc[1, 0])
                ref_atual = excel_data.iloc[1, 0]
                
                # Adiciona o valor de excel_data.iloc[1, 0] ao array refs
                refs.append(excel_data.iloc[1, 0])
                descricoes.append(desc_atual)
                
                tema = extract_theme(excel_data.iloc[1, 1].split('Tema:')[1])

                # Verifica se a linha contém a informação da peça e realiza o tratamento
                if peca_pattern.search(str(excel_data.iloc[2, 0])):
                    peca = excel_data.iloc[2, 0].split('Peça:')[1].strip().upper()
                else:
                    peca = excel_data.iloc[3, 0].split('Peça:')[1].strip().upper()

                # Substitui a barra por hífen no nome da peça
                peca = peca.replace('/', '-')
                    
                ref_match = re.search(r'Ref: (.+)', excel_data.iloc[1, 0])
                if ref_match:
                    ref = ref_match.group(1).replace(' ', '_').replace('/', '-')
                else:
                    ref = ""

                # Gera o novo nome do arquivo PDF
                new_file_name = f"{tema}_{peca}_{ref}.pdf"

                # Verifica as condições especificadas
                if len(refs) > 1 and ref_atual == refs[-2] and desc_atual != descricoes[-2]:
                        # Ação a ser executada se as condições forem verdadeiras
                        print("Condições atendidas: ref_atual é igual ao último valor em refs e desc_atual é diferente do último valor em descricoes.")
                        # Extrai a última palavra da descrição atual
                       
                        ultima_palavra = desc_atual.split()[-1]

                        # Adiciona a última palavra ao nome do arquivo
                        new_file_name = f"{tema}_{peca}_{ref}_{ultima_palavra}.pdf"
                        
                if len(refs) > 1 and ref_atual == refs[-2] and desc_atual == descricoes[-2]:
                    # Ação a ser executada se as condições forem verdadeiras
                    print("Condições atendidas: ref_atual é igual ao último valor em refs e desc_atual é igual do último valor em descricoes.")
                    # Extrai a última palavra da descrição atual
                    
                    ultima_palavra = desc_atual.split()[-1]

                    # Adiciona a última palavra ao nome do arquivo
                    counter = 1
                    new_file_name = f"{tema}_{peca}_{ref}_{ultima_palavra}_outra_cor_({counter}).pdf"
                    new_pdf_path = os.path.join(output_folder, new_file_name)
                    while os.path.exists(new_pdf_path):
                        counter += 1
                        new_file_name = f"{tema}_{peca}_{ref}_{ultima_palavra} ({counter}).pdf"
                        new_pdf_path = os.path.join(output_folder, new_file_name)
                        
                first_file_name = new_file_name

                # Renomeia o arquivo PDF
                new_pdf_path = os.path.join(output_folder, new_file_name)
                os.rename(pdf_path, new_pdf_path)
        else:
            if first_file_name:
                base_name = os.path.splitext(first_file_name)[0]
                counter = 2
                new_file_name = f"{base_name} ({counter}).pdf"
                while os.path.exists(os.path.join(output_folder, new_file_name)):
                    counter += 1
                    new_file_name = f"{base_name} ({counter}).pdf"
                
                new_pdf_path = os.path.join(output_folder, new_file_name)
                os.rename(pdf_path, new_pdf_path)
                
                
def extract_theme(text):
    # Remove caracteres especiais após "Tema:"
    theme = re.split(r'Tema:\s*', text)[-1].strip()

    # Remove "Estilista:" e qualquer texto após
    theme = re.split(r'Estilista:', theme)[0].strip()

    # Remove caracteres especiais
    theme = re.sub(r'[^\w\s]', '', theme)
    
    return theme.strip()


# Pasta de entrada contendo arquivos PDF
input_folder = r'C:\Users\Gabss\Desktop\wrong'

# Pasta de saída para os arquivos Excel
output_folder = r'C:\Users\Gabss\Desktop\wrong'

# Chama a função para converter os PDFs para Excel
convert_pdfs_to_excel(input_folder, output_folder)

