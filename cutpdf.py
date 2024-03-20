import os
import fitz  # PyMuPDF
from datetime import datetime

def separar_contra_cheques(input_folder, output_folder):
    # Verifica se a pasta de saída existe, se não, cria ela
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    
    
    # Obter o mês e o ano atuais
    now = datetime.now()
    mes_ano = now.strftime('%m_%Y')
    
    
    # Itera pelos arquivos na pasta de entrada
    for filename in os.listdir(input_folder):
        if filename.endswith('.pdf'):
            pdf_path = os.path.join(input_folder, filename)

            # Abre o PDF com PyMuPDF (fitz)
            pdf_document = fitz.open(pdf_path)

            # Itera pelas páginas do PDF
            for page_num in range(len(pdf_document)):
                page = pdf_document.load_page(page_num)
                text = page.get_text()

                # Verifica se o texto contém os campos "Func.:" e "Período:" e extrai o nome do colaborador
                if 'Func.:' in text and 'Período:' in text:
                    start_index = text.find('Func.:') + len('Func.:')
                    end_index = text.find('Período:')
                    nome_colaborador = text[start_index:end_index].strip()

                    # Cria a pasta do colaborador, se ainda não existir
                    colaborador_folder = os.path.join(output_folder, nome_colaborador)
                    if not os.path.exists(colaborador_folder):
                        os.makedirs(colaborador_folder)

                    # Salva a página como um novo PDF na pasta do colaborador
                    output_path = os.path.join(colaborador_folder, f'ContraCheque_{page_num + 1}_Periodo-{mes_ano}.pdf')
                    pdf_writer = fitz.open()
                    pdf_writer.insert_pdf(pdf_document, from_page=page_num, to_page=page_num)
                    pdf_writer.save(output_path)

                    print(f'Contra cheque do colaborador {nome_colaborador} salvo em {output_path}')

            pdf_document.close()

# Exemplo de uso
input_folder = '/home/pdfconvert/Input/'
output_folder = '/home/pdfconvert/Output'
separar_contra_cheques(input_folder, output_folder)
