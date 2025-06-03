from flask import Flask, render_template, request
import threading
import time
import tempfile
import os
import ast
import re
import logging
import unicodedata
from datetime import datetime
import pdfplumber
from PyPDF2 import PdfReader
from pdfminer.high_level import extract_text_to_fp
from io import StringIO
import shutil
from html import escape
import sys
import socket


app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # Limite de tamanho para uploads (100MB)


ALLOWED_EXTENSIONS = {'pdf'}


# Função para obter o caminho correto no ambiente empacotado
def obter_caminho_recurso(relativo):
    if hasattr(sys, '_MEIPASS'):
        # Se estiver rodando o executável, os arquivos estão no diretório temporário
        return os.path.join(sys._MEIPASS, relativo)
    else:
        # Se estiver rodando no ambiente de desenvolvimento
        return os.path.join(os.path.abspath("."), relativo)

# Ajustar a localização dos templates e arquivos estáticos
app.template_folder = obter_caminho_recurso("templates")
app.static_folder = obter_caminho_recurso("static")

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


# Configura o logging para registrar informações e erros
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(message)s')

# Função que extrai o texto de um PDF utilizando PDFMiner e mantém o layout original (SICAF2)
def extract_text_with_pdfminer_layout(pdf_path):
    try:
        # Usar StringIO para capturar a saída da função extract_text_to_fp
        output_string = StringIO()

        # Extrair o texto do PDF mantendo o layout original
        with open(pdf_path, 'rb') as arquivo_pdf:
            extract_text_to_fp(arquivo_pdf, output_string, laparams=None)

        # Pegar o valor do texto extraído
        texto_extraido = output_string.getvalue()

        # Fechar o StringIO
        output_string.close()

        # Remove todos os espaços utilizando replace
        texto_extraido = texto_extraido.replace(" ", "")

        # Remove linhas em branco
        texto_extraido = "\n".join([line for line in texto_extraido.splitlines() if line.strip()])

        return texto_extraido
    except Exception as e:
        logging.error(f"Erro ao extrair texto com PDFMiner mantendo layout do PDF {pdf_path}: {e}")
        return ""


# Função que extrai o texto de um PDF e faz ajustes no formato (adição de espaço após "Formato:" e remoção de linhas em branco)
def extract_text_with_format_adjustment(pdf_path):
    try:
        with pdfplumber.open(pdf_path) as pdf:
            # Concatena o texto de todas as páginas
            text = "\n".join([page.extract_text() for page in pdf.pages])
        # Adiciona um espaço após "Formato:" se não houver
        text = re.sub(r'(Formato:)(\S)', r'\1 \2', text)
        # Remove linhas em branco
        text = "\n".join([line for line in text.splitlines() if line.strip()])
        return text
    except Exception as e:
        logging.error(f"Erro ao extrair texto do PDF {pdf_path}: {e}")
        return ""


# Função que extrai o texto de um PDF usando PyPDF e faz ajustes no formato (adição de espaço após "Formato:" e remoção de linhas em branco)
def extract_text_with_format_adjustment_py(pdf_path):
    try:
        # Abre o PDF e extrai o texto usando PyPDF
        with open(pdf_path, "rb") as file:
            reader = PdfReader(file)
            # Concatena o texto de todas as páginas
            text = "\n".join([page.extract_text() for page in reader.pages])

        # Adiciona um espaço após "Formato:" se não houver
        text = re.sub(r'(Formato:)(\S)', r'\1 \2', text)
        # Remove linhas em branco
        text = "\n".join([line for line in text.splitlines() if line.strip()])

        return text
    except Exception as e:
        logging.error(f"Erro ao extrair texto do PDF {pdf_path}: {e}")
        return ""


# Função para extrair o valor de um campo específico, com suporte para várias opções, como data, números, e delimitadores
def extract_field_value(text, field_names, below=False, below_lines=1, first_n_chars=None, date_only=False,
                        exclude_pattern=None, exclude_numbers=False, after_dash=False, stop_before=None,
                        stop_after=None, only_numbers=False, line_range=None, check_next_line_if_empty=False,
                        skip_empty_lines=True, split_by=None):
    lines = text.split('\n')

    # Se for fornecido um intervalo de linhas, ele ajusta a captura de linhas conforme necessário
    if line_range is not None:
        if isinstance(line_range, int):
            lines = [lines[line_range - 1]]  # Captura apenas a linha especificada (indexada a partir de 1)
        else:
            lines = lines[line_range[0] - 1:line_range[1]]  # Ajusta intervalo de linhas

    # Procura o campo e extrai o valor dele, com várias condições (como linhas abaixo, parar antes de algo, etc.)
    for field_name in field_names:
        for i, line in enumerate(lines):
            start_index = line.find(field_name)
            if start_index != -1:
                if below and i + below_lines < len(lines):  # Captura a linha abaixo, considerando below_lines
                    field_value = lines[i + below_lines].strip()
                    if skip_empty_lines:
                        # Pula linhas vazias
                        while not field_value and i + below_lines < len(lines):
                            i += 1
                            field_value = lines[i + below_lines].strip()

                    if check_next_line_if_empty and not field_value and i + below_lines + 1 < len(lines):
                        # Se a linha seguinte estiver vazia e houver uma linha após ela, captura essa linha
                        field_value = lines[i + below_lines + 1].strip()
                else:
                    field_value = line[start_index + len(field_name):].strip()
                break
        else:
            continue
        break
    else:
        return None

    # Diversas opções de ajustes no valor do campo extraído
    if stop_before:
        # Se stop_before for uma lista, verifique todas as opções
        if isinstance(stop_before, list):
            for stop_word in stop_before:
                field_value = field_value.split(stop_word)[0].strip()
        else:
            # Se for uma string, aplique diretamente
            field_value = field_value.split(stop_before)[0].strip()

    if stop_after:
        stop_index = field_value.find(stop_after)
        if stop_index != -1:
            field_value = field_value[:stop_index].strip()

    if date_only:
        # Extração de apenas uma data
        date_match = re.search(r'\d{2}/\d{2}/\d{4}', field_value)
        if date_match:
            return date_match.group()

    if first_n_chars:
        field_value = field_value[:first_n_chars].strip(': ').strip()

    if exclude_pattern:
        field_value = re.sub(exclude_pattern, '', field_value).strip()

    if exclude_numbers:
        field_value = re.sub(r'\d+', '', field_value).strip()

    if after_dash:
        parts = field_value.split('-')
        if len(parts) > 1:
            field_value = parts[1].strip()
        else:
            field_value = ''

    if only_numbers:
        field_value = re.sub(r'\D+', '', field_value).strip()

    # Se split_by for fornecido, divide o valor extraído em múltiplos valores
    if split_by:
        # Dividir por 'split_by' com espaços opcionais ao redor
        field_value = [part.strip() for part in re.split(r' *' + re.escape(split_by) + r' *', field_value) if
                       part.strip()]

    return field_value.strip() if isinstance(field_value, str) else field_value


# Função para extrair múltiplos valores de campos, com suporte para captura abaixo do rótulo (similar ao extract_field_value, mas múltiplos campos)
def extract_field_values(text, field_names, line_range=None, stop_before=None, below=False, below_lines=1,
                         after_dash=False, exclude_pattern=None, avoid_duplicates_for_formato=True):
    values = []
    lines = text.split('\n')

    # Se line_range for fornecido, ajusta o intervalo de linhas para capturar
    if line_range:
        lines = lines[line_range[0] - 1:line_range[1]]  # Ajusta intervalo de linhas (indexado a partir de zero)

    # Procura o campo e extrai o valor dele, adicionando-o a uma lista
    for field_name in field_names:
        for i, line in enumerate(lines):
            if field_name in line:
                if below and i + below_lines < len(lines):
                    value = lines[i + below_lines].strip()
                else:
                    match = re.search(re.escape(field_name) + r'.*', line, re.IGNORECASE)
                    if match:
                        value = match.group()[len(field_name):].strip()
                    else:
                        value = None

                if value:
                    # Normaliza aspas
                    value = re.sub(r'[“”″\'"‘’]', '"', value)

                    if stop_before:
                        value = value.split(stop_before)[0].strip()

                    # Ajustes especiais para campos que contenham "PEÇA" ou "FORMATO"
                    if 'PECA' in field_name.upper() or 'PEÇA' in field_name.upper():
                        value = re.sub(r'-[A-Z] ', '', value).strip()
                        value = re.sub(r'- [A-Z] ', '', value).strip()
                        value = value.split('FORMATO')[0].strip()

                    if 'FORMATO' in field_name.upper():
                        parts = value.split('-')
                        if len(parts) > 1:
                            value = parts[1].strip()

                    if exclude_pattern:
                        value = re.sub(exclude_pattern, '', value).strip()

                    if after_dash:
                        parts = value.split('-')
                        if len(parts) > 1:
                            value = parts[1].strip()

                    # Se o campo for "FORMATO", evitar duplicatas
                    if 'FORMATO' in field_name.upper() and avoid_duplicates_for_formato and value in values:
                        continue

                    values.append(value)

    return values




# Determina o tipo de SICAF com base na presença da palavra "Relatório"
def determine_sicaf_type(text):
    if 'Relatório' in text or 'RELATORIO' in text:
        return 'SICAF1'
    else:
        return 'SICAF2'


# Determina o tipo de OS com base na presença de "E-mail de Leiaute"
def determine_os_type(text):
    if 'E-mail de Leiaute' in text:
        return 'OS1'
    else:
        return 'OS2'


# Função para extrair a Razão Social de um texto de SICAF, lidando com quebras de linha
def extract_razao_social_from_sicaf(text):
    lines = text.split('\n')
    razao_social = ""
    for line in lines:
        if line.strip():  # Ignora linhas vazias
            razao_social += " " + line.strip()
    return razao_social.strip()


def normalize_razao_social(razao_social):
    """Normaliza a Razão Social para comparação."""
    if not razao_social:
        return ''

    # Remover acentos e caracteres especiais
    razao_social = unicodedata.normalize('NFKD', razao_social).encode('ASCII', 'ignore').decode('ASCII')

    # Converter para maiúsculas
    razao_social = razao_social.upper()

    # Substituir variações de "S.A." por "SOCIEDADE ANONIMA" (com espaço)
    razao_social = re.sub(r'\bS[./]?[ ]?[A]\b', 'SOCIEDADE ANONIMA', razao_social)

    # Inserir espaço após preposições se seguido de letra sem espaço
    razao_social = re.sub(r'\b(DA|DE|DO|DAS|DOS)([A-Z])', r'\1 \2', razao_social)

    # Remover pontuações e caracteres especiais restantes, mas manter espaços
    razao_social = re.sub(r'[^A-Z0-9 ]', '', razao_social)

    # Remover espaços extras
    razao_social = re.sub(r'\s+', ' ', razao_social).strip()

    # Remover todos os espaços
    razao_social = razao_social.replace(" ", "")

    return razao_social


# Função para verificar se o formato está presente no corpo de um documento AT
def check_format_in_at(at_text, formato_ap):
    if not formato_ap:
        return False

    formato_ap = re.sub(r'[“”″\'"‘’]', '"', formato_ap.strip())
    at_text_normalized = re.sub(r'[“”″\'"‘’]', '"', at_text)

    formato_ap = formato_ap.strip('"')

    pattern = re.compile(re.escape(formato_ap), re.IGNORECASE)
    lines = at_text_normalized.split('\n')
    for i, line in enumerate(lines):
        if pattern.search(line):
            return True

    return False


# Procura diretamente o valor de um formato em um PDF
def search_format_in_pdf(pdf_path, search_text):
    try:
        with pdfplumber.open(pdf_path) as pdf:
            text = "\n".join([page.extract_text() for page in pdf.pages])

        normalized_text = re.sub(r'[“”″\'"‘’]', '"', text)
        normalized_search_text = re.sub(r'[“”″\'"‘’]', '"', search_text)

        if normalized_search_text in normalized_text:
            return True
    except Exception as e:
        logging.error(f"Erro ao procurar formato no PDF {pdf_path}: {e}")

    return False


# Função para procurar uma peça diretamente no PDF
def search_peca_in_pdf(pdf_path, peca):
    try:
        with pdfplumber.open(pdf_path) as pdf:
            text = "\n".join([page.extract_text() for page in pdf.pages])

        normalized_text = re.sub(r'[“”″\'"‘’]', '"', text)
        normalized_search_text = re.sub(r'[“”″\'"‘’]', '"', peca)

        if normalized_search_text in normalized_text:
            return True
    except Exception as e:
        logging.error(f"Erro ao procurar a peça no PDF {pdf_path}: {e}")

    return False


def extract_cnpj(text):
    """
    SICAF 2

    Extrai o CNPJ do texto usando uma expressão regular.
    """
    cnpj_pattern = r'\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}'
    match = re.search(cnpj_pattern, text)
    if match:
        return match.group(0)
    return None


# Função que extrai os campos de documentos com base no tipo de documento (OS, AP, AT, SICAF)
def extract_fields(document_text, document_type):
    fields = {}
    lines = document_text.split('\n')  # Garante que 'lines' seja sempre definido
    try:
        if document_type == 'OS':
            # Extrai os campos específicos para o documento OS (variações entre OS1 e OS2)
            os_type = determine_os_type(document_text)
            fields['OS_TYPE'] = os_type
            if os_type == 'OS1':
                # Extração para OS1
                fields.update({
                    'OS N°': extract_field_value(document_text, ['OS Nº', 'OS N°'], only_numbers=True),
                    'DATA DE INICIO': extract_field_value(document_text, ['DATA DE INICIO:', 'DATA DE INÍCIO'],
                                                          below=True,
                                                          date_only=True),
                    'TITULO DA OS': extract_field_value(document_text, ['TITULO DA OS:', 'TÍTULO DA OS'], below=True,
                                                        check_next_line_if_empty=True),
                    'ORGAO': extract_field_value(document_text, ['ORGAO', 'ÓRGÃO'], below=True,
                                                 exclude_pattern=r'\d{2}/\d{2}/\d{4}', after_dash=True),
                    'TIPO DA CAMPANHA': extract_field_value(document_text, ['Nº DO PROCESSO DE SELEÇÃO INTERNA:'], below=True,
                                                            stop_before=' N° ', exclude_numbers=True )
                })
            else:
                # Extração para OS2
                fields.update({
                    'OS N°': extract_field_value(document_text, ['OS N', 'OS N°'], below=True, only_numbers=True),
                    'DATA DE INICIO': extract_field_value(document_text, ['DATA DE INÍCIO:'], below=True,
                                                          date_only=True),
                    'TITULO DA OS': extract_field_value(document_text, ['TÍTULO DA OS:'], below=True,
                                                        check_next_line_if_empty=True),
                    'ORGAO': extract_field_value(document_text, ['ÓRGÃO'], below=True,
                                                 exclude_pattern=r'\d{2}/\d{2}/\d{4}', after_dash=True),
                    'TIPO DA CAMPANHA': extract_field_value(document_text, ['TIPO DA CAMPANHA'], below=True,
                                                            exclude_numbers=True)
                })
        elif document_type == 'AP':
            # Extração de campos para o documento AP
            fields.update({
                'OS N°': extract_field_value(document_text, ['OS N°', 'OS Nº','OSNº'], stop_before='VALOR', only_numbers=True),
                'DATA EMISSAO': extract_field_value(document_text, ['DATA EMISSAO', 'DATA EMISSÃO', 'DATAEMISSÃO','DATA  EMISSÃO:'],
                                                    date_only=True),
                'CAMPANHA': extract_field_value(document_text, ['CAMPANHA:'], stop_before=['AUT.','MEIO:']),
                'PRODUTO': extract_field_value(document_text, ['PRODUTO:'], stop_before=' '),
                'AUT.CLIENTE': extract_field_value(document_text, ['AUT.CLIENTE:'], check_next_line_if_empty=True),
                'AT DE PRODUCAO': extract_field_value(document_text,
                                                      ['AT DE PRODUCAO:', 'AT DE PRODUÇÃO:', 'AT DE PRODUCAO','AT DE PRODUÇÃO','ATDEPRODUÇÃO', "AT'SDEPRODUÇÃO", "AT'S DE PRODUÇÃO"], stop_before='-', exclude_pattern=':', split_by='E'),
            })

            # Ignorar qualquer CNPJ que contenha "16.088.593"
            cnpj_value = extract_field_value(document_text, ['Cnpj: ', 'CNPJ:'])
            if "16.088.593" in cnpj_value:
                # Tentar encontrar o próximo CNPJ que não contenha "16.088.593"
                cnpj_pattern = re.compile(r'\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}')
                for line in lines:
                    cnpj_match = cnpj_pattern.search(line)
                    if cnpj_match and "16.088.593" not in cnpj_match.group():
                        cnpj_value = cnpj_match.group()
                        break

            fields['CNPJ'] = cnpj_value

            # Procurar a linha que contém o CNPJ e então pegar o Município três linhas acima e Razão Social uma linha
            # acima
            cnpj_line_index = None
            for i, line in enumerate(lines):
                if cnpj_value in line:
                    cnpj_line_index = i
                    break

            if cnpj_line_index is not None:
                # Ignorar linhas vazias e capturar Razão Social uma linha acima
                razao_social_lines = [line.strip() for line in lines[max(0, cnpj_line_index - 1):cnpj_line_index] if
                                      line.strip()]
                if razao_social_lines:
                    fields['Razão social'] = razao_social_lines[0].strip()  # Hífen preservado

                # Ignorar linhas vazias e capturar Município três linhas acima
                municipio_lines = [line.strip() for line in lines[max(0, cnpj_line_index - 3):cnpj_line_index] if
                                   line.strip()]
                fields['Município'] = municipio_lines[0].split('-')[0].strip() if municipio_lines else ""

            # Extração dos formatos e peças na linha de baixo
            pecas = extract_field_values(document_text, ['PEÇA', 'PECA'], stop_before='FORMATO', after_dash=True)
            for i, peca in enumerate(pecas):
                fields[f'PECA{i + 1}'] = peca

            formatos = extract_field_values(document_text, ['FORMATO'])
            for i, formato in enumerate(formatos):
                fields[f'FORMATO{i + 1}'] = formato.strip().strip(':').replace(' ', '')

        elif document_type == 'AT':
            # Extração de campos para o documento AT
            fields.update({
                'AT': extract_field_value(document_text, ['AT '], stop_before='DATA'),
                'TITULO': extract_field_value(document_text, ['TITULO: ', 'TÍTULO:', 'Título:'], stop_before=['Cores','CORES']),
            })
            formatos = extract_field_values(document_text, ['FORMATO:', 'Formato'])
            for i, formato in enumerate(formatos):
                fields[f'FORMATO{i + 1}'] = formato
            fields.update({
                'Data da AT': extract_field_value(document_text, ['Data:', 'DATA:'], date_only=True)
            })
        elif document_type == 'SICAF':
            # Extração de campos para o documento SICAF (variação entre SICAF1 e SICAF2)
            sicaf_type = determine_sicaf_type(document_text)
            fields['SICAF_TYPE'] = sicaf_type
            if sicaf_type == 'SICAF1':
                # Extração para SICAF1
                fields.update({
                    'Razão social': extract_field_value(document_text, ['Razao Social:','Razão Social:']),
                    'CNPJ': extract_field_value(document_text, ['CNPJ: ', 'CNPJ:'], stop_before='Data'),
                    'Município': extract_field_value(document_text, ['Municipio: ', 'Munícipio:'], stop_before=' N°')
                })
            else:
                # Extração para SICAF2
                fields.update({
                    'CNPJ': extract_cnpj(document_text)
                })
    except Exception as e:
        logging.error(f"Erro ao extrair campos para {document_type}: {e}")
    return fields


# Verifica se a Razão Social do SICAF está presente no AP nas linhas 21 ou 22
def check_razao_social_in_ap(ap_text, razao_social):
    if not razao_social:
        return False
    lines = ap_text.split('\n')
    if len(lines) >= 22:
        if razao_social in lines[20] or razao_social in lines[21]:
            return True
    return False


# Verifica se o Município do SICAF está presente no AP entre as linhas 18 e 23
def check_municipio_in_ap(ap_text, municipio):
    if not municipio:
        return False
    lines = ap_text.split('\n')
    for i in range(17, 23):  # Os números das linhas são indexados a partir de zero
        if i < len(lines) and municipio in lines[i]:
            return True
    return False


# Verifica se a peça está presente no documento AT
def check_peca_in_at(at_text, peca):
    if not peca:
        return False

    peca_normalized = re.sub(r'[“”″\'"‘’]', '"', peca.strip())
    at_text_normalized = re.sub(r'[“”″\'"‘’]', '"', at_text)

    peca_normalized = re.sub(r'\s+', ' ', peca_normalized)
    at_text_normalized = re.sub(r'\s+', ' ', at_text_normalized)

    if peca_normalized in at_text_normalized:
        return True

    fragments = peca_normalized.split()
    for fragment in fragments:
        if fragment not in at_text_normalized:
            return False

    return True


def determine_overall_status(field_statuses, required_pieces, found_pieces):
    """
    Determina o status geral do processo.
    1. Verifica se todos os campos estão OK.
    2. Se todos os campos estiverem OK, verifica se todas as peças encontradas correspondem às do AP.
    Retorna uma tupla com o status geral e a classe CSS correspondente.
    """
    # Log para depuração
    print("Field statuses:", field_statuses)
    print("Required pieces:", required_pieces)
    print("Found pieces:", found_pieces)

    # Verificar se TODOS os campos estão 'OK', normalizando os valores
    all_fields_ok = all(str(status).strip().upper() == 'OK' for status in field_statuses.values() if status is not None)
    print(f"All fields OK: {all_fields_ok}")

    # Se todos os campos estiverem OK, então verifica as peças
    if all_fields_ok:
        # Verificar se todas as peças do AP estão no AT e vice-versa
        all_pieces_match = set(required_pieces) == set(found_pieces)
        print(f"All pieces match: {all_pieces_match}")

        # Se todas as peças estiverem corretas, status será 'OK'
        if all_pieces_match:
            overall_status = 'OK'
            status_class = 'status-ok'
        else:
            overall_status = 'NC'
            status_class = 'status-nc'
    else:
        # Se algum campo estiver incorreto, status será 'NC'
        overall_status = 'NC'
        status_class = 'status-nc'

    print(f"Overall status: {overall_status}")
    return overall_status, status_class



def generate_html_report(report, subfolder_name, overall_status, status_class):
    """
    Gera o relatório HTML com base no relatório de texto, nome da subpasta,
    status geral e classe CSS correspondente.
    """
    # Inicia a div do processo com a classe de status
    html_report = f"""
    <div class="report {status_class}">
        <h2>{escape(subfolder_name)}</h2>
        <div class="document-sections-container">
    """

    lines = report.split('\n')
    index = 0
    while index < len(lines):
        line = lines[index].strip()
        if line.startswith('- OS:'):
            # Extrai os dados da OS
            os_data_str = line[len('- OS:'):].strip()
            try:
                os_data = ast.literal_eval(os_data_str)
                # Cria uma div para os dados da OS
                html_report += '<div class="document-section">'
                html_report += '<h3>OS</h3>'
                html_report += '<ul>'
                for key, value in os_data.items():
                    if isinstance(value, list):
                        value_str = ', '.join([escape(str(v)) for v in value])
                        html_report += f'<li><strong>{escape(key)}:</strong> {value_str}</li>'
                    else:
                        html_report += f'<li><strong>{escape(key)}:</strong> {escape(str(value))}</li>'
                html_report += '</ul>'
                html_report += '</div>'
            except Exception as e:
                html_report += f'<p>Erro ao processar dados da OS: {e}</p>'
        elif line.startswith('- AP:'):
            # Extrai os dados da AP
            ap_data_str = line[len('- AP:'):].strip()
            try:
                ap_data = ast.literal_eval(ap_data_str)
                # Cria uma div para os dados da AP
                html_report += '<div class="document-section">'
                html_report += '<h3>AP</h3>'
                html_report += '<ul>'
                for key, value in ap_data.items():
                    if isinstance(value, list):
                        value_str = ', '.join([escape(str(v)) for v in value])
                        html_report += f'<li><strong>{escape(key)}:</strong> {value_str}</li>'
                    else:
                        html_report += f'<li><strong>{escape(key)}:</strong> {escape(str(value))}</li>'
                html_report += '</ul>'
                html_report += '</div>'
            except Exception as e:
                html_report += f'<p>Erro ao processar dados da AP: {e}</p>'
        elif line.startswith('- AT ('):
            # Extrai os dados do AT
            at_match = re.match(r'- AT \((.*?)\): (.*)', line)
            if at_match:
                at_file_name = at_match.group(1)
                at_data_str = at_match.group(2).strip()
                try:
                    at_data = ast.literal_eval(at_data_str)
                    # Cria uma div para os dados do AT
                    html_report += '<div class="document-section">'
                    html_report += f'<h3>AT ({escape(at_file_name)})</h3>'
                    html_report += '<ul>'
                    for key, value in at_data.items():
                        if key != 'FILE_NAME':
                            if isinstance(value, list):
                                value_str = ', '.join([escape(str(v)) for v in value])
                                html_report += f'<li><strong>{escape(key)}:</strong> {value_str}</li>'
                            else:
                                html_report += f'<li><strong>{escape(key)}:</strong> {escape(str(value))}</li>'
                    html_report += '</ul>'
                    html_report += '</div>'
                except Exception as e:
                    html_report += f'<p>Erro ao processar dados do AT: {e}</p>'
            else:
                html_report += f"<p>{escape(line)}</p>"
        elif line.startswith('- AT: Nenhum arquivo AT encontrado.'):
            html_report += f"<p>{escape(line)}</p>"
        elif line.startswith('- SICAF:'):
            # Extrai os dados do SICAF
            sicaf_data_str = line[len('- SICAF:'):].strip()
            try:
                sicaf_data = ast.literal_eval(sicaf_data_str)
                # Cria uma div para os dados do SICAF
                html_report += '<div class="document-section">'
                html_report += '<h3>SICAF</h3>'
                html_report += '<ul>'
                for key, value in sicaf_data.items():
                    if isinstance(value, list):
                        value_str = ', '.join([escape(str(v)) for v in value])
                        html_report += f'<li><strong>{escape(key)}:</strong> {value_str}</li>'
                    else:
                        html_report += f'<li><strong>{escape(key)}:</strong> {escape(str(value))}</li>'
                html_report += '</ul>'
                html_report += '</div>'
            except Exception as e:
                html_report += f'<p>Erro ao processar dados do SICAF: {e}</p>'
        index += 1

    # Fecha o container das seções dos documentos
    html_report += '</div>'

    # Processa os checks (verificações)
    for line in lines:
        line = line.strip()
        if 'CHECK' in line:
            index = line.find('CHECK')
            field = escape(line[:index].strip())
            value = escape(line[index:].strip())

            # Determina a classe com base no resultado
            if 'OK' in value:
                result_class = 'ok'
            elif 'Non-conformity' in value:
                result_class = 'non-conformity'
            else:
                result_class = ''

            html_report += f"""
            <div class="check-row">
                <div class="field">{field}</div>
                <div class="value {result_class}">{value}</div>
            </div>
            """
        else:
            continue  # Ignora outras linhas que não são checks

    # Adiciona o status geral do processo ao relatório
    html_report += f"""
    <div class="overall-status {status_class}">
        <strong>Status do Processo:</strong> {overall_status}
    </div>
    """
    html_report += "</div>"  # Fecha a div class="report"

    return html_report



def save_and_open_report(html_report):
    """Retorna o relatório HTML para ser renderizado."""
    return html_report


    try:
        os.startfile(report_path)
    except Exception as e:
        logging.error(f"Erro ao abrir o relatório: {e}")
        messagebox.showerror("Erro", f"Não foi possível abrir o relatório: {e}")


# Salva o texto extraído do AP em um arquivo de texto
def save_ap_text(ap_text):
    with open("../ap_text.txt", "w", encoding="utf-8") as file:
        file.write(ap_text)


# Procura o texto diretamente no PDF e retorna o texto concatenado das ocorrências
def search_text_in_pdf(pdf_path, search_text):
    concatenated_result = ""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            text = "\n".join([page.extract_text() for page in pdf.pages])
        normalized_text = re.sub(r'\s+', '', text.lower())
        normalized_search_text = re.sub(r'\s+', '', search_text.lower())
        if normalized_search_text in normalized_text:
            concatenated_result = search_text
    except Exception as e:
        logging.error(f"Erro ao procurar texto no PDF {pdf_path}: {e}")
    return concatenated_result.strip()


# Função que realiza várias verificações nos campos extraídos de OS, AP, AT e SICAF, gerando um relatório detalhado
def check_fields(os_fields, ap_fields, at_fields_list, sicaf_fields, sicaf_pdf_path, at_files, missing_at_numbers, found_pieces):
    report = ["Transcrição das informações extraídas:", f"- OS: {os_fields}", f"- AP: {ap_fields}"]

    if not os_fields or not ap_fields or not sicaf_fields:
        error_message = "Erro ao extrair campos dos documentos."
        logging.error(error_message)
        return "", error_message

    # Adiciona informações dos ATs se houver
    if at_fields_list:
        for at_fields in at_fields_list:
            at_file_name = at_fields.get('FILE_NAME', f"- AT {at_fields_list.index(at_fields) + 1}")
            report.append(f"- AT ({at_file_name}): {at_fields}")
    else:
        report.append("- AT: Nenhum arquivo AT encontrado.")

    report.append(f"- SICAF: {sicaf_fields}")

    # Função para normalizar formatos (usada tanto na extração quanto na checagem)
    def normalize_format(value):
        return value.upper()

    # CHECK 1 - Verificações comuns
    try:
        report.append(
            f"OS N°                                   CHECK 1.1: {'OK - OS N° ' + os_fields.get('OS N°') if os_fields.get('OS N°') == ap_fields.get('OS N°') else 'Non-conformity - OS N° ' + ap_fields.get('OS N°')}")
        os_data_inicio = os_fields.get('DATA DE INICIO')
        ap_data_emissao = ap_fields.get('DATA EMISSAO')
        if os_data_inicio and ap_data_emissao:
            os_data_inicio_dt = datetime.strptime(os_data_inicio, '%d/%m/%Y')
            ap_data_emissao_dt = datetime.strptime(ap_data_emissao, '%d/%m/%Y')
            report.append(
                f"DATAS                                   CHECK 1.2: {'OK' if ap_data_emissao_dt > os_data_inicio_dt else 'Non-conformity'}")
        else:
            report.append("DATAS                                   CHECK 1.2: Non-conformity")
    except Exception as e:
        report.append(f"DATAS                                   CHECK 1.2: Error {e}")

    report.append(
        f"TITULO DA OS/CAMPANHA                   CHECK 1.3: {'OK' if os_fields.get('TITULO DA OS') == ap_fields.get('CAMPANHA') else 'Non-conformity'}")

    orgao = os_fields.get('ORGAO', '')
    produto = ap_fields.get('PRODUTO', '')
    if orgao and produto:
        report.append(
            f"ORGAO/PRODUTO                           CHECK 1.4: {'OK' if orgao == produto else 'Non-conformity'}")
    else:
        report.append("CHECK 1.4: Non-conformity")

    report.append(
        f"TIPO DA CAMPANHA/AUT.CLIENTE            CHECK 1.5: {'OK' if os_fields.get('TIPO DA CAMPANHA') == ap_fields.get('AUT.CLIENTE') else 'Non-conformity'}")


    # CHECK 2 - Verificação de ATs

    for i, at_file in enumerate(at_files):
        at_file_name = os.path.basename(at_file)  # Nome do arquivo AT
        at_number_match = re.search(r'AT\s*(\d+)', at_file_name)  # Encontra a numeração do AT no nome do arquivo

        if at_number_match:
            at_number = at_number_match.group(1)  # Numeração do AT extraída do nome do arquivo
        else:
            # Se o número do AT não for encontrado, continue para o próximo AT
            continue


        # Verifica se o número do AT atual existe na lista de ATs do AP
        ap_value_list = ap_fields.get('AT DE PRODUCAO', [])

        # Garante que ap_value_list seja uma lista, mesmo se o campo for None
        if ap_value_list is None:
            ap_value_list = []

        # Se ap_value_list for uma string, converta-a para uma lista com um único elemento
        if isinstance(ap_value_list, str):
            ap_value_list = [ap_value_list.strip()]
        else:
            ap_value_list = [str(item).strip() for item in ap_value_list]

        # Se o número do AT estiver presente no AP, ele será processado
        if at_number in ap_value_list:
            report.append(f"AT {at_number} - ({at_file_name}) /AT DE PRODUCAO CHECK 2.{i + 1}.1: OK")

            # Verificação de Peças no AT
            for peca_key in [f'PECA{j + 1}' for j in range(len(ap_fields))]:
                peca = ap_fields.get(peca_key)
                if peca:
                    matched = False
                    peca_no_spaces = re.sub(r'\s+', '', peca)  # Remove espaços da peça

                    with pdfplumber.open(at_file) as pdf:
                        # Remove espaços do texto completo do PDF do AT
                        at_text_no_spaces = re.sub(r'\s+', '', "".join([page.extract_text() for page in pdf.pages if page.extract_text()]))

                        # Verifica se a peça sem espaços está presente no texto do PDF
                        if peca_no_spaces in at_text_no_spaces:
                            matched = True
                            report.append(
                                f"AT {at_number} - ({at_file_name}) PECA/PDF AT CHECK 2.{i + 1}.2: OK - {peca_key} ({peca}) encontrada no {at_file_name} (sem espaços)")
                            found_pieces.append(peca.strip().upper())  # Adiciona a peça encontrada
                        else:
                            report.append(
                                f"AT {at_number} - ({at_file_name}) PECA/PDF AT CHECK 2.{i + 1}.2: Non-conformity - {peca_key} ({peca}) não encontrada no PDF (mesmo sem espaços)")
        else:
            # Não adiciona o AT ao relatório se ele não estiver presente no AP
            continue  # Continua para o próximo arquivo AT sem exibir nada

        # Normalizando formatos tanto do AT quanto do AP antes da checagem
        formatos_ap = [ap_fields.get(f'FORMATO{j + 1}') for j in range(len(ap_fields)) if
                       ap_fields.get(f'FORMATO{j + 1}')]
        formatos_at = [at_fields.get(f'FORMATO{j + 1}') for j in range(len(at_fields)) if
                       at_fields.get(f'FORMATO{j + 1}')]

        # Remover formatos duplicados do AP
        formatos_ap_unicos = list(set(formatos_ap))  # Remove duplicatas

        for formato_ap in formatos_ap_unicos:
            matched = False  # Inicializa a variável matched como False

            if formato_ap:
                # Verifica se o formato está presente no texto extraído do OCR
                if formato_ap in formatos_at:
                    matched = True

                # Se não encontrou no texto do OCR, verifica diretamente no PDF
                elif not matched:
                    if search_format_in_pdf(at_files[i], formato_ap):
                        matched = True

            # Adiciona a verificação ao relatório com base no resultado da verificação
            report.append(
                f"AT - ({at_file_name}) FORMATO/FORMATO                         CHECK 2.{i + 1}.3: {'OK' if matched else 'Non-conformity'} - Formato {formato_ap}")

        try:
            ap_data_emissao_dt = datetime.strptime(ap_fields.get('DATA EMISSAO', '01/01/1900'), '%d/%m/%Y')
            at_data_at_dt = datetime.strptime(at_fields.get('Data da AT', '01/01/1900'), '%d/%m/%Y')
            report.append(
                f"DATA EMISSAO/Data da AT                 CHECK 2.{i + 1}.4: {'OK' if ap_data_emissao_dt >= at_data_at_dt else 'Non-conformity'}")
        except Exception as e:
            report.append(f"DATA EMISSAO/Data da AT                 CHECK 2.{i + 1}.4: Error {e}")

    # CHECK 3 - SICAF verificações
    if sicaf_fields.get('SICAF_TYPE') == 'SICAF1':
        razao_social_sicaf = normalize_razao_social(sicaf_fields.get('Razão social', ''))
        razao_social_ap = normalize_razao_social(ap_fields.get('Razão social', ''))

        report.append(f"Razão social                            CHECK 3.1: {'OK' if razao_social_sicaf == razao_social_ap else 'Non-conformity'}")
        report.append(f"CNPJ                                    CHECK 3.2: {'OK' if sicaf_fields.get('CNPJ') == ap_fields.get('CNPJ') else 'Non-conformity'}")
        report.append(f"Município                               CHECK 3.3: {'OK' if sicaf_fields.get('Município') == ap_fields.get('Município') else 'Non-conformity'}")
    else:
        # SICAF 2 - Pesquisa direta no PDF
        razao_social_ap = ap_fields.get('Razão social')
        if razao_social_ap:
            razao_social_ap_no_spaces = razao_social_ap.replace(" ", "").upper()

            try:
                sicaf_text = extract_text_with_pdfminer_layout(sicaf_pdf_path)
                sicaf_text_no_spaces = sicaf_text.replace(" ", "").upper()

                if razao_social_ap_no_spaces in sicaf_text_no_spaces:
                    report.append("Razão social                            CHECK 3.1: OK - Razão Social do AP encontrada no SICAF (sem espaços).")
                else:
                    report.append("Razão social                            CHECK 3.1: Non-conformity - Razão Social do AP não encontrada no SICAF.")
            except Exception as e:
                logging.error(f"Erro ao procurar razão social no SICAF 2: {e}")
                report.append("Razão social                            CHECK 3.1: Error - Não foi possível verificar a Razão Social no SICAF.")
        else:
            report.append("Razão social                            CHECK 3.1: Non-conformity - Razão Social do AP não foi encontrada.")

        cnpj_ap = ap_fields.get('CNPJ')
        if cnpj_ap and search_text_in_pdf(sicaf_pdf_path, cnpj_ap):
            report.append("CNPJ                                    CHECK 3.2: OK - CNPJ do AP encontrado no SICAF.")
        else:
            report.append("CNPJ                                    CHECK 3.2: Non-conformity - CNPJ do AP não encontrado no SICAF.")

        municipio_ap = ap_fields.get('Município')
        if municipio_ap and search_text_in_pdf(sicaf_pdf_path, municipio_ap):
            report.append("Município                               CHECK 3.3: OK - Município do AP encontrado no SICAF.")
        else:
            report.append("Município                               CHECK 3.3: Non-conformity - Município do AP não encontrado no SICAF.")

    report_text = "\n".join(report)
    return report_text, None  # Retorna o relatório e None indicando que não houve erro


def save_text_to_file(text, file_name, folder_path):
    """Salva o texto extraído em um arquivo de texto."""
    if not folder_path:
        logging.error("folder_path is None")
        return
    file_path = os.path.join(folder_path, file_name)
    with open(file_path, "w", encoding="utf-8") as file:
        file.write(text)
    logging.info(f"Texto salvo em {file_path}")


# Função que realiza a verificação dos documentos, extraindo campos e gerando um relatório
def verify_documents(file_paths, subfolder_name, temp_pdf_dir, move_os_at_files=True):
    os_file = file_paths.get('OS')
    ap_file = file_paths.get('AP')
    at_files = file_paths.get('AT', [])  # Pode ser uma lista vazia se não houver AT
    sicaf_file = file_paths.get('SICAF')

    if not all([os_file, ap_file, sicaf_file]):
        error_message = f"Todos os arquivos (OS, AP, SICAF) devem estar presentes na pasta {subfolder_name}."
        logging.error(error_message)
        return "", None, error_message

    folder_path = os.path.dirname(__file__)

    # Extrai e salva o texto do documento OS
    os_text = extract_text_with_format_adjustment(os_file)
    determine_os_type(os_text)
    os_fields = extract_fields(os_text, 'OS')
    save_text_to_file(os_text, f"os_text_{subfolder_name}.txt", temp_pdf_dir)

    # Extrai e salva o texto do documento AP
    ap_text = extract_text_with_format_adjustment_py(ap_file)
    ap_fields = extract_fields(ap_text, 'AP')
    save_text_to_file(ap_text, f"ap_text_{subfolder_name}.txt", temp_pdf_dir)

    # Processa e salva o texto de cada AT
    at_fields_list = []
    at_de_producao = ap_fields.get('AT DE PRODUCAO')
    if isinstance(at_de_producao, list):
        at_numbers_in_ap = [num.strip() for num in at_de_producao]
    elif at_de_producao:
        at_numbers_in_ap = [at_de_producao.strip()]
    else:
        at_numbers_in_ap = []

    at_numbers_found = []
    at_number_to_file = {}  # Dicionário para mapear números de AT para arquivos AT correspondentes

    for at_file in at_files:
        at_text = extract_text_with_format_adjustment(at_file)
        at_fields = extract_fields(at_text, 'AT')
        at_fields['FILE_NAME'] = os.path.basename(at_file)
        at_number = at_fields.get('AT').strip()
        if at_number in at_numbers_in_ap:
            at_fields_list.append(at_fields)
            at_numbers_found.append(at_number)
            at_number_to_file[at_number] = at_file  # Mapeia o número de AT para o arquivo AT
            # Salva o texto extraído do AT
            save_text_to_file(at_text, f"at_text_{os.path.basename(at_file)}.txt", temp_pdf_dir)
        else:
            # AT número não está na AP, não processar
            pass

    # Encontrar ATs que estão na AP mas não possuem arquivo correspondente
    missing_at_numbers = set(at_numbers_in_ap) - set(at_numbers_found)

    # Extrai e salva o texto do SICAF
    sicaf_text = extract_text_with_format_adjustment(sicaf_file)

    sicaf_type = determine_sicaf_type(sicaf_text)
    logging.info(f"Tipo de SICAF identificado: {sicaf_type}")

    if sicaf_type == 'SICAF2':
        logging.info("Usando pdfminer para SICAF 2 mantendo layout")
        sicaf_text = extract_text_with_pdfminer_layout(sicaf_file)
    # Salva o texto extraído do SICAF
    save_text_to_file(sicaf_text, f"sicaf_text_{subfolder_name}.txt", temp_pdf_dir)

    sicaf_fields = extract_fields(sicaf_text, 'SICAF')

    # Inicializa a lista para armazenar as peças encontradas
    found_pieces = []

    # Gera o relatório
    report, error_message = check_fields(os_fields, ap_fields, at_fields_list, sicaf_fields, sicaf_file, at_files, missing_at_numbers, found_pieces)
    if error_message:
        return "", None, error_message

    # Determina se todas as peças do AP foram encontradas em algum AT
    required_pieces = [value.strip().upper() for key, value in ap_fields.items() if key.startswith('PECA')]

    all_pieces_found = all(piece in found_pieces for piece in required_pieces)

    # Log para depuração
    logging.info(f"Required pieces: {required_pieces}")
    logging.info(f"Pieces found: {found_pieces}")
    logging.info(f"All pieces found: {all_pieces_found}")

    # Criar o dicionário de field_statuses com base no relatório
    field_statuses = {
        'OS N°': 'OK' if 'CHECK 1.1' in report and 'OK' in report else 'NC',
        'DATAS': 'OK' if 'CHECK 1.2' in report and 'OK' in report else 'NC',
        'TITULO DA OS/CAMPANHA': 'OK' if 'CHECK 1.3' in report and 'OK' in report else 'NC',
        'ORGAO/PRODUTO': 'OK' if 'CHECK 1.4' in report and 'OK' in report else 'NC',
        'TIPO DA CAMPANHA/AUT.CLIENTE': 'OK' if 'CHECK 1.5' in report and 'OK' in report else 'NC',

        # Para AT / AT DE PRODUCAO, verifica todos os checks possíveis (2.1.1, 2.2.1, 2.3.1, etc.)
        'AT /AT DE PRODUCAO': 'OK' if any(
            f'CHECK 2.{i}.1' in report and 'OK' in report for i in range(1, 10)) else 'NC',

        # Para AT FORMATO/FORMATO, verifica todos os checks possíveis (2.1.3, 2.2.3, 2.3.3, etc.)
        'AT FORMATO/FORMATO': 'OK' if any(
            f'CHECK 2.{i}.3' in report and 'OK' in report for i in range(1, 10)) else 'NC',

        # Para DATA EMISSAO/Data da AT, verifica todos os checks possíveis (2.1.4, 2.2.4, 2.3.4, etc.)
        'DATA EMISSAO/Data da AT': 'OK' if any(
            f'CHECK 2.{i}.4' in report and 'OK' in report for i in range(1, 10)) else 'NC',

        'Razão social': 'OK' if 'CHECK 3.1' in report and 'OK' in report else 'NC',
        'CNPJ': 'OK' if 'CHECK 3.2' in report and 'OK' in report else 'NC',
        'Município': 'OK' if 'CHECK 3.3' in report and 'OK' in report else 'NC'
    }

    # Determina o status geral com base nas verificações e nas peças encontradas
    overall_status, status_class = determine_overall_status(
        field_statuses,  # status de cada campo
        required_pieces,  # Peças que são esperadas a partir do AP
        found_pieces  # Peças que foram realmente encontradas nos ATs
    )

    # Define a pasta "Relatorios" onde as pastas "OK" e "Non-conformity" serão armazenadas
    relatorios_folder = os.path.join(folder_path, "Relatorios")
    os.makedirs(relatorios_folder, exist_ok=True)

    # Determina a pasta de destino com base no status geral
    if overall_status == 'OK':
        status_folder = os.path.join(relatorios_folder, "OK", subfolder_name)
    else:
        status_folder = os.path.join(relatorios_folder, "Non-conformity", subfolder_name)

    # Cria a pasta de destino se ela não existir
    os.makedirs(status_folder, exist_ok=True)

    # Move cada arquivo para a pasta de destino correspondente
    try:
        # Sempre mover os arquivos AP e SICAF
        shutil.move(ap_file, os.path.join(status_folder, os.path.basename(ap_file)))
        shutil.move(sicaf_file, os.path.join(status_folder, os.path.basename(sicaf_file)))

        if overall_status == 'Non-conformity' and move_os_at_files:
            # Na pasta 'Non-conformity', mover todos os arquivos
            shutil.move(os_file, os.path.join(status_folder, os.path.basename(os_file)))

            # Mover arquivos AT, se houver
            for at_file in at_files:
                shutil.move(at_file, os.path.join(status_folder, os.path.basename(at_file)))

        logging.info(f"Arquivos movidos para a pasta: {status_folder}")
    except Exception as e:
        logging.error(f"Erro ao mover os arquivos: {e}")

    # Gerar o relatório HTML
    html_report = generate_html_report(report, subfolder_name, overall_status, status_class)

    return html_report, overall_status, None  # Retorna o relatório, o status e None indicando que não houve erro



def move_relatorios_folder(destination_path):
    """Move a pasta 'Relatorios' para o caminho de destino especificado, incluindo data e hora no nome."""
    try:
        folder_path = os.path.dirname(__file__)
        relatorios_folder = os.path.join(folder_path, "Relatorios")
        if os.path.exists(relatorios_folder):
            # Obtém a data e hora atual no formato desejado
            now = datetime.now()
            timestamp = now.strftime("%Y%m%d_%H%M%S")
            # Novo nome da pasta incluindo data e hora
            new_folder_name = f"Relatorios_{timestamp}"
            # Define o novo caminho para a pasta 'Relatorios'
            destination_folder = os.path.join(destination_path, new_folder_name)
            # Verifica se o destino existe, caso contrário, cria o diretório
            if not os.path.exists(destination_path):
                os.makedirs(destination_path)
            # Move a pasta 'Relatorios' para o destino com o novo nome
            shutil.move(relatorios_folder, destination_folder)
            logging.info(f"Pasta 'Relatorios' movida para {destination_folder}")
            # Não recriar a pasta 'Relatorios' após movê-la
        else:
            logging.warning("A pasta 'Relatorios' não existe e não pode ser movida.")
    except Exception as e:
        logging.error(f"Erro ao mover a pasta 'Relatorios': {e}")



def delete_temp_folder():
    # Aguarda 45 segundos
    time.sleep(45)
    temp_folder = os.path.join(os.path.dirname(__file__), 'temp_pdf')
    try:
        # Remove a pasta temporária e todo o conteúdo
        shutil.rmtree(temp_folder)
        logging.info(f"Pasta temporária {temp_folder} apagada com sucesso.")
    except Exception as e:
        logging.error(f"Erro ao tentar apagar a pasta temporária {temp_folder}: {e}")

@app.route('/', methods=['GET', 'POST'])
def upload_files():
    if request.method == 'POST':
        # Verifica se há arquivos enviados
        if 'files' not in request.files:
            error_message = "Faltando arquivos do diretório."
            return render_template('error.html', error_message=error_message), 400

        # Define o caminho da pasta temp_pdf
        temp_pdf_dir = os.path.join(os.path.dirname(__file__), 'temp_pdf')

        # Verifica se a pasta temp_pdf existe e, se existir, apaga diretamente
        if os.path.exists(temp_pdf_dir):
            try:
                shutil.rmtree(temp_pdf_dir)
                logging.info(f"Pasta temporária {temp_pdf_dir} apagada no início da execução.")
            except Exception as e:
                logging.error(f"Erro ao tentar apagar a pasta temporária {temp_pdf_dir}: {e}")

        # Cria a pasta temporária "temp_pdf"
        os.makedirs(temp_pdf_dir, exist_ok=True)

        # Inicia o timer para apagar a pasta temp_pdf após 45 segundos
        threading.Thread(target=delete_temp_folder).start()

        # Salva cada arquivo na pasta temp_pdf
        files = request.files.getlist('files')
        for file in files:
            file_path = os.path.join(temp_pdf_dir, file.filename)
            os.makedirs(os.path.dirname(file_path), exist_ok=True)
            file.save(file_path)

        full_html_report = ""
        plano_dirs = []

        # Variáveis para o resumo
        total_subfolders_sent = 0
        total_subfolders_processed = 0
        total_subfolders_ignored = 0
        ignored_subfolders = []
        root_folder_name = ""

        # Inicializa as listas para os processos OK e NC
        ok_processes = []
        nc_processes = []

        # Obtém a lista de subpastas imediatas dentro de temp_pdf
        immediate_subdirs = [d for d in os.listdir(temp_pdf_dir) if os.path.isdir(os.path.join(temp_pdf_dir, d))]

        # Define o nome da pasta raiz como a primeira subpasta dentro de temp_pdf
        if immediate_subdirs:
            root_folder_name = immediate_subdirs[0]
        else:
            root_folder_name = ""

        # Percorre todas as subpastas na pasta temp_pdf
        for root, dirs, files in os.walk(temp_pdf_dir):
            # Ignora a própria pasta temp_pdf
            if root == temp_pdf_dir:
                continue

            # Verifica se 'plano' está no nome da pasta
            if 'plano' in os.path.basename(root).lower():
                plano_dir = root
                plano_dirs.append(plano_dir)

                # Coleta os arquivos AT e OS na pasta 'plano'
                at_files = []
                os_file = None

                for file_name in os.listdir(plano_dir):
                    file_path = os.path.join(plano_dir, file_name)
                    if os.path.isfile(file_path):
                        if 'AT' in file_name.upper():
                            at_files.append(file_path)
                        elif 'OS' in file_name.upper():
                            os_file = file_path

                # Percorre as subpastas dentro da pasta 'plano'
                subdirs = [d for d in os.listdir(plano_dir) if os.path.isdir(os.path.join(plano_dir, d))]
                total_subfolders_sent += len(subdirs)  # Atualiza o total de subpastas enviadas

                for subdir in subdirs:
                    subdir_path = os.path.join(plano_dir, subdir)
                    # Coleta os arquivos SICAF e AP na subpasta
                    sicaf_file = None
                    ap_file = None

                    for file_name in os.listdir(subdir_path):
                        if 'SICAF' in file_name.upper():
                            sicaf_file = os.path.join(subdir_path, file_name)
                        elif 'AP' in file_name.upper():
                            ap_file = os.path.join(subdir_path, file_name)

                    # Monta o dicionário de arquivos
                    file_paths = {
                        'OS': os_file,
                        'AT': at_files,
                        'SICAF': sicaf_file,
                        'AP': ap_file  # Pode ser None
                    }

                    subfolder_name = subdir

                    # Chama a função de verificação
                    result, status, error_message = verify_documents(file_paths, subfolder_name, temp_pdf_dir)

                    if error_message:
                        # Não exibe a mensagem de erro se 'plano' estiver no nome da pasta
                        pass
                    else:
                        full_html_report += result
                        total_subfolders_processed += 1  # Atualiza o total de subpastas processadas

                        # Adiciona o nome da subpasta à lista correspondente
                        if status == 'OK':
                            ok_processes.append(subfolder_name)
                        elif status == 'NC':
                            nc_processes.append(subfolder_name)

                # Opcionalmente, mover os arquivos AT e OS após processar todas as subpastas
                # Neste exemplo, não movemos os arquivos AT e OS para evitar conflitos
                continue  # Continua para a próxima iteração, evitando processamento adicional

            # Evita processar subpastas das pastas 'plano' novamente
            elif root != temp_pdf_dir and not any(root.startswith(plano_dir + os.sep) for plano_dir in plano_dirs):
                file_paths = {'AT': []}
                for file_name in files:
                    # Verifica se os arquivos possuem "OS", "AP", "AT" ou "SICAF" no nome e armazena os caminhos
                    if 'OS' in file_name.upper():
                        file_paths['OS'] = os.path.join(root, file_name)
                    elif 'AP' in file_name.upper():
                        file_paths['AP'] = os.path.join(root, file_name)
                    elif 'SICAF' in file_name.upper():
                        file_paths['SICAF'] = os.path.join(root, file_name)
                    elif 'AT' in file_name.upper():
                        file_paths['AT'].append(os.path.join(root, file_name))

                subfolder_name = os.path.basename(root)
                total_subfolders_sent += 1  # Atualiza o total de subpastas enviadas

                # Verifica se a subpasta tem um AP válido
                if 'AP' not in file_paths or not file_paths['AP']:
                    total_subfolders_ignored += 1
                    ignored_subfolders.append(subfolder_name)
                    continue  # Ignora esta subpasta

                result, status, error_message = verify_documents(file_paths, subfolder_name, temp_pdf_dir)

                if error_message:
                    # Exibe a mensagem de erro apenas quando a pasta não contém 'plano' no nome
                    full_html_report += f"<h2>Erro no conjunto {subfolder_name}</h2><p>{error_message}</p>"
                else:
                    full_html_report += result
                    total_subfolders_processed += 1  # Atualiza o total de subpastas processadas

                    # Adiciona o nome da subpasta à lista correspondente
                    if status == 'OK':
                        ok_processes.append(subfolder_name)
                    elif status == 'NC':
                        nc_processes.append(subfolder_name)

        # Cria o relatório resumido
        summary_report = f"""
        <div class="summary">
            <h2>Resumo do Processamento</h2>
            <p>Total de subpastas enviadas: <strong>{total_subfolders_sent}</strong></p>
            <p>Total de subpastas processadas: <strong>{total_subfolders_processed}</strong></p>
            <p>Total de subpastas ignoradas (sem AP válido): <strong>{total_subfolders_ignored}</strong></p>
        """

        # Lista os nomes das subpastas ignoradas, se houver
        if total_subfolders_ignored > 0:
            summary_report += "<p>Subpastas Ignoradas:</p><ul>"
            for ignored in ignored_subfolders:
                summary_report += f"<li>{escape(ignored)}</li>"
            summary_report += "</ul>"

        # Adiciona o nome da pasta raiz
        if root_folder_name:
            summary_report += f"<p>Pasta Raiz: <strong>{escape(root_folder_name)}</strong></p>"

        # Processos OK e NC - report
        # Adiciona as listas de processos OK e NC
        if ok_processes:
            summary_report += "<p>Processos OK:</p><ul>"
            for process in ok_processes:
                summary_report += f"<li class='ok'><span class='icon'></span>{escape(process)}</li>"
            summary_report += "</ul>"

        if nc_processes:
            summary_report += "<p>Processos NC:</p><ul>"
            for process in nc_processes:
                summary_report += f"<li class='nc'><span class='icon'></span>{escape(process)}</li>"
            summary_report += "</ul>"

        summary_report += "</div>"

        # Adiciona o resumo ao relatório completo
        full_html_report += summary_report

        # Move a pasta 'Relatorios' para o caminho especificado
        destination_path = r"G:\Shared drives\AUTOMACAO\CHECKIN_MIDIA"  # Atualize conforme necessário
        move_relatorios_folder(destination_path)

        # Renderiza o relatório HTML
        return render_template('report.html', report_content=full_html_report)

    return render_template('upload.html')


if __name__ == '__main__':
    # Obtém o endereço IP local da máquina automaticamente
    host_ip = socket.gethostbyname(socket.gethostname())
    app.run(host=host_ip, debug=True, port='80')
