# services.py

import os
import sys
import re
import logging
import unicodedata
import shutil
import ast
import threading
import time
import tempfile
from datetime import datetime
import pdfplumber
from PyPDF2 import PdfReader
from pdfminer.high_level import extract_text_to_fp
from io import StringIO
from html import escape

# Ajuste o nível de logging conforme necessário
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(message)s')

ALLOWED_EXTENSIONS = {'pdf'}

def obter_caminho_recurso(relativo):
    """
    Função para obter o caminho correto no ambiente empacotado ou em ambiente de desenvolvimento.
    Se estiver rodando o executável, os arquivos ficam em sys._MEIPASS.
    """
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relativo)
    else:
        return os.path.join(os.path.abspath("."), relativo)


def allowed_file(filename):
    """
    Verifica se o arquivo enviado possui extensão permitida.
    """
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def extract_text_with_pdfminer_layout(pdf_path):
    """
    Função que extrai o texto de um PDF utilizando PDFMiner,
    mantendo (tanto quanto possível) o layout original. (SICAF2)
    """
    try:
        output_string = StringIO()
        with open(pdf_path, 'rb') as arquivo_pdf:
            extract_text_to_fp(arquivo_pdf, output_string, laparams=None)

        texto_extraido = output_string.getvalue()
        output_string.close()

        # Remove todos os espaços
        texto_extraido = texto_extraido.replace(" ", "")
        # Remove linhas em branco
        texto_extraido = "\n".join([line for line in texto_extraido.splitlines() if line.strip()])

        return texto_extraido
    except Exception as e:
        logging.error(f"Erro ao extrair texto com PDFMiner mantendo layout do PDF {pdf_path}: {e}")
        return ""


def extract_text_with_format_adjustment(pdf_path):
    """
    Extrai o texto de um PDF usando pdfplumber e faz ajustes:
    - Adiciona espaço após 'Formato:' se não houver.
    - Remove linhas em branco.
    """
    try:
        with pdfplumber.open(pdf_path) as pdf:
            text = "\n".join([page.extract_text() for page in pdf.pages])

        text = re.sub(r'(Formato:)(\S)', r'\1 \2', text)
        text = "\n".join([line for line in text.splitlines() if line.strip()])
        return text
    except Exception as e:
        logging.error(f"Erro ao extrair texto do PDF {pdf_path}: {e}")
        return ""


def extract_text_with_format_adjustment_py(pdf_path):
    """
    Extrai o texto de um PDF usando PyPDF2 (PdfReader) e faz ajustes:
    - Adiciona espaço após 'Formato:' se não houver.
    - Remove linhas em branco.
    """
    try:
        with open(pdf_path, "rb") as file:
            reader = PdfReader(file)
            text = "\n".join([page.extract_text() for page in reader.pages])

        text = re.sub(r'(Formato:)(\S)', r'\1 \2', text)
        text = "\n".join([line for line in text.splitlines() if line.strip()])
        return text
    except Exception as e:
        logging.error(f"Erro ao extrair texto do PDF {pdf_path}: {e}")
        return ""


def extract_field_value(text, field_names, below=False, below_lines=1, first_n_chars=None, date_only=False,
                        exclude_pattern=None, exclude_numbers=False, after_dash=False, stop_before=None,
                        stop_after=None, only_numbers=False, line_range=None, check_next_line_if_empty=False,
                        skip_empty_lines=True, split_by=None):
    """
    Extrai o valor de um campo específico no texto, com suporte para várias opções de filtro
    e manipulação de strings (como capturar linha abaixo, extrair somente números, data, etc.).
    """
    lines = text.split('\n')

    if line_range is not None:
        if isinstance(line_range, int):
            lines = [lines[line_range - 1]]
        else:
            lines = lines[line_range[0] - 1:line_range[1]]

    for field_name in field_names:
        for i, line in enumerate(lines):
            start_index = line.find(field_name)
            if start_index != -1:
                if below and i + below_lines < len(lines):
                    field_value = lines[i + below_lines].strip()
                    if skip_empty_lines:
                        while not field_value and i + below_lines < len(lines):
                            i += 1
                            field_value = lines[i + below_lines].strip()

                    if check_next_line_if_empty and not field_value and i + below_lines + 1 < len(lines):
                        field_value = lines[i + below_lines + 1].strip()
                else:
                    field_value = line[start_index + len(field_name):].strip()
                break
        else:
            continue
        break
    else:
        return None

    # Aplicando opções de manipulação
    if stop_before:
        if isinstance(stop_before, list):
            for stop_word in stop_before:
                field_value = field_value.split(stop_word)[0].strip()
        else:
            field_value = field_value.split(stop_before)[0].strip()

    if stop_after:
        stop_index = field_value.find(stop_after)
        if stop_index != -1:
            field_value = field_value[:stop_index].strip()

    if date_only:
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

    if split_by:
        field_value = [
            part.strip()
            for part in re.split(r' *' + re.escape(split_by) + r' *', field_value)
            if part.strip()
        ]

    return field_value.strip() if isinstance(field_value, str) else field_value


def extract_field_values(text, field_names, line_range=None, stop_before=None, below=False, below_lines=1,
                         after_dash=False, exclude_pattern=None, avoid_duplicates_for_formato=True):
    """
    Extrai múltiplos valores de campos no texto, com suporte para diversos filtros.
    Retorna uma lista com todos os valores encontrados.
    """
    values = []
    lines = text.split('\n')

    if line_range:
        lines = lines[line_range[0] - 1:line_range[1]]

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

                    # Evitar duplicidades para formato
                    if 'FORMATO' in field_name.upper() and avoid_duplicates_for_formato and value in values:
                        continue

                    values.append(value)
    return values


def determine_sicaf_type(text):
    """
    Determina o tipo de SICAF com base na presença da palavra 'Relatório'.
    """
    if 'Relatório' in text or 'RELATORIO' in text:
        return 'SICAF1'
    else:
        return 'SICAF2'


def determine_os_type(text):
    """
    Determina o tipo de OS com base na presença de 'E-mail de Leiaute'.
    """
    if 'E-mail de Leiaute' in text:
        return 'OS1'
    else:
        return 'OS2'


def extract_razao_social_from_sicaf(text):
    """
    Extrai a Razão Social de um texto SICAF, tratando quebras de linha.
    """
    lines = text.split('\n')
    razao_social = ""
    for line in lines:
        if line.strip():
            razao_social += " " + line.strip()
    return razao_social.strip()


def normalize_razao_social(razao_social):
    """
    Normaliza a Razão Social para comparação (remove acentos, caracteres especiais, espaços, etc.).
    """
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


def check_format_in_at(at_text, formato_ap):
    """
    Verifica se o formato está presente no corpo de um documento AT.
    """
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


def search_format_in_pdf(pdf_path, search_text):
    """
    Procura diretamente o valor de um formato em um PDF, retornando True se encontrado.
    """
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


def search_peca_in_pdf(pdf_path, peca):
    """
    Procura diretamente o valor de uma peça em um PDF, retornando True se encontrado.
    """
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
    Extrai o CNPJ do texto usando uma expressão regular (padrão XX.XXX.XXX/XXXX-XX).
    Usada especificamente no SICAF2.
    """
    cnpj_pattern = r'\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}'
    match = re.search(cnpj_pattern, text)
    if match:
        return match.group(0)
    return None


def extract_fields(document_text, document_type):
    """
    Extrai campos importantes de acordo com o tipo de documento (OS, AP, AT ou SICAF).
    """
    fields = {}
    lines = document_text.split('\n')
    try:
        if document_type == 'OS':
            os_type = determine_os_type(document_text)
            fields['OS_TYPE'] = os_type

            if os_type == 'OS1':
                # Extração para OS1
                fields.update({
                    'OS N°': extract_field_value(document_text, ['OS Nº', 'OS N°'], only_numbers=True),
                    'DATA DE INICIO': extract_field_value(document_text, ['DATA DE INICIO:', 'DATA DE INÍCIO'],
                                                          below=True, date_only=True),
                    'TITULO DA OS': extract_field_value(document_text, ['TITULO DA OS:', 'TÍTULO DA OS'],
                                                        below=True, check_next_line_if_empty=True),
                    'ORGAO': extract_field_value(document_text, ['ORGAO', 'ÓRGÃO'], below=True,
                                                 exclude_pattern=r'\d{2}/\d{2}/\d{4}', after_dash=True),
                    'TIPO DA CAMPANHA': extract_field_value(document_text,
                                                            ['Nº DO PROCESSO DE SELEÇÃO INTERNA:'],
                                                            below=True,
                                                            stop_before=' N° ',
                                                            exclude_numbers=True)
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
            # Extração para o AP
            fields.update({
                'OS N°': extract_field_value(document_text, ['OS N°', 'OS Nº', 'OSNº'],
                                             stop_before='VALOR', only_numbers=True),
                'DATA EMISSAO': extract_field_value(document_text,
                                                    ['DATA EMISSAO', 'DATA EMISSÃO', 'DATAEMISSÃO', 'DATA  EMISSÃO:'],
                                                    date_only=True),
                'CAMPANHA': extract_field_value(document_text, ['CAMPANHA:'],
                                                stop_before=['AUT.', 'MEIO:']),
                'PRODUTO': extract_field_value(document_text, ['PRODUTO:'], stop_before=' '),
                'AUT.CLIENTE': extract_field_value(document_text, ['AUT.CLIENTE:'], check_next_line_if_empty=True),
                'AT DE PRODUCAO': extract_field_value(document_text,
                                                      ['AT DE PRODUCAO:', 'AT DE PRODUÇÃO:', 'AT DE PRODUCAO',
                                                       'AT DE PRODUÇÃO', 'ATDEPRODUÇÃO', "AT'SDEPRODUÇÃO",
                                                       "AT'S DE PRODUÇÃO"],
                                                      stop_before='-',
                                                      exclude_pattern=':',
                                                      split_by='E'),
            })

            # Ajuste de CNPJ, ignorando 16.088.593
            cnpj_value = extract_field_value(document_text, ['Cnpj: ', 'CNPJ:'])
            if cnpj_value and "16.088.593" in cnpj_value:
                cnpj_pattern = re.compile(r'\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}')
                for line in lines:
                    cnpj_match = cnpj_pattern.search(line)
                    if cnpj_match and "16.088.593" not in cnpj_match.group():
                        cnpj_value = cnpj_match.group()
                        break
            fields['CNPJ'] = cnpj_value

            # Busca índice da linha do CNPJ para extrair Razão Social e Município
            cnpj_line_index = None
            for i, line in enumerate(lines):
                if cnpj_value and cnpj_value in line:
                    cnpj_line_index = i
                    break

            if cnpj_line_index is not None:
                # Razão Social (uma linha acima não vazia)
                razao_social_lines = [
                    line.strip()
                    for line in lines[max(0, cnpj_line_index - 1):cnpj_line_index]
                    if line.strip()
                ]
                if razao_social_lines:
                    fields['Razão social'] = razao_social_lines[0].strip()

                # Município (três linhas acima)
                municipio_lines = [
                    line.strip()
                    for line in lines[max(0, cnpj_line_index - 3):cnpj_line_index]
                    if line.strip()
                ]
                if municipio_lines:
                    fields['Município'] = municipio_lines[0].split('-')[0].strip()
                else:
                    fields['Município'] = ""

            # Extração dos campos de PEÇAS
            pecas = extract_field_values(document_text, ['PEÇA', 'PECA'], stop_before='FORMATO', after_dash=True)
            for i, peca in enumerate(pecas):
                fields[f'PECA{i + 1}'] = peca

            # Extração dos campos de FORMATO
            formatos = extract_field_values(document_text, ['FORMATO'])
            for i, formato in enumerate(formatos):
                fields[f'FORMATO{i + 1}'] = formato.strip().strip(':').replace(' ', '')

        elif document_type == 'AT':
            # Extração para AT
            fields.update({
                'AT': extract_field_value(document_text, ['AT '], stop_before='DATA'),
                'TITULO': extract_field_value(document_text, ['TITULO: ', 'TÍTULO:', 'Título:'],
                                              stop_before=['Cores', 'CORES'])
            })
            formatos = extract_field_values(document_text, ['FORMATO:', 'Formato'])
            for i, formato in enumerate(formatos):
                fields[f'FORMATO{i + 1}'] = formato
            fields.update({
                'Data da AT': extract_field_value(document_text, ['Data:', 'DATA:'], date_only=True)
            })

        elif document_type == 'SICAF':
            # Extração para SICAF
            sicaf_type = determine_sicaf_type(document_text)
            fields['SICAF_TYPE'] = sicaf_type
            if sicaf_type == 'SICAF1':
                fields.update({
                    'Razão social': extract_field_value(document_text, ['Razao Social:', 'Razão Social:']),
                    'CNPJ': extract_field_value(document_text, ['CNPJ: ', 'CNPJ:'], stop_before='Data'),
                    'Município': extract_field_value(document_text, ['Municipio: ', 'Munícipio:'],
                                                     stop_before=' N°')
                })
            else:
                fields.update({
                    'CNPJ': extract_cnpj(document_text)
                })
    except Exception as e:
        logging.error(f"Erro ao extrair campos para {document_type}: {e}")
    return fields


def check_razao_social_in_ap(ap_text, razao_social):
    """
    Verifica se a Razão Social do SICAF está presente no AP, especificamente nas linhas 21 ou 22.
    """
    if not razao_social:
        return False
    lines = ap_text.split('\n')
    if len(lines) >= 22:
        if razao_social in lines[20] or razao_social in lines[21]:
            return True
    return False


def check_municipio_in_ap(ap_text, municipio):
    """
    Verifica se o Município do SICAF está presente no AP entre as linhas 18 e 23.
    """
    if not municipio:
        return False
    lines = ap_text.split('\n')
    for i in range(17, 23):
        if i < len(lines) and municipio in lines[i]:
            return True
    return False


def check_peca_in_at(at_text, peca):
    """
    Verifica se a peça (string) está presente no texto de um documento AT.
    """
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


def determine_overall_status(field_statuses, required_pieces, found_pieces, fields_to_verify=None):
    """
    Determina o status geral do processo (OK ou NC) com base nos campos checados e nas peças encontradas.
    """
    relevant = field_statuses
    if fields_to_verify:
        relevant = {k: v for k, v in field_statuses.items() if k in fields_to_verify}

    all_fields_ok = all(
        str(status).strip().upper() == 'OK'
        for status in relevant.values()
        if status is not None
    )

    if all_fields_ok:
        all_pieces_match = set(required_pieces) == set(found_pieces)
        if all_pieces_match:
            overall_status = 'OK'
            status_class = 'status-ok'
        else:
            overall_status = 'NC'
            status_class = 'status-nc'
    else:
        overall_status = 'NC'
        status_class = 'status-nc'

    return overall_status, status_class


def generate_html_report(report, subfolder_name, overall_status, status_class, fields_to_verify=None):
    """
    Gera o relatório HTML final com base em um relatório de texto,
    nome da subpasta, status geral do processo e classe CSS de status.
    """
    html_report = f"""
    <div class="report {status_class}">
        <h2>{escape(subfolder_name)}</h2>
        <div class="document-sections-container">
    """

    lines = report.split('\n')
    index = 0
    while index < len(lines):
        line = lines[index].strip()

        # ---------------- OS ----------------
        if line.startswith('- OS:'):
            os_data_str = line[len('- OS:'):].strip()
            try:
                os_data = ast.literal_eval(os_data_str)
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

        # ---------------- AP ----------------
        elif line.startswith('- AP:'):
            ap_data_str = line[len('- AP:'):].strip()
            try:
                ap_data = ast.literal_eval(ap_data_str)
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

        # ---------------- AT ----------------
        elif line.startswith('- AT ('):
            at_match = re.match(r'- AT \((.*?)\): (.*)', line)
            if at_match:
                at_file_name = at_match.group(1)
                at_data_str = at_match.group(2).strip()
                try:
                    at_data = ast.literal_eval(at_data_str)
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

        # ---------------- SICAF ----------------
        elif line.startswith('- SICAF:'):
            sicaf_data_str = line[len('- SICAF:'):].strip()
            try:
                sicaf_data = ast.literal_eval(sicaf_data_str)
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
        else:
            pass

        index += 1

    html_report += '</div>'  # Fecha .document-sections-container

    # Processa as linhas de CHECK
    for line in lines:
        line = line.strip()
        if 'CHECK' in line:
            idx = line.find('CHECK')
            field = escape(line[:idx].strip())
            if fields_to_verify and field not in fields_to_verify:
                continue
            value = escape(line[idx:].strip())

            # Determina classe CSS baseada no resultado
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

    # Adiciona o status geral
    html_report += f"""
    <div class="overall-status {status_class}">
        <strong>Status do Processo:</strong> {overall_status}
    </div>
    """
    html_report += "</div>"  # Fecha a div .report

    return html_report


def save_and_open_report(html_report):
    """
    Retorna o relatório HTML gerado para ser renderizado ou salvo em arquivo.
    (Nesta implementação, apenas retornamos a string gerada).
    """
    return html_report


def save_ap_text(ap_text):
    """
    Salva o texto extraído do AP em um arquivo texto (exemplo simples).
    Caso precise armazenar em outro local, basta ajustar o caminho.
    """
    with open("../ap_text.txt", "w", encoding="utf-8") as file:
        file.write(ap_text)


def search_text_in_pdf(pdf_path, search_text):
    """
    Procura o texto no PDF e retorna o resultado concatenado das ocorrências (exemplo simples).
    Se o texto existir sem espaços, retorna a própria string procurada.
    """
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


def check_fields(os_fields, ap_fields, at_fields_list, sicaf_fields,
                 sicaf_pdf_path, at_files, missing_at_numbers, found_pieces):
    """
    Função principal de verificação e comparação de campos extraídos de OS, AP, AT e SICAF.
    Gera um relatório (em texto) com base nos checks realizados.
    """
    report = [
        "Transcrição das informações extraídas:",
        f"- OS: {os_fields}",
        f"- AP: {ap_fields}"
    ]

    if not os_fields or not ap_fields or not sicaf_fields:
        error_message = "Erro ao extrair campos dos documentos."
        logging.error(error_message)
        return "", error_message

    if at_fields_list:
        for at_fields in at_fields_list:
            at_file_name = at_fields.get('FILE_NAME', f"- AT {at_fields_list.index(at_fields) + 1}")
            report.append(f"- AT ({at_file_name}): {at_fields}")
    else:
        report.append("- AT: Nenhum arquivo AT encontrado.")

    report.append(f"- SICAF: {sicaf_fields}")

    # -------------------------
    # CHECK 1 - Verificações comuns (OS x AP)
    # -------------------------
    try:
        # 1.1 OS N° vs AP OS N°
        if os_fields.get('OS N°') == ap_fields.get('OS N°'):
            report.append(f"OS N°                                   CHECK 1.1: OK - OS N° {os_fields.get('OS N°')}")
        else:
            report.append(f"OS N°                                   CHECK 1.1: Non-conformity - OS N° {ap_fields.get('OS N°')}")

        # 1.2 Data de início (OS) vs Data de emissão (AP)
        os_data_inicio = os_fields.get('DATA DE INICIO')
        ap_data_emissao = ap_fields.get('DATA EMISSAO')
        if os_data_inicio and ap_data_emissao:
            os_data_inicio_dt = datetime.strptime(os_data_inicio, '%d/%m/%Y')
            ap_data_emissao_dt = datetime.strptime(ap_data_emissao, '%d/%m/%Y')
            if ap_data_emissao_dt > os_data_inicio_dt:
                report.append("DATAS                                   CHECK 1.2: OK")
            else:
                report.append("DATAS                                   CHECK 1.2: Non-conformity")
        else:
            report.append("DATAS                                   CHECK 1.2: Non-conformity")
    except Exception as e:
        report.append(f"DATAS                                   CHECK 1.2: Error {e}")

    # 1.3 TÍTULO DA OS vs CAMPANHA
    if os_fields.get('TITULO DA OS') == ap_fields.get('CAMPANHA'):
        report.append("TITULO DA OS/CAMPANHA                   CHECK 1.3: OK")
    else:
        report.append("TITULO DA OS/CAMPANHA                   CHECK 1.3: Non-conformity")

    # 1.4 ORGAO vs PRODUTO
    orgao = os_fields.get('ORGAO', '')
    produto = ap_fields.get('PRODUTO', '')
    if orgao and produto and orgao == produto:
        report.append("ORGAO/PRODUTO                           CHECK 1.4: OK")
    else:
        report.append("CHECK 1.4: Non-conformity")

    # 1.5 TIPO DA CAMPANHA vs AUT.CLIENTE
    if os_fields.get('TIPO DA CAMPANHA') == ap_fields.get('AUT.CLIENTE'):
        report.append("TIPO DA CAMPANHA/AUT.CLIENTE            CHECK 1.5: OK")
    else:
        report.append("TIPO DA CAMPANHA/AUT.CLIENTE            CHECK 1.5: Non-conformity")
   # ------------------------- CHECK 2 - Verificação de ATs (AP vs AT) -------------------------

    # 1) Identificar peças do AP
    required_pieces = []
    for key, value in ap_fields.items():
        if key.startswith('PECA') and value:
            required_pieces.append(value.strip())

    # 2) Dicionário p/ dizer em quais ATs cada peça foi encontrada
    found_in_at = {}
    for p in required_pieces:
        found_in_at[p] = set()

    # 3) ATs declarados no AP
    ap_value_list = ap_fields.get('AT DE PRODUCAO', [])
    if ap_value_list is None:
        ap_value_list = []
    if isinstance(ap_value_list, str):
        ap_value_list = [ap_value_list.strip()]
    else:
        ap_value_list = [str(item).strip() for item in ap_value_list]

    # 4) Vamos também montar os formatos do AP (para checar 1x por AT)
    formatos_ap = [
        ap_fields.get(f'FORMATO{j + 1}')
        for j in range(len(ap_fields))
        if ap_fields.get(f'FORMATO{j + 1}')
    ]
    formatos_ap_unicos = list(set(formatos_ap))

    # 5) Iterar cada arquivo AT
    for i, at_file in enumerate(at_files):
        at_file_name = os.path.basename(at_file)
        match = re.search(r'AT\s*(\d+)', at_file_name)
        if not match:
            continue
        at_number = match.group(1)

        # Checa se esse AT está no AP
        if at_number in ap_value_list:
            # Exibe OK do /AT DE PRODUCAO
            report.append(f"AT {at_number} - ({at_file_name}) /AT DE PRODUCAO CHECK 2.{i+1}.1: OK")

            # Lê texto do PDF (sem espaços) p/ comparar peças
            with pdfplumber.open(at_file) as pdf:
                at_text_no_spaces = re.sub(
                    r'\s+',
                    '',
                    "".join([page.extract_text() or "" for page in pdf.pages])
                ).upper()

            # (A) Marcar quais peças do AP aparecem neste AT
            for peca in required_pieces:
                peca_no_spaces = re.sub(r'\s+', '', peca.upper())
                if peca_no_spaces in at_text_no_spaces:
                    found_in_at[peca].add(at_number)

            # (B) Verificar Formatos (1x por AT)
            # Acha o dicionário at_fields correspondente
            at_fields = None
            for atf in at_fields_list:
                if atf.get('FILE_NAME') == at_file_name:
                    at_fields = atf
                    break

            if at_fields:
                # Formatos do AT
                formatos_at = [
                    at_fields.get(f'FORMATO{j+1}')
                    for j in range(len(at_fields))
                    if at_fields.get(f'FORMATO{j+1}')
                ]
            else:
                formatos_at = []

            # Checar cada formato do AP
            for formato_ap in formatos_ap_unicos:
                matched = False
                if formato_ap:
                    # Se está no dicionário at_fields
                    if formato_ap in formatos_at:
                        matched = True
                    else:
                        # Ou checar diretamente no PDF
                        if search_format_in_pdf(at_file, formato_ap):
                            matched = True

                if matched:
                    report.append(
                        f"AT - ({at_file_name}) FORMATO/FORMATO "
                        f"CHECK 2.{i+1}.3: OK - Formato {formato_ap}"
                    )
                else:
                    report.append(
                        f"AT - ({at_file_name}) FORMATO/FORMATO "
                        f"CHECK 2.{i+1}.3: Non-conformity - Formato {formato_ap}"
                    )

            # (C) Verificação data (1x por AT)
            try:
                ap_data_emissao_dt = datetime.strptime(
                    ap_fields.get('DATA EMISSAO', '01/01/1900'),
                    '%d/%m/%Y'
                )
                if at_fields:
                    at_data_at_dt = datetime.strptime(
                        at_fields.get('Data da AT', '01/01/1900'),
                        '%d/%m/%Y'
                    )
                    if ap_data_emissao_dt >= at_data_at_dt:
                        report.append(f"DATA EMISSAO/Data da AT                 CHECK 2.{i+1}.4: OK")
                    else:
                        report.append(f"DATA EMISSAO/Data da AT                 CHECK 2.{i+1}.4: Non-conformity")
                else:
                    # Se at_fields é None
                    report.append(
                        f"DATA EMISSAO/Data da AT                 CHECK 2.{i+1}.4: Non-conformity - "
                        f"at_fields não encontrado para '{at_file_name}', data não verificada."
                    )
            except Exception as e:
                report.append(f"DATA EMISSAO/Data da AT                 CHECK 2.{i+1}.4: Error {e}")

        else:
            # Se o AT não consta na lista do AP
            continue

    # 6) Ao final, gerar OK ou NC para cada peça
    for peca in required_pieces:
        if found_in_at[peca]:
            # Apareceu em algum AT
            ats_encontrados = ', '.join(found_in_at[peca])
            report.append(
                f"CHECK 2.2: OK - A peça '{peca}' foi encontrada nos ATs: {ats_encontrados}"
            )
            found_pieces.append(peca.upper())
        else:
            report.append(
                f"CHECK 2.2: Non-conformity - A peça '{peca}' não foi encontrada em nenhum AT"
            )

    # -------------------------
    # CHECK 3 - SICAF verificações
    # -------------------------
    if sicaf_fields.get('SICAF_TYPE') == 'SICAF1':
        # SICAF1
        razao_social_sicaf = normalize_razao_social(sicaf_fields.get('Razão social', ''))
        razao_social_ap = normalize_razao_social(ap_fields.get('Razão social', ''))
        if razao_social_sicaf == razao_social_ap:
            report.append("Razão social                            CHECK 3.1: OK")
        else:
            report.append("Razão social                            CHECK 3.1: Non-conformity")

        if sicaf_fields.get('CNPJ') == ap_fields.get('CNPJ'):
            report.append("CNPJ                                    CHECK 3.2: OK")
        else:
            report.append("CNPJ                                    CHECK 3.2: Non-conformity")

        if sicaf_fields.get('Município') == ap_fields.get('Município'):
            report.append("Município                               CHECK 3.3: OK")
        else:
            report.append("Município                               CHECK 3.3: Non-conformity")

    else:
        # SICAF2 - Pesquisa direta no PDF
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
    return report_text, None


def save_text_to_file(text, file_name, folder_path):
    """
    Salva o texto extraído em um arquivo de texto no 'folder_path'.
    """
    if not folder_path:
        logging.error("folder_path is None")
        return
    file_path = os.path.join(folder_path, file_name)
    with open(file_path, "w", encoding="utf-8") as file:
        file.write(text)
    logging.info(f"Texto salvo em {file_path}")


def verify_documents(file_paths, subfolder_name, temp_pdf_dir, fields_to_verify=None, move_os_at_files=True):
    """
    Função principal que faz a verificação dos documentos:
    - Extrai texto e campos de OS, AP, AT e SICAF.
    - Gera relatório de não conformidades ou OK.
    - Move arquivos para as pastas OK ou Non-conformity, caso necessário.
    - Retorna o relatório HTML final e o status geral.
    """
    os_file = file_paths.get('OS')
    ap_file = file_paths.get('AP')
    at_files = file_paths.get('AT', [])
    sicaf_file = file_paths.get('SICAF')

    if not all([os_file, ap_file, sicaf_file]):
        error_message = f"Todos os arquivos (OS, AP, SICAF) devem estar presentes na pasta {subfolder_name}."
        logging.error(error_message)
        return "", None, error_message

    folder_path = os.path.dirname(__file__)

    # Extrai texto OS
    os_text = extract_text_with_format_adjustment(os_file)
    determine_os_type(os_text)  # Força detecção do tipo de OS
    os_fields = extract_fields(os_text, 'OS')
    save_text_to_file(os_text, f"os_text_{subfolder_name}.txt", temp_pdf_dir)

    # Extrai texto AP
    ap_text = extract_text_with_format_adjustment_py(ap_file)
    ap_fields = extract_fields(ap_text, 'AP')
    save_text_to_file(ap_text, f"ap_text_{subfolder_name}.txt", temp_pdf_dir)

    # Processa os ATs
    at_fields_list = []
    at_de_producao = ap_fields.get('AT DE PRODUCAO')
    if isinstance(at_de_producao, list):
        at_numbers_in_ap = [num.strip() for num in at_de_producao]
    elif at_de_producao:
        at_numbers_in_ap = [at_de_producao.strip()]
    else:
        at_numbers_in_ap = []

    at_numbers_found = []

    for at_file in at_files:
        at_text = extract_text_with_format_adjustment(at_file)
        at_fields = extract_fields(at_text, 'AT')
        at_fields['FILE_NAME'] = os.path.basename(at_file)
        at_number = (at_fields.get('AT') or "").strip()

        if at_number in at_numbers_in_ap:
            at_fields_list.append(at_fields)
            at_numbers_found.append(at_number)
            save_text_to_file(at_text, f"at_text_{os.path.basename(at_file)}.txt", temp_pdf_dir)
        else:
            # Se o número do AT não está no AP, não processa
            pass

    missing_at_numbers = set(at_numbers_in_ap) - set(at_numbers_found)

    # Extrai texto SICAF
    sicaf_text = extract_text_with_format_adjustment(sicaf_file)
    sicaf_type = determine_sicaf_type(sicaf_text)
    if sicaf_type == 'SICAF2':
        sicaf_text = extract_text_with_pdfminer_layout(sicaf_file)
    save_text_to_file(sicaf_text, f"sicaf_text_{subfolder_name}.txt", temp_pdf_dir)
    sicaf_fields = extract_fields(sicaf_text, 'SICAF')

    found_pieces = []
    report, error_message = check_fields(
        os_fields, ap_fields, at_fields_list,
        sicaf_fields, sicaf_file,
        at_files, missing_at_numbers,
        found_pieces
    )
    if error_message:
        return "", None, error_message

    required_pieces = [
        value.strip().upper()
        for key, value in ap_fields.items()
        if key.startswith('PECA')
    ]

    # Monta dicionário dos status
    field_statuses = {
        'OS N°': 'OK' if 'CHECK 1.1' in report and 'OK' in report else 'NC',
        'DATAS': 'OK' if 'CHECK 1.2' in report and 'OK' in report else 'NC',
        'TITULO DA OS/CAMPANHA': 'OK' if 'CHECK 1.3' in report and 'OK' in report else 'NC',
        'ORGAO/PRODUTO': 'OK' if 'CHECK 1.4' in report and 'OK' in report else 'NC',
        'TIPO DA CAMPANHA/AUT.CLIENTE': 'OK' if 'CHECK 1.5' in report and 'OK' in report else 'NC',
        'AT /AT DE PRODUCAO': 'OK' if any(f'CHECK 2.{i}.1' in report and 'OK' in report for i in range(1, 10)) else 'NC',
        'AT FORMATO/FORMATO': 'OK' if any(f'CHECK 2.{i}.3' in report and 'OK' in report for i in range(1, 10)) else 'NC',
        'DATA EMISSAO/Data da AT': 'OK' if any(f'CHECK 2.{i}.4' in report and 'OK' in report for i in range(1, 10)) else 'NC',
        'Razão social': 'OK' if 'CHECK 3.1' in report and 'OK' in report else 'NC',
        'CNPJ': 'OK' if 'CHECK 3.2' in report and 'OK' in report else 'NC',
        'Município': 'OK' if 'CHECK 3.3' in report and 'OK' in report else 'NC'
    }

    overall_status, status_class = determine_overall_status(
        field_statuses, required_pieces, found_pieces, fields_to_verify
    )

    # Pasta "Relatorios" onde "OK" e "Non-conformity" serão criadas
    relatorios_folder = os.path.join(folder_path, "Relatorios")
    os.makedirs(relatorios_folder, exist_ok=True)

    if overall_status == 'OK':
        status_folder = os.path.join(relatorios_folder, "OK", subfolder_name)
    else:
        status_folder = os.path.join(relatorios_folder, "Non-conformity", subfolder_name)

    os.makedirs(status_folder, exist_ok=True)

    # Move os arquivos para a pasta de destino
    try:
        shutil.move(ap_file, os.path.join(status_folder, os.path.basename(ap_file)))
        shutil.move(sicaf_file, os.path.join(status_folder, os.path.basename(sicaf_file)))

        if overall_status == 'Non-conformity' and move_os_at_files:
            shutil.move(os_file, os.path.join(status_folder, os.path.basename(os_file)))
            for at_file in at_files:
                shutil.move(at_file, os.path.join(status_folder, os.path.basename(at_file)))

        logging.info(f"Arquivos movidos para a pasta: {status_folder}")
    except Exception as e:
        logging.error(f"Erro ao mover os arquivos: {e}")

    html_report = generate_html_report(report, subfolder_name, overall_status, status_class, fields_to_verify)
    return html_report, overall_status, None


def move_relatorios_folder(destination_path):
    """
    Move a pasta 'Relatorios' para o 'destination_path', renomeando com data e hora.
    """
    try:
        folder_path = os.path.dirname(__file__)
        relatorios_folder = os.path.join(folder_path, "Relatorios")
        if os.path.exists(relatorios_folder):
            now = datetime.now()
            timestamp = now.strftime("%Y%m%d_%H%M%S")
            new_folder_name = f"Relatorios_{timestamp}"
            destination_folder = os.path.join(destination_path, new_folder_name)
            if not os.path.exists(destination_path):
                os.makedirs(destination_path)
            shutil.move(relatorios_folder, destination_folder)
            logging.info(f"Pasta 'Relatorios' movida para {destination_folder}")
        else:
            logging.warning("A pasta 'Relatorios' não existe e não pode ser movida.")
    except Exception as e:
        logging.error(f"Erro ao mover a pasta 'Relatorios': {e}")


def delete_temp_folder():
    """
    Aguarda 45 segundos e remove a pasta temporária 'temp_pdf' caso exista.
    """
    time.sleep(45)
    temp_folder = os.path.join(os.path.dirname(__file__), 'temp_pdf')
    try:
        shutil.rmtree(temp_folder)
        logging.info(f"Pasta temporária {temp_folder} apagada com sucesso.")
    except Exception as e:
        logging.error(f"Erro ao tentar apagar a pasta temporária {temp_folder}: {e}")
