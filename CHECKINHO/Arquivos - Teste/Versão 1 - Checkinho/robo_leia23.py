import pdfplumber
from datetime import datetime
import logging
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import re
import webbrowser

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(message)s')


def extract_text_with_format_adjustment(pdf_path):
    """Extrai o texto de um PDF usando pdfplumber, ajusta o formato conforme necessário e remove linhas em branco."""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            text = "\n".join([page.extract_text() for page in pdf.pages])
        # Adiciona um espaço após "Formato:" se não houver
        text = re.sub(r'(Formato:)(\S)', r'\1 \2', text)
        # Remove linhas em branco
        text = "\n".join([line for line in text.splitlines() if line.strip()])
        return text
    except Exception as e:
        logging.error(f"Erro ao extrair texto do PDF {pdf_path}: {e}")
        return ""


def extract_field_value(text, field_names, below=False, below_lines=1, first_n_chars=None, date_only=False,
                        exclude_pattern=None, exclude_numbers=False, after_dash=False, stop_before=None,
                        stop_after=None, only_numbers=False, line_range=None, check_next_line_if_empty=False,
                        skip_empty_lines=True, split_by=None):
    """Extrair o valor de um campo do texto, com suporte à separação por um delimitador se split_by for fornecido."""
    lines = text.split('\n')

    if line_range is not None:
        if isinstance(line_range, int):
            lines = [lines[line_range - 1]]  # Captura apenas a linha especificada (indexada a partir de 1)
        else:
            lines = lines[line_range[0] - 1:line_range[1]]  # Ajustar intervalo de linhas (indexado a partir de zero)

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

    if stop_before:
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

    # Se split_by for fornecido, dividir o valor extraído em múltiplos valores
    if split_by:
        field_value = [part.strip() for part in re.split(f"[{re.escape(split_by)}]", field_value) if part.strip()]

    return field_value.strip() if isinstance(field_value, str) else field_value



def extract_field_values(text, field_names, line_range=None, stop_before=None, below=False, below_lines=1,
                         after_dash=False):
    """Extrair vários valores de campos do texto, com suporte a captura abaixo do rótulo."""
    values = []
    lines = text.split('\n')

    if line_range:
        lines = lines[line_range[0] - 1:line_range[1]]  # Ajustar intervalo de linhas (indexado a partir de zero)

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
                        # Ignorar o "-" e o que há depois (por exemplo, "-A", "-B", "-C", etc.)
                        value = re.sub(r'-[A-Z] ', '', value).strip()
                        # Extrair valor até antes da palavra "FORMATO"
                        value = value.split('FORMATO')[0].strip()

                    if 'FORMATO' in field_name.upper():
                        # Ignorar o "-" e o que há antes dele
                        parts = value.split('-')
                        if len(parts) > 1:
                            value = parts[1].strip()

                    if after_dash:
                        parts = value.split('-')
                        if len(parts) > 1:
                            value = parts[1].strip()

                    values.append(value)

    return values


def determine_sicaf_type(text):
    """Determina o tipo de SICAF com base na presença da palavra 'Relatorio'."""
    if 'Relatório' in text or 'RELATORIO' in text:
        return 'SICAF1'
    else:
        return 'SICAF2'


def determine_os_type(text):
    """Determina o tipo de OS com base na presença da frase 'E-mail de Leiaute'."""
    if 'E-mail de Leiaute' in text:
        return 'OS1'
    else:
        return 'OS2'


def extract_razao_social_from_sicaf(text):
    """Extrai a Razão Social do SICAF, lidando com quebras de linha."""
    lines = text.split('\n')
    razao_social = ""
    for line in lines:
        if line.strip():  # Ignorar linhas vazias
            razao_social += " " + line.strip()
    return razao_social.strip()


def normalize_razao_social(razao_social):
    """Normaliza a Razão Social para comparação."""
    return re.sub(r'\s+', ' ', razao_social.strip().upper())


def check_format_in_at(at_text, formato_ap):
    """Verifica se o formato está presente no corpo do documento AT, lidando com variações de aspas e outros
    caracteres."""
    if not formato_ap:
        return False

    # Normaliza o formato e o texto do AT substituindo diferentes tipos de aspas e apóstrofes por aspas padrão
    formato_ap = re.sub(r'[“”″\'"‘’]', '"', formato_ap.strip())
    at_text_normalized = re.sub(r'[“”″\'"‘’]', '"', at_text)

    # Remove quaisquer aspas no final do formato normalizado
    formato_ap = formato_ap.strip('"')

    # Verifica se o formato aparece no texto
    pattern = re.compile(re.escape(formato_ap), re.IGNORECASE)
    lines = at_text_normalized.split('\n')
    for i, line in enumerate(lines):  # Verifica em todas as linhas para garantir abrangência
        if pattern.search(line):
            return True

    return False


def search_format_in_pdf(pdf_path, search_text):
    """Procura o valor do formato diretamente no PDF usando pdfplumber."""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            text = "\n".join([page.extract_text() for page in pdf.pages])

        # Normaliza o texto e o valor de pesquisa para remover variações de aspas
        normalized_text = re.sub(r'[“”″\'"‘’]', '"', text)
        normalized_search_text = re.sub(r'[“”″\'"‘’]', '"', search_text)

        # Verifica se o valor de formato normalizado aparece no texto do PDF
        if normalized_search_text in normalized_text:
            return True
    except Exception as e:
        logging.error(f"Erro ao procurar formato no PDF {pdf_path}: {e}")

    return False


def search_peca_in_pdf(pdf_path, peca):
    """Procura a peça diretamente no PDF usando pdfplumber."""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            text = "\n".join([page.extract_text() for page in pdf.pages])

        # Normaliza o texto e o valor da peça para remover variações de aspas e espaços
        normalized_text = re.sub(r'[“”″\'"‘’]', '"', text)
        normalized_search_text = re.sub(r'[“”″\'"‘’]', '"', peca)

        # Verifica se a peça normalizada aparece no texto do PDF
        if normalized_search_text in normalized_text:
            return True
    except Exception as e:
        logging.error(f"Erro ao procurar a peça no PDF {pdf_path}: {e}")

    return False


def extract_fields(document_text, document_type):
    """Extraia campos do texto do documento com base no tipo de documento."""
    fields = {}
    lines = document_text.split('\n')  # Garante que 'lines' seja sempre definido
    try:
        if document_type == 'OS':
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
                    'TIPO DA CAMPANHA': extract_field_value(document_text, ['TIPO DA CAMPANHA'], below=True,
                                                            stop_before=' N° ', below_lines=2, )
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
            fields.update({
                'OS N°': extract_field_value(document_text, ['OS N°', 'OS Nº'], stop_before='VALOR', only_numbers=True),
                'DATA EMISSAO': extract_field_value(document_text, ['DATA EMISSAO', 'DATA EMISSÃO', 'DATAEMISSÃO'],
                                                    date_only=True),
                'CAMPANHA': extract_field_value(document_text, ['CAMPANHA:'], stop_before='AUT.'),
                'PRODUTO': extract_field_value(document_text, ['PRODUTO:'], stop_before=' '),
                'AUT.CLIENTE': extract_field_value(document_text, ['AUT.CLIENTE:'], check_next_line_if_empty=True),
                'AT DE PRODUCAO': extract_field_value(document_text,
                                                      ['S DE PRODUÇÃO:', 'AT DE PRODUCAO', 'AT DE PRODUÇÃO'],
                                                      split_by="E""/", stop_before='-'),
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
                # Ignorar linhas vazias e capturar Município três linhas acima
                municipio_lines = [line.strip() for line in lines[max(0, cnpj_line_index - 3):cnpj_line_index] if
                                   line.strip()]
                fields['Município'] = municipio_lines[0].split('-')[0].strip() if municipio_lines else ""

                # Ignorar linhas vazias e capturar Razão Social uma linha acima
                razao_social_lines = [line.strip() for line in lines[max(0, cnpj_line_index - 1):cnpj_line_index] if
                                      line.strip()]
                fields['Razão social'] = razao_social_lines[0] if razao_social_lines else ""

            # Extração dos formatos e peças na linha de baixo
            pecas = extract_field_values(document_text, ['PEÇA', 'PECA'], stop_before='FORMATO', after_dash=True)
            for i, peca in enumerate(pecas):
                fields[f'PECA{i + 1}'] = peca

            formatos = extract_field_values(document_text, ['FORMATO'])
            for i, formato in enumerate(formatos):
                fields[f'FORMATO{i + 1}'] = formato.strip().strip('"')

        elif document_type == 'AT':
            fields.update({
                'AT': extract_field_value(document_text, ['AT '], stop_before='DATA'),
                'TITULO': extract_field_value(document_text, ['TITULO: ', 'TÍTULO:', 'Título:'], stop_before='Cores'),
            })
            formatos = extract_field_values(document_text, ['FORMATO:', 'Formato'])
            for i, formato in enumerate(formatos):
                fields[f'FORMATO{i + 1}'] = formato
            fields.update({
                'Data da AT': extract_field_value(document_text, ['Data:', 'DATA:'], date_only=True)
            })
        elif document_type == 'SICAF':
            sicaf_type = determine_sicaf_type(document_text)
            fields['SICAF_TYPE'] = sicaf_type
            if sicaf_type == 'SICAF1':
                # Extração para SICAF1
                fields.update({
                    'Razao Social': extract_field_value(document_text, ['Razao Social:','Razão Social:']),
                    'CNPJ': extract_field_value(document_text, ['CNPJ: ', 'CNPJ:'], stop_before='Data'),
                    'Município': extract_field_value(document_text, ['Municipio: ', 'Munícipio:'], stop_before=' N°')
                })
            else:
                # Extração para SICAF2
                fields.update({
                    'Razao Social': None,
                    'CNPJ': None,
                    'Endereco': None,
                })
    except Exception as e:
        logging.error(f"Erro ao extrair campos para {document_type}: {e}")
    return fields


def check_razao_social_in_ap(ap_text, razao_social):
    """Verifique se a Razão Social do SICAF está presente no texto do AP nas linhas 21 ou 22."""
    if not razao_social:
        return False
    lines = ap_text.split('\n')
    if len(lines) >= 22:
        if razao_social in lines[20] or razao_social in lines[21]:
            return True
    return False


def check_municipio_in_ap(ap_text, municipio):
    """Verifique se o Município do SICAF está presente no texto do AP entre as linhas 18 e 23."""
    if not municipio:
        return False
    lines = ap_text.split('\n')
    for i in range(17, 23):  # Os números das linhas são indexados a partir de zero
        if i < len(lines) and municipio in lines[i]:
            return True
    return False


def check_peca_in_at(at_text, peca):
    """Verifica se a peça está presente no corpo do documento AT, considerando variações."""
    if not peca:
        return False

    # Normaliza as aspas, apóstrofes e espaços no peca e no texto do AT
    peca_normalized = re.sub(r'[“”″\'"‘’]', '"', peca.strip())
    at_text_normalized = re.sub(r'[“”″\'"‘’]', '"', at_text)

    # Remove espaços extras para garantir comparação correta
    peca_normalized = re.sub(r'\s+', ' ', peca_normalized)
    at_text_normalized = re.sub(r'\s+', ' ', at_text_normalized)

    # Verifica se a peça completa ou fragmentos dela estão no texto
    if peca_normalized in at_text_normalized:
        return True

    # Divide a peça em fragmentos e verifica cada um
    fragments = peca_normalized.split()
    for fragment in fragments:
        if fragment not in at_text_normalized:
            return False

    return True


def generate_html_report(report, subfolder_name):
    """Gera um relatório em HTML a partir do texto do relatório, separando por subpastas."""
    html_report = f"""
    <div>
        <h2>{subfolder_name}</h2>
        <pre>{report}</pre>
    </div>
    """
    return html_report


def save_and_open_report(html_report):
    """Salva o relatório em HTML e o abre no navegador."""
    with open("../report.html", "w", encoding="utf-8") as file:
        file.write(html_report)
    webbrowser.open("../report.html")


def save_ap_text(ap_text):
    """Salva o texto extraído do AP em um arquivo de texto."""
    with open("../ap_text.txt", "w", encoding="utf-8") as file:
        file.write(ap_text)


def search_text_in_pdf(pdf_path, search_text):
    """Procura o texto diretamente no PDF usando pdfplumber e concatena todas as ocorrências da Razão Social."""
    concatenated_result = ""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            text = "\n".join([page.extract_text() for page in pdf.pages])
        # Normaliza o texto para remover quebras de linha e espaços
        normalized_text = re.sub(r'\s+', '', text.lower())
        normalized_search_text = re.sub(r'\s+', '', search_text.lower())
        if normalized_search_text in normalized_text:
            concatenated_result = search_text  # Caso encontre a string normalizada
    except Exception as e:
        logging.error(f"Erro ao procurar texto no PDF {pdf_path}: {e}")
    return concatenated_result.strip()


def check_fields(os_fields, ap_fields, at_fields_list, sicaf_fields, sicaf_pdf_path, at_files):
    """Executa verificações nos campos extraídos e retorna um relatório."""
    report = ["Transcrição das informações extraídas:", f"OS: {os_fields}", f"AP: {ap_fields}"]
    # Transcreve as informações extraídas
    for i, at_fields in enumerate(at_fields_list):
        report.append(f"AT {i + 1}: {at_fields}")
    report.append(f"SICAF: {sicaf_fields}")

    # CHECK 1
    try:
        report.append(
            f"OS N°                                   CHECK 1.1: {'OK' if os_fields.get('OS N°') == ap_fields.get(
                'OS N°') else 'Non-conformity'}")
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

    # CHECK 2
    for i, at_fields in enumerate(at_fields_list):
        if at_fields.get('AT ') and ap_fields.get('AT DE PRODUCAO'):
            report.append(
                f"AT/AT DE PRODUCAO                       CHECK 2.{i + 1}.1: {'OK' if at_fields['AT '] in ap_fields['AT DE PRODUCAO'] else 'Non-conformity'}")
        else:
            report.append(f"AT/AT DE PRODUCAO                       CHECK 2.{i + 1}.1: Non-conformity")

        if at_fields.get('TITULO'):
            for peca_key in [f'PECA{i + 1}' for i in range(len(ap_fields))]:
                peca = ap_fields.get(peca_key)
                if peca:
                    if check_peca_in_at(at_fields.get('TITULO', ''),
                                        peca):  # Verifica se a peça está no texto extraído pelo OCR
                        report.append(
                            f"PECA/TITULO                             CHECK 2.{i + 1}.2: OK - {peca_key} ({peca})")
                    elif search_peca_in_pdf(at_files[i], peca):  # Se não, busca diretamente no PDF
                        report.append(
                            f"PECA/TITULO                             CHECK 2.{i + 1}.2: OK - {peca_key} ({peca}) encontrada no PDF")
                    else:
                        report.append(
                            f"PECA/TITULO                             CHECK 2.{i + 1}.2: Non-conformity - {peca_key} ({peca}) não encontrada")
        else:
            report.append(f"PECA/TITULO                             CHECK 2.{i + 1}.2: Non-conformity")

        # Verifica os formatos diretamente no PDF da AT
        formatos_ap = [ap_fields.get(f'FORMATO{j + 1}') for j in range(len(ap_fields)) if
                       ap_fields.get(f'FORMATO{j + 1}')]
        for formato_ap in formatos_ap:
            matched = False
            if formato_ap:
                # Verifica se o formato está presente no texto extraído do OCR
                if formato_ap in at_fields.get(f'FORMATO{i + 1}', ''):
                    matched = True
                elif not matched:
                    # Se não encontrou no texto do OCR, verifica diretamente no PDF
                    if search_format_in_pdf(at_files[i], formato_ap):
                        matched = True

            report.append(
                f"FORMATO/FORMATO                         CHECK 2.{i + 1}.3: {'OK' if matched else 'Non-conformity'} - {formato_ap}")
        try:
            ap_data_emissao_dt = datetime.strptime(ap_fields.get('DATA EMISSAO', '01/01/1900'), '%d/%m/%Y')
            at_data_at_dt = datetime.strptime(at_fields.get('Data da AT', '01/01/1900'), '%d/%m/%Y')
            report.append(
                f"DATA EMISSAO/Data da AT                 CHECK 2.{i + 1}.4: {'OK' if ap_data_emissao_dt >= at_data_at_dt else 'Non-conformity'}")
        except Exception as e:
            report.append(f"DATA EMISSAO/Data da AT                 CHECK 2.{i + 1}.4: Error {e}")

    # CHECK 3 (SICAF 1 ou SICAF 2)
    if sicaf_fields.get('SICAF_TYPE') == 'SICAF1':
        report.append(
            f"Razão social                            CHECK 3.1: {'OK' if sicaf_fields.get('Razão social') == ap_fields.get('Razão social') else 'Non-conformity'}")
        report.append(
            f"CNPJ                                    CHECK 3.2: {'OK' if sicaf_fields.get('CNPJ') == ap_fields.get('CNPJ') else 'Non-conformity'}")
        report.append(
            f"Município                               CHECK 3.3: {'OK' if sicaf_fields.get('Município') == ap_fields.get('Município') else 'Non-conformity'}")

    else:  # SICAF 2 - Pesquisa direta no PDF
        # Verificação 3.1: Razão Social
        razao_social_ap = ap_fields.get('Razão social')
        concatenated_razao_social = search_text_in_pdf(sicaf_pdf_path, normalize_razao_social(razao_social_ap))
        if razao_social_ap and concatenated_razao_social:
            report.append(
                "Razão social                            CHECK 3.1: OK - Razão Social do AP encontrada no SICAF.")
        else:
            report.append(
                "Razão social                            CHECK 3.1: Non-conformity - Razão Social do AP não encontrada no SICAF.")

        # Verificação 3.2: CNPJ
        cnpj_ap = ap_fields.get('CNPJ')
        if cnpj_ap and search_text_in_pdf(sicaf_pdf_path, cnpj_ap):
            report.append("CNPJ                                    CHECK 3.2: OK - CNPJ do AP encontrado no SICAF.")
        else:
            report.append(
                "CNPJ                                    CHECK 3.2: Non-conformity - CNPJ do AP não encontrado no SICAF.")

        # Verificação 3.3: Município de AP
        municipio_ap = ap_fields.get('Município')
        if municipio_ap and search_text_in_pdf(sicaf_pdf_path, municipio_ap):
            report.append(
                "Município                               CHECK 3.3: OK - Município do AP encontrado no SICAF.")
        else:
            report.append(
                "Município                               CHECK 3.3: Non-conformity - Município do AP não encontrado no SICAF.")

    # CHECK 4
    orgao = os_fields.get('ORGAO', '')
    produto = ap_fields.get('PRODUTO', '')
    if orgao and produto:
        report.append(
            f"ORGAO/PRODUTO                           CHECK 4.1: {'OK' if any(prod in orgao for prod in produto.split()) else 'Non-conformity'}")
    else:
        report.append("ORGAO/PRODUTO                           CHECK 4.1: Non-conformity")

    report_text = "\n".join(report)
    return report_text


def save_text_to_file(text, file_name, folder_path):
    """Salva o texto extraído em um arquivo de texto."""
    file_path = os.path.join(folder_path, file_name)
    with open(file_path, "w", encoding="utf-8") as file:
        file.write(text)
    logging.info(f"Texto salvo em {file_path}")


def verify_documents(file_paths, subfolder_name):
    """Verifique os documentos realizando a extração de texto e verificando os campos."""
    os_file = file_paths.get('OS')
    ap_file = file_paths.get('AP')
    at_files = file_paths.get('AT', [])
    sicaf_file = file_paths.get('SICAF')

    if not all([os_file, ap_file, sicaf_file]) or not at_files:
        messagebox.showerror("Erro",
                             f"Todos os arquivos (OS, AP, AT, SICAF) devem estar presentes na pasta {subfolder_name}.")
        return ""

    # Diretório para salvar os arquivos de texto
    folder_path = os.path.dirname("../report.html")

    # Extrai o texto para o documento OS e verifica se é do tipo OS2
    os_text = extract_text_with_format_adjustment(os_file)
    determine_os_type(os_text)
    os_fields = extract_fields(os_text, 'OS')

    # Salvar texto de OS
    save_text_to_file(os_text, f"OS_text_{subfolder_name}.txt", folder_path)

    # Extrai o texto para o documento AP
    ap_text = extract_text_with_format_adjustment(ap_file)
    ap_fields = extract_fields(ap_text, 'AP')

    # Salvar texto de AP
    save_text_to_file(ap_text, f"AP_text_{subfolder_name}.txt", folder_path)

    # Extrai o texto para os documentos AT
    at_fields_list = []
    for i, at_file in enumerate(at_files):
        at_text = extract_text_with_format_adjustment(at_file)
        at_fields = extract_fields(at_text, 'AT')
        at_fields_list.append(at_fields)

        # Salvar texto de cada AT
        save_text_to_file(at_text, f"AT_text_{subfolder_name}_{i + 1}.txt", folder_path)

    # Extrai o texto para o documento SICAF
    sicaf_text = extract_text_with_format_adjustment(sicaf_file)
    sicaf_fields = extract_fields(sicaf_text, 'SICAF')

    # Salvar texto de SICAF
    save_text_to_file(sicaf_text, f"SICAF_text_{subfolder_name}.txt", folder_path)

    # Gerar e salvar o relatório
    report = check_fields(os_fields, ap_fields, at_fields_list, sicaf_fields, sicaf_file, at_files)

    return generate_html_report(report, subfolder_name)


def select_folder():
    folder_path = filedialog.askdirectory(title="Selecione a pasta contendo os PDFs")
    if not folder_path:
        messagebox.showerror("Erro", "Pasta não selecionada.")
        return

    # Inicializar conteúdo do relatório HTML
    full_html_report =("<html><head><title>Relatório de Verificação</title></head><body><h1>Relatório de "
                        "Verificação</h1>")

    for root, dirs, files in os.walk(folder_path):
        for subdir in dirs:
            subdir_path = os.path.join(root, subdir)
            file_paths = {'AT': []}

            for file_name in os.listdir(subdir_path):
                if 'OS' in file_name.upper():
                    file_paths['OS'] = os.path.join(subdir_path, file_name)
                elif 'AP' in file_name.upper():
                    file_paths['AP'] = os.path.join(subdir_path, file_name)
                elif 'AT' in file_name.upper():
                    file_paths['AT'].append(os.path.join(subdir_path, file_name))
                elif 'SICAF' in file_name.upper():
                    file_paths['SICAF'] = os.path.join(subdir_path, file_name)

            # Verifica os documentos e adiciona o resultado ao relatório HTML
            subfolder_name = os.path.basename(subdir_path)
            result = verify_documents(file_paths, subfolder_name)
            full_html_report += result

    full_html_report += "</body></html>"

    # Salvar e abrir o relatório HTML completo
    save_and_open_report(full_html_report)


# Interface gráfica
root = tk.Tk()
root.title("Verificação de Documentos PDF")

select_button = tk.Button(root, text="Selecionar Pasta com PDFs", command=select_folder)
select_button.pack(pady=20)

root.mainloop()
