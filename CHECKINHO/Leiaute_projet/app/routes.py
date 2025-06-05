# routes.py

import os
import logging
import shutil
import threading
from flask import Blueprint, render_template, request
from flask_login import login_required, current_user
from html import escape
from db import get_db_connection
from services import (
    delete_temp_folder,
    verify_documents,
    move_relatorios_folder,
    allowed_file
)

bp = Blueprint('main', __name__)

@bp.route('/', methods=['GET', 'POST'])
@login_required
def upload_files():
    if request.method == 'POST':
        if 'files' not in request.files:
            error_message = "Faltando arquivos do diretório."
            return render_template('error.html', error_message=error_message), 400

        temp_pdf_dir = os.path.join(os.path.dirname(__file__), 'temp_pdf')

        if os.path.exists(temp_pdf_dir):
            try:
                shutil.rmtree(temp_pdf_dir)
                logging.info(f"Pasta temporária {temp_pdf_dir} apagada no início da execução.")
            except Exception as e:
                logging.error(f"Erro ao tentar apagar a pasta temporária {temp_pdf_dir}: {e}")

        os.makedirs(temp_pdf_dir, exist_ok=True)

        # Inicia a thread para apagar a pasta temp_pdf após 45 segundos
        threading.Thread(target=delete_temp_folder).start()

        selected_fields = request.form.getlist('fields')
        files = request.files.getlist('files')
        for file in files:
            file_path = os.path.join(temp_pdf_dir, file.filename)
            os.makedirs(os.path.dirname(file_path), exist_ok=True)
            file.save(file_path)

        full_html_report = ""
        campanha_dirs = []

        total_subfolders_sent = 0
        total_subfolders_processed = 0
        total_subfolders_ignored = 0
        ignored_subfolders = []
        root_folder_name = ""

        ok_processes = []
        nc_processes = []

        immediate_subdirs = [
            d for d in os.listdir(temp_pdf_dir)
            if os.path.isdir(os.path.join(temp_pdf_dir, d))
        ]
        if immediate_subdirs:
            root_folder_name = immediate_subdirs[0]

        # Percorre subpastas
        for root, dirs, files in os.walk(temp_pdf_dir):
            if root == temp_pdf_dir:
                continue

            # Se for uma pasta chamada "campanha"
            if 'campanha' in os.path.basename(root).lower():
                campanha_dir = root
                campanha_dirs.append(campanha_dir)

                at_files = []
                os_file = None

                for file_name in os.listdir(campanha_dir):
                    file_path = os.path.join(campanha_dir, file_name)
                    if os.path.isfile(file_path):
                        if 'AT' in file_name.upper():
                            at_files.append(file_path)
                        elif 'OS' in file_name.upper():
                            os_file = file_path

                subdirs = [
                    d for d in os.listdir(campanha_dir)
                    if os.path.isdir(os.path.join(campanha_dir, d))
                ]
                total_subfolders_sent += len(subdirs)

                for subdir in subdirs:
                    subdir_path = os.path.join(campanha_dir, subdir)
                    sicaf_file = None
                    ap_file = None

                    for file_name in os.listdir(subdir_path):
                        if 'SICAF' in file_name.upper():
                            sicaf_file = os.path.join(subdir_path, file_name)
                        elif 'AP' in file_name.upper():
                            ap_file = os.path.join(subdir_path, file_name)

                    file_paths = {
                        'OS': os_file,
                        'AT': at_files,
                        'SICAF': sicaf_file,
                        'AP': ap_file
                    }
                    subfolder_name = subdir

                    result, status, error_message = verify_documents(
                        file_paths,
                        subfolder_name,
                        temp_pdf_dir,
                        selected_fields
                    )
                    if error_message:
                        logging.warning(f"Erro em '{subfolder_name}': {error_message}")
                    else:
                        full_html_report += result
                        total_subfolders_processed += 1

                        if status == 'OK':
                            ok_processes.append(subfolder_name)
                        elif status == 'NC':
                            nc_processes.append(subfolder_name)

                continue  # Próximo "root"

            # Se não for pasta "campanha" nem subpasta dela
            elif (root != temp_pdf_dir and
                  not any(root.startswith(campanha_dir + os.sep) for campanha_dir in campanha_dirs)):

                file_paths = {'AT': []}
                for file_name in files:
                    if 'OS' in file_name.upper():
                        file_paths['OS'] = os.path.join(root, file_name)
                    elif 'AP' in file_name.upper():
                        file_paths['AP'] = os.path.join(root, file_name)
                    elif 'SICAF' in file_name.upper():
                        file_paths['SICAF'] = os.path.join(root, file_name)
                    elif 'AT' in file_name.upper():
                        file_paths['AT'].append(os.path.join(root, file_name))

                subfolder_name = os.path.basename(root)
                total_subfolders_sent += 1

                if 'AP' not in file_paths or not file_paths['AP']:
                    total_subfolders_ignored += 1
                    ignored_subfolders.append(subfolder_name)
                    continue

                result, status, error_message = verify_documents(
                    file_paths,
                    subfolder_name,
                    temp_pdf_dir,
                    selected_fields
                )
                if error_message:
                    full_html_report += f"<h2>Erro no conjunto {escape(subfolder_name)}</h2><p>{escape(error_message)}</p>"
                else:
                    full_html_report += result
                    total_subfolders_processed += 1
                    if status == 'OK':
                        ok_processes.append(subfolder_name)
                    elif status == 'NC':
                        nc_processes.append(subfolder_name)

        summary_report = f"""
        <div class="summary">
            <h2>Resumo do Processamento</h2>
            <p>Total de subpastas enviadas: <strong>{total_subfolders_sent}</strong></p>
            <p>Total de subpastas processadas: <strong>{total_subfolders_processed}</strong></p>
            <p>Total de subpastas ignoradas (sem AP válido): <strong>{total_subfolders_ignored}</strong></p>
        """

        if total_subfolders_ignored > 0:
            summary_report += "<p>Subpastas Ignoradas:</p><ul>"
            for ignored in ignored_subfolders:
                summary_report += f"<li>{escape(ignored)}</li>"
            summary_report += "</ul>"

        if root_folder_name:
            summary_report += f"<p>Pasta Raiz: <strong>{escape(root_folder_name)}</strong></p>"

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
        full_html_report += summary_report

        conn = get_db_connection()
        conn.execute(
            'INSERT INTO results (user_id, subfolder_name, report) VALUES (?, ?, ?)',
            (current_user.id, root_folder_name, full_html_report)
        )
        conn.commit()
        conn.close()

        # Move a pasta de relatórios
        destination_path = os.environ.get('OUTPUT_PATH', r"G:\\Shared drives\\AUTOMACAO\\CHECKIN_MIDIA")
        move_relatorios_folder(destination_path)

        return render_template('report.html', report_content=full_html_report)

    # Se GET, apenas exibe a página de upload
    return render_template('upload.html')


@bp.route('/history')
@login_required
def history():
    conn = get_db_connection()
    results = conn.execute(
        'SELECT id, subfolder_name, created_at FROM results WHERE user_id = ? ORDER BY id DESC',
        (current_user.id,)
    ).fetchall()
    conn.close()
    return render_template('history.html', results=results)


@bp.route('/result/<int:result_id>')
@login_required
def view_result(result_id):
    conn = get_db_connection()
    result = conn.execute(
        'SELECT report FROM results WHERE id = ? AND user_id = ?',
        (result_id, current_user.id)
    ).fetchone()
    conn.close()
    if result:
        return render_template('report.html', report_content=result['report'])
    return 'Resultado não encontrado', 404

