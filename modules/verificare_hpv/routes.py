import os
import uuid
import shutil
import time
import threading
from flask import Blueprint, render_template, request, redirect, url_for, session, send_from_directory, flash, jsonify
from modules.verificare_hpv.process_verify import read_master, process_all, generate_report

bp = Blueprint('hpv', __name__, url_prefix='/hpv', template_folder='../../templates/hpv')

BASE_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
UPLOADS_DIR = os.path.join(BASE_DIR, 'uploads')
OUTPUT_DIR = os.path.join(BASE_DIR, 'output')

progress_store = {}
results_store = {}


def get_session_dir(subdir):
    if 'session_id' not in session:
        session['session_id'] = str(uuid.uuid4())
    path = os.path.join(BASE_DIR, subdir, session['session_id'])
    os.makedirs(path, exist_ok=True)
    return path


@bp.route('/')
def index():
    return render_template('hpv/index.html')


@bp.route('/upload', methods=['POST'])
def upload():
    master_file = request.files.get('master')
    if not master_file or master_file.filename == '':
        flash('Fisierul MASTER este obligatoriu!', 'error')
        return redirect(url_for('hpv.index'))

    doc_files = request.files.getlist('docs')
    if not doc_files or all(f.filename == '' for f in doc_files):
        flash('Incarcati cel putin un fisier Word (.docx)!', 'error')
        return redirect(url_for('hpv.index'))

    upload_dir = get_session_dir('uploads')
    shutil.rmtree(upload_dir, ignore_errors=True)
    os.makedirs(upload_dir, exist_ok=True)

    docs_dir = os.path.join(upload_dir, 'docs')
    os.makedirs(docs_dir, exist_ok=True)

    master_path = os.path.join(upload_dir, 'master.xlsx')
    master_file.save(master_path)

    doc_count = 0
    doc_paths = []
    for f in doc_files:
        if f.filename and f.filename.lower().endswith('.docx'):
            fpath = os.path.join(docs_dir, f.filename)
            f.save(fpath)
            doc_paths.append(fpath)
            doc_count += 1

    session['hpv_master_path'] = master_path
    session['hpv_doc_paths'] = doc_paths
    session['hpv_doc_count'] = doc_count

    return redirect(url_for('hpv.process'))


@bp.route('/process')
def process():
    if 'hpv_master_path' not in session:
        flash('Incarcati fisierele mai intai.', 'error')
        return redirect(url_for('hpv.index'))
    return render_template('hpv/processing.html', doc_count=session.get('hpv_doc_count', 0))


@bp.route('/run-process', methods=['POST'])
def run_process():
    if 'hpv_master_path' not in session:
        return jsonify({'error': 'No files'}), 400

    master_path = session['hpv_master_path']
    doc_paths = session['hpv_doc_paths']
    sid = session.get('session_id', '')
    output_dir_path = get_session_dir('output')

    progress_store[sid] = {'current': 0, 'total': 0, 'file': 'Pornire...', 'done': False}
    results_store[sid] = None

    def background_process():
        def progress_cb(current, total, filename):
            progress_store[sid] = {'current': current, 'total': total, 'file': filename, 'done': False}

        try:
            results, master_entries = process_all(master_path, doc_paths, progress_callback=progress_cb)
            report_path = os.path.join(output_dir_path, 'Raport neconcordante.xlsx')
            generate_report(results, report_path)

            results_store[sid] = {
                'results': results,
                'report_path': report_path,
                'total_ok': sum(1 for r in results if not r['issues']),
                'total_issues': sum(1 for r in results if r['issues']),
                'total_neconcordante': sum(len(r['issues']) for r in results),
            }
            progress_store[sid] = {'current': len(results), 'total': len(results), 'file': 'Finalizat', 'done': True}
        except Exception as e:
            progress_store[sid] = {'current': 0, 'total': 0, 'file': str(e), 'done': True, 'error': str(e)}

    thread = threading.Thread(target=background_process)
    thread.start()
    return jsonify({'ok': True})


@bp.route('/progress')
def progress():
    sid = session.get('session_id', '')
    p = progress_store.get(sid, {'current': 0, 'total': 0, 'file': '', 'done': False})
    return jsonify(p)


@bp.route('/results')
def results():
    sid = session.get('session_id', '')
    data = results_store.get(sid)
    if not data:
        flash('Nu exista rezultate.', 'error')
        return redirect(url_for('hpv.index'))

    session['hpv_report_path'] = data['report_path']

    return render_template('hpv/results.html',
        results=data['results'], total_ok=data['total_ok'],
        total_issues=data['total_issues'], total_neconcordante=data['total_neconcordante'])


@bp.route('/download-report')
def download_report():
    if 'hpv_report_path' not in session:
        flash('Nu exista raport.', 'error')
        return redirect(url_for('hpv.index'))
    path = session['hpv_report_path']
    return send_from_directory(os.path.dirname(path), os.path.basename(path), as_attachment=True)
