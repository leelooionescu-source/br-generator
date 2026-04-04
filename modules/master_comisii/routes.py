import os
import uuid
import shutil
import zipfile
from flask import Blueprint, render_template, request, redirect, url_for, session, send_from_directory, flash
from modules.master_comisii.process_master import (
    detect_format, preview_import, import_to_master,
    get_uats_from_master, get_master_stats, update_cols_30_38,
    preview_merge, generate_word_merge, convert_doc_to_docx
)

bp = Blueprint('master', __name__, url_prefix='/master', template_folder='../../templates/master')

BASE_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


def get_session_dir(subdir):
    if 'session_id' not in session:
        session['session_id'] = str(uuid.uuid4())
    path = os.path.join(BASE_DIR, subdir, session['session_id'])
    os.makedirs(path, exist_ok=True)
    return path


@bp.route('/')
def index():
    return render_template('master/index.html')


@bp.route('/upload', methods=['POST'])
def upload():
    hg_number = request.form.get('hg_number', '').strip()
    if not hg_number:
        flash('Introduceti numarul HG!', 'error')
        return redirect(url_for('master.index'))

    master_file = request.files.get('master')
    sursa_file = request.files.get('sursa')
    if not master_file or master_file.filename == '':
        flash('Fisierul MASTER este obligatoriu!', 'error')
        return redirect(url_for('master.index'))
    if not sursa_file or sursa_file.filename == '':
        flash('Fisierul sursa date este obligatoriu!', 'error')
        return redirect(url_for('master.index'))

    upload_dir = get_session_dir('uploads')
    # Clear previous
    for f in os.listdir(upload_dir):
        fp = os.path.join(upload_dir, f)
        if os.path.isfile(fp):
            os.remove(fp)

    master_path = os.path.join(upload_dir, 'master.xlsx')
    sursa_path = os.path.join(upload_dir, 'sursa.xlsx')
    master_file.save(master_path)
    sursa_file.save(sursa_path)

    data_sedinta = request.form.get('data_sedinta', '').strip()

    session['master_hg_number'] = hg_number
    session['master_data_sedinta'] = data_sedinta
    session['master_files'] = {
        'master': master_path,
        'sursa': sursa_path,
    }

    return redirect(url_for('master.step1_preview'))


@bp.route('/step1')
def step1_preview():
    if 'master_files' not in session:
        flash('Incarcati fisierele mai intai.', 'error')
        return redirect(url_for('master.index'))

    paths = session['master_files']
    hg_number = session.get('master_hg_number', '')

    try:
        fmt = detect_format(paths['sursa'])
        stats = preview_import(paths['sursa'], fmt)
        session['master_format'] = fmt
    except Exception as e:
        flash(f'Eroare la analiza fisierului sursa: {str(e)}', 'error')
        return redirect(url_for('master.index'))

    return render_template('master/step1_preview.html',
        hg_number=hg_number, stats=stats)


@bp.route('/do-import', methods=['POST'])
def do_import():
    if 'master_files' not in session:
        flash('Sesiune expirata. Reincarcati fisierele.', 'error')
        return redirect(url_for('master.index'))

    paths = session['master_files']
    hg_number = session.get('master_hg_number', '')
    data_sedinta = session.get('master_data_sedinta', '')
    fmt = session.get('master_format', 'A')

    upload_dir = get_session_dir('uploads')
    updated_master = os.path.join(upload_dir, 'master_updated.xlsx')

    try:
        result = import_to_master(
            paths['master'], paths['sursa'], fmt,
            hg_number, data_sedinta, updated_master
        )
        session['master_updated'] = updated_master
        session['master_import_result'] = result
    except Exception as e:
        flash(f'Eroare la import: {str(e)}', 'error')
        return redirect(url_for('master.step1_preview'))

    return redirect(url_for('master.step2'))


@bp.route('/step2')
def step2():
    master_path = session.get('master_updated') or (session.get('master_files', {}).get('master'))
    if not master_path:
        flash('Incarcati fisierele mai intai.', 'error')
        return redirect(url_for('master.index'))

    try:
        uats = get_uats_from_master(master_path)
        master_stats = get_master_stats(master_path)
    except Exception as e:
        flash(f'Eroare la citirea MASTER: {str(e)}', 'error')
        return redirect(url_for('master.index'))

    import_result = session.pop('master_import_result', None)

    # Prefill from session if available
    prefill = session.get('master_prefill', {})

    return render_template('master/step2_config.html',
        uats=uats, master_stats=master_stats,
        import_result=import_result, prefill=prefill)


@bp.route('/do-update-cols', methods=['POST'])
def do_update_cols():
    master_path = session.get('master_updated') or (session.get('master_files', {}).get('master'))
    if not master_path:
        flash('Sesiune expirata.', 'error')
        return redirect(url_for('master.index'))

    # Build config from form
    config = {
        'fixed': {
            30: request.form.get('col_30', '').strip(),
            31: request.form.get('col_31', '').strip(),
            32: request.form.get('col_32', '').strip(),
        },
        'per_uat': {}
    }

    # Save prefill for back navigation
    session['master_prefill'] = {
        'col_30': config['fixed'][30],
        'col_31': config['fixed'][31],
        'col_32': config['fixed'][32],
    }

    # Collect per-UAT values
    i = 0
    while True:
        uat_name = request.form.get(f'uat_{i}_name')
        if uat_name is None:
            break
        uat_vals = {}
        for col in range(33, 39):
            val = request.form.get(f'uat_{i}_col_{col}', '').strip()
            if val:
                uat_vals[col] = val
        col_33 = request.form.get(f'uat_{i}_col_33', '').strip()
        if col_33:
            uat_vals[33] = col_33
        if uat_vals:
            config['per_uat'][uat_name] = uat_vals
        i += 1

    upload_dir = get_session_dir('uploads')
    output_path = os.path.join(upload_dir, 'master_cols_updated.xlsx')

    try:
        count = update_cols_30_38(master_path, config, output_path)
        session['master_updated'] = output_path
        session['master_update_count'] = count
    except Exception as e:
        flash(f'Eroare la actualizare coloane: {str(e)}', 'error')
        return redirect(url_for('master.step2'))

    return redirect(url_for('master.step3'))


@bp.route('/step3')
def step3():
    master_path = session.get('master_updated') or (session.get('master_files', {}).get('master'))
    if not master_path:
        flash('Incarcati fisierele mai intai.', 'error')
        return redirect(url_for('master.index'))

    try:
        merge_prev = preview_merge(master_path)
    except Exception as e:
        flash(f'Eroare la citirea MASTER: {str(e)}', 'error')
        return redirect(url_for('master.index'))

    update_count = session.pop('master_update_count', None)

    return render_template('master/step3_preview.html',
        merge_preview=merge_prev, update_count=update_count)


@bp.route('/do-generate', methods=['POST'])
def do_generate():
    master_path = session.get('master_updated') or (session.get('master_files', {}).get('master'))
    if not master_path:
        flash('Sesiune expirata.', 'error')
        return redirect(url_for('master.index'))

    template_file = request.files.get('template_word')
    if not template_file or template_file.filename == '':
        flash('Fisierul sablon Word este obligatoriu!', 'error')
        return redirect(url_for('master.step3'))

    output_dir = get_session_dir('output')
    # Clear previous output
    for f in os.listdir(output_dir):
        fp = os.path.join(output_dir, f)
        if os.path.isfile(fp):
            os.remove(fp)
        elif os.path.isdir(fp):
            shutil.rmtree(fp)

    upload_dir = get_session_dir('uploads')
    filename = template_file.filename
    ext = os.path.splitext(filename)[1].lower()
    template_path = os.path.join(upload_dir, f'template_word{ext}')
    template_file.save(template_path)

    # Convert .doc to .docx if needed
    if ext == '.doc':
        try:
            template_path = convert_doc_to_docx(template_path, upload_dir)
        except Exception as e:
            flash(f'Eroare la conversia .doc: {str(e)}', 'error')
            return redirect(url_for('master.step3'))

    try:
        word_result = generate_word_merge(master_path, template_path, output_dir)
        session['master_word_result'] = word_result
    except Exception as e:
        flash(f'Eroare la generare Word: {str(e)}', 'error')
        return redirect(url_for('master.step3'))

    return redirect(url_for('master.results'))


@bp.route('/results')
def results():
    master_path = session.get('master_updated') or (session.get('master_files', {}).get('master'))
    if not master_path:
        flash('Nu exista date de afisat.', 'error')
        return redirect(url_for('master.index'))

    word_result = session.get('master_word_result')
    hg_number = session.get('master_hg_number', '')

    return render_template('master/results.html',
        hg_number=hg_number, word_result=word_result)


@bp.route('/download-master')
def download_master():
    master_path = session.get('master_updated')
    if not master_path or not os.path.exists(master_path):
        flash('Nu exista MASTER actualizat.', 'error')
        return redirect(url_for('master.index'))
    hg_clean = session.get('master_hg_number', 'HG').replace('/', '-')
    return send_from_directory(
        os.path.dirname(master_path), os.path.basename(master_path),
        as_attachment=True,
        download_name=f'MASTER actualizat HG {hg_clean}.xlsx')


@bp.route('/download/<filename>')
def download_file(filename):
    output_dir = get_session_dir('output')
    return send_from_directory(output_dir, filename, as_attachment=True)


@bp.route('/download-all')
def download_all():
    output_dir = get_session_dir('output')
    master_path = session.get('master_updated')
    hg_clean = session.get('master_hg_number', 'HG').replace('/', '-')
    zip_path = os.path.join(output_dir, f'MASTER_COMISII_HG_{hg_clean}.zip')

    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        # Add MASTER
        if master_path and os.path.exists(master_path):
            zf.write(master_path, f'MASTER actualizat HG {hg_clean}.xlsx')
        # Add Word files
        for f in os.listdir(output_dir):
            if f.endswith('.docx'):
                zf.write(os.path.join(output_dir, f), f)

    return send_from_directory(output_dir, os.path.basename(zip_path), as_attachment=True)
