import os
import uuid
import shutil
import zipfile
import time
from flask import Flask, render_template, request, redirect, url_for, session, send_from_directory, flash
from process_br import analyze_master, update_situatie_col_n, analyze_situatie, generate_all_br, parse_recipise

app = Flask(__name__)
app.secret_key = 'br-generator-secret-key-2026'
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOADS_DIR = os.path.join(BASE_DIR, 'uploads')
OUTPUT_DIR = os.path.join(BASE_DIR, 'output')
os.makedirs(UPLOADS_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)


def get_session_dir(subdir):
    if 'session_id' not in session:
        session['session_id'] = str(uuid.uuid4())
    path = os.path.join(BASE_DIR, subdir, session['session_id'])
    os.makedirs(path, exist_ok=True)
    return path


def cleanup_old_sessions():
    """Sterge sesiuni mai vechi de 24h."""
    cutoff = time.time() - 86400
    for base in [UPLOADS_DIR, OUTPUT_DIR]:
        if not os.path.exists(base):
            continue
        for d in os.listdir(base):
            path = os.path.join(base, d)
            if os.path.isdir(path) and os.path.getmtime(path) < cutoff:
                shutil.rmtree(path, ignore_errors=True)


@app.route('/')
def index():
    cleanup_old_sessions()
    return render_template('index.html')


@app.route('/upload', methods=['POST'])
def upload():
    hg_number = request.form.get('hg_number', '').strip()
    if not hg_number:
        flash('Introduceti numarul HG!', 'error')
        return redirect(url_for('index'))

    files = {
        'master': request.files.get('master'),
        'situatie': request.files.get('situatie'),
        'template_br1': request.files.get('template_br1'),
        'template_br11': request.files.get('template_br11'),
    }

    for name, f in files.items():
        if not f or f.filename == '':
            labels = {'master': 'MASTER', 'situatie': 'SITUATIE', 'template_br1': 'Template BR nr. 1', 'template_br11': 'Template BR nr. 1.1'}
            flash(f'Fisierul {labels[name]} este obligatoriu!', 'error')
            return redirect(url_for('index'))

    upload_dir = get_session_dir('uploads')
    # Clear previous uploads
    for f in os.listdir(upload_dir):
        os.remove(os.path.join(upload_dir, f))

    paths = {}
    for name, f in files.items():
        path = os.path.join(upload_dir, f'{name}.xlsx')
        f.save(path)
        paths[name] = path

    # Optional RECIPISE file
    recipise_file = request.files.get('recipise')
    if recipise_file and recipise_file.filename != '':
        rec_path = os.path.join(upload_dir, 'recipise.xlsx')
        recipise_file.save(rec_path)
        paths['recipise'] = rec_path

    session['hg_number'] = hg_number
    session['files'] = paths
    return redirect(url_for('preview'))


@app.route('/preview')
def preview():
    if 'files' not in session:
        flash('Incarcati fisierele mai intai.', 'error')
        return redirect(url_for('index'))

    paths = session['files']
    hg_number = session.get('hg_number', '')

    try:
        master_data = analyze_master(paths['master'])
        master_stats = {name: len(entries) for name, entries in master_data.items()}

        # Update SITUATIE col N from MASTER
        upload_dir = get_session_dir('uploads')
        updated_situatie = os.path.join(upload_dir, 'situatie_updated.xlsx')
        update_result = update_situatie_col_n(paths['situatie'], master_data, updated_situatie)
        session['updated_situatie'] = updated_situatie

        # Analyze updated SITUATIE
        stats = analyze_situatie(updated_situatie)

        # Parse RECIPISE if provided
        recipise_stats = None
        if 'recipise' in paths:
            recipise_lookup = parse_recipise(paths['recipise'])
            recipise_stats = len(recipise_lookup)
    except Exception as e:
        flash(f'Eroare la analiza fisierelor: {str(e)}', 'error')
        return redirect(url_for('index'))

    return render_template('preview.html',
        hg_number=hg_number,
        master_stats=master_stats,
        update_result=update_result,
        stats=stats,
        recipise_stats=recipise_stats)


@app.route('/generate', methods=['POST'])
def generate():
    if 'files' not in session or 'updated_situatie' not in session:
        flash('Sesiune expirata. Reincarcati fisierele.', 'error')
        return redirect(url_for('index'))

    paths = session['files']
    hg_number = session.get('hg_number', '')
    updated_situatie = session['updated_situatie']

    output_dir = get_session_dir('output')
    # Clear previous output
    for f in os.listdir(output_dir):
        fp = os.path.join(output_dir, f)
        if os.path.isfile(fp):
            os.remove(fp)

    try:
        # Parse RECIPISE if provided
        recipise_lookup = {}
        if 'recipise' in paths:
            recipise_lookup = parse_recipise(paths['recipise'])

        generated = generate_all_br(
            updated_situatie,
            paths['template_br1'],
            paths['template_br11'],
            hg_number,
            output_dir,
            recipise_lookup=recipise_lookup
        )
        session['generated'] = generated
    except Exception as e:
        flash(f'Eroare la generare: {str(e)}', 'error')
        return redirect(url_for('preview'))

    return redirect(url_for('results'))


@app.route('/results')
def results():
    if 'generated' not in session:
        flash('Nu exista fisiere generate.', 'error')
        return redirect(url_for('index'))

    generated = session['generated']
    hg_number = session.get('hg_number', '')
    total_match = sum(g['count'] for g in generated if g['br_type'] == 'BR nr. 1')
    total_mismatch = sum(g['count'] for g in generated if g['br_type'] == 'BR nr. 1.1')

    return render_template('results.html',
        hg_number=hg_number,
        generated=generated,
        total_match=total_match,
        total_mismatch=total_mismatch)


@app.route('/download/<filename>')
def download(filename):
    output_dir = get_session_dir('output')
    return send_from_directory(output_dir, filename, as_attachment=True)


@app.route('/download-situatie')
def download_situatie():
    if 'updated_situatie' not in session:
        flash('Nu exista SITUATIE actualizata.', 'error')
        return redirect(url_for('index'))
    path = session['updated_situatie']
    return send_from_directory(os.path.dirname(path), os.path.basename(path),
                               as_attachment=True,
                               download_name=f'SITUATIE actualizata HG {session.get("hg_number","")}.xlsx')


@app.route('/download-all')
def download_all():
    output_dir = get_session_dir('output')
    hg_clean = session.get('hg_number', 'HG').replace('/', '-')
    zip_path = os.path.join(output_dir, f'BR_toate_HG_{hg_clean}.zip')

    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for f in os.listdir(output_dir):
            if f.endswith('.xlsx'):
                zf.write(os.path.join(output_dir, f), f)

    return send_from_directory(output_dir, os.path.basename(zip_path), as_attachment=True)


if __name__ == '__main__':
    print('\n' + '=' * 50)
    print('  Generator BR - Borderou de Reconsemnare')
    print('=' * 50)
    print(f'  Local:  http://localhost:5050')
    print(f'  Retea:  http://<IP-ul-PC-ului>:5050')
    print('=' * 50 + '\n')
    app.run(host='0.0.0.0', port=5050, debug=False)
