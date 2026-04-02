import os
import uuid
import shutil
import zipfile
import time
from flask import Blueprint, render_template, request, redirect, url_for, session, send_from_directory, flash
from modules.br_generator.process_br import analyze_master, update_situatie_col_n, analyze_situatie, generate_all_br, parse_recipise

bp = Blueprint('br', __name__, url_prefix='/br', template_folder='../../templates/br')

BASE_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
UPLOADS_DIR = os.path.join(BASE_DIR, 'uploads')
OUTPUT_DIR = os.path.join(BASE_DIR, 'output')


def get_session_dir(subdir):
    if 'session_id' not in session:
        session['session_id'] = str(uuid.uuid4())
    path = os.path.join(BASE_DIR, subdir, session['session_id'])
    os.makedirs(path, exist_ok=True)
    return path


@bp.route('/')
def index():
    return render_template('br/index.html')


@bp.route('/upload', methods=['POST'])
def upload():
    hg_number = request.form.get('hg_number', '').strip()
    if not hg_number:
        flash('Introduceti numarul HG!', 'error')
        return redirect(url_for('br.index'))

    required = {
        'master': ('MASTER', request.files.get('master')),
        'situatie': ('SITUATIE', request.files.get('situatie')),
        'template_br1': ('Template BR nr. 1', request.files.get('template_br1')),
        'template_br11': ('Template BR nr. 1.1', request.files.get('template_br11')),
    }
    for name, (label, f) in required.items():
        if not f or f.filename == '':
            flash(f'Fisierul {label} este obligatoriu!', 'error')
            return redirect(url_for('br.index'))

    upload_dir = get_session_dir('uploads')
    for f in os.listdir(upload_dir):
        os.remove(os.path.join(upload_dir, f))

    paths = {}
    for name, (label, f) in required.items():
        path = os.path.join(upload_dir, f'{name}.xlsx')
        f.save(path)
        paths[name] = path

    for opt_name in ['recipise', 'template_bp']:
        opt_file = request.files.get(opt_name)
        if opt_file and opt_file.filename != '':
            opt_path = os.path.join(upload_dir, f'{opt_name}.xlsx')
            opt_file.save(opt_path)
            paths[opt_name] = opt_path

    session['hg_number'] = hg_number
    session['files'] = paths
    return redirect(url_for('br.preview'))


@bp.route('/preview')
def preview():
    if 'files' not in session:
        flash('Incarcati fisierele mai intai.', 'error')
        return redirect(url_for('br.index'))

    paths = session['files']
    hg_number = session.get('hg_number', '')

    try:
        master_data = analyze_master(paths['master'])
        master_stats = {name: len(entries) for name, entries in master_data.items()}

        tip_counts = {'PLATA': 0, 'LCA': 0, 'CONSEMNARE': 0, 'ALTELE': 0}
        for entries in master_data.values():
            for e in entries:
                tip = e.get('tip_hsd', '')
                if tip in tip_counts:
                    tip_counts[tip] += 1
                elif tip:
                    tip_counts['ALTELE'] += 1

        upload_dir = get_session_dir('uploads')
        updated_situatie = os.path.join(upload_dir, 'situatie_updated.xlsx')
        update_result = update_situatie_col_n(paths['situatie'], master_data, updated_situatie)
        session['updated_situatie'] = updated_situatie

        stats = analyze_situatie(updated_situatie)

        recipise_stats = None
        if 'recipise' in paths:
            recipise_lookup = parse_recipise(paths['recipise'])
            recipise_stats = len(recipise_lookup)
    except Exception as e:
        flash(f'Eroare la analiza fisierelor: {str(e)}', 'error')
        return redirect(url_for('br.index'))

    has_bp_template = 'template_bp' in paths

    return render_template('br/preview.html',
        hg_number=hg_number, master_stats=master_stats,
        update_result=update_result, stats=stats,
        recipise_stats=recipise_stats, tip_counts=tip_counts,
        has_bp_template=has_bp_template)


@bp.route('/generate', methods=['POST'])
def generate():
    if 'files' not in session or 'updated_situatie' not in session:
        flash('Sesiune expirata. Reincarcati fisierele.', 'error')
        return redirect(url_for('br.index'))

    paths = session['files']
    hg_number = session.get('hg_number', '')
    updated_situatie = session['updated_situatie']

    output_dir = get_session_dir('output')
    for f in os.listdir(output_dir):
        fp = os.path.join(output_dir, f)
        if os.path.isfile(fp):
            os.remove(fp)

    try:
        recipise_lookup = {}
        if 'recipise' in paths:
            recipise_lookup = parse_recipise(paths['recipise'])

        master_data = analyze_master(paths['master'])

        generated = generate_all_br(
            updated_situatie, paths['template_br1'], paths['template_br11'],
            hg_number, output_dir, recipise_lookup=recipise_lookup,
            master_data=master_data, template_bp=paths.get('template_bp'))
        session['generated'] = generated
    except Exception as e:
        flash(f'Eroare la generare: {str(e)}', 'error')
        return redirect(url_for('br.preview'))

    return redirect(url_for('br.results'))


@bp.route('/results')
def results():
    if 'generated' not in session:
        flash('Nu exista fisiere generate.', 'error')
        return redirect(url_for('br.index'))

    generated = session['generated']
    hg_number = session.get('hg_number', '')
    total_match = sum(g['count'] for g in generated if g['br_type'] == 'BR nr. 1')
    total_mismatch = sum(g['count'] for g in generated if g['br_type'] == 'BR nr. 1.1')
    total_bp = sum(g['count'] for g in generated if g['br_type'] == 'BP nr. 1')

    return render_template('br/results.html',
        hg_number=hg_number, generated=generated,
        total_match=total_match, total_mismatch=total_mismatch, total_bp=total_bp)


@bp.route('/download/<filename>')
def download(filename):
    output_dir = get_session_dir('output')
    return send_from_directory(output_dir, filename, as_attachment=True)


@bp.route('/download-situatie')
def download_situatie():
    if 'updated_situatie' not in session:
        flash('Nu exista SITUATIE actualizata.', 'error')
        return redirect(url_for('br.index'))
    path = session['updated_situatie']
    return send_from_directory(os.path.dirname(path), os.path.basename(path),
                               as_attachment=True,
                               download_name=f'SITUATIE actualizata HG {session.get("hg_number","")}.xlsx')


@bp.route('/download-all')
def download_all():
    output_dir = get_session_dir('output')
    hg_clean = session.get('hg_number', 'HG').replace('/', '-')
    zip_path = os.path.join(output_dir, f'BR_toate_HG_{hg_clean}.zip')
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for f in os.listdir(output_dir):
            if f.endswith('.xlsx'):
                zf.write(os.path.join(output_dir, f), f)
    return send_from_directory(output_dir, os.path.basename(zip_path), as_attachment=True)
