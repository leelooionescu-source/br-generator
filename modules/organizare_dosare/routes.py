import os
import uuid
import shutil
from flask import Blueprint, render_template, request, redirect, url_for, session, send_from_directory, flash
from modules.organizare_dosare.process_organize import (
    parse_borderou, scan_hpv_files, scan_doc_cadastrale,
    build_matching_preview, organize_files, create_output_zip, extract_zip_contents
)

bp = Blueprint('org', __name__, url_prefix='/org', template_folder='../../templates/org')

BASE_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


def get_session_dir(subdir):
    if 'session_id' not in session:
        session['session_id'] = str(uuid.uuid4())
    path = os.path.join(BASE_DIR, subdir, session['session_id'])
    os.makedirs(path, exist_ok=True)
    return path


@bp.route('/')
def index():
    # Check if we have generated BR/BP files from current session
    has_generated = 'generated' in session
    return render_template('org/index.html', has_generated=has_generated)


@bp.route('/upload', methods=['POST'])
def upload():
    upload_dir = get_session_dir('uploads')
    org_dir = os.path.join(upload_dir, 'org')
    if os.path.exists(org_dir):
        shutil.rmtree(org_dir)
    os.makedirs(org_dir, exist_ok=True)

    use_session = request.form.get('use_session_br') == '1'

    # --- Borderourile ---
    borderou_paths = []
    if use_session and 'generated' in session:
        output_dir = get_session_dir('output')
        for g in session['generated']:
            bp_path = os.path.join(output_dir, g['filename'])
            if os.path.exists(bp_path):
                borderou_paths.append(bp_path)
    else:
        br_files = request.files.getlist('borderourile')
        if not br_files or all(f.filename == '' for f in br_files):
            flash('Incarcati cel putin un fisier borderou!', 'error')
            return redirect(url_for('org.index'))
        br_dir = os.path.join(org_dir, 'borderourile')
        os.makedirs(br_dir, exist_ok=True)
        for f in br_files:
            if f.filename:
                path = os.path.join(br_dir, f.filename)
                f.save(path)
                borderou_paths.append(path)

    if not borderou_paths:
        flash('Nu s-au gasit borderourile!', 'error')
        return redirect(url_for('org.index'))

    # --- Hot/PV files ---
    hpv_files = request.files.getlist('hpv_files')
    hpv_zip = request.files.get('hpv_zip')
    hpv_dir = os.path.join(org_dir, 'hpv')
    os.makedirs(hpv_dir, exist_ok=True)
    hpv_paths = []

    if hpv_zip and hpv_zip.filename:
        zip_path = os.path.join(org_dir, 'hpv.zip')
        hpv_zip.save(zip_path)
        files, _ = extract_zip_contents(zip_path, hpv_dir)
        hpv_paths.extend(files)
        os.remove(zip_path)
    elif hpv_files and any(f.filename for f in hpv_files):
        for f in hpv_files:
            if f.filename:
                path = os.path.join(hpv_dir, f.filename)
                f.save(path)
                hpv_paths.append(path)

    # --- Doc Cadastrale ---
    doc_files = request.files.getlist('doc_files')
    doc_zip = request.files.get('doc_zip')
    doc_dir = os.path.join(org_dir, 'doc_cad')
    os.makedirs(doc_dir, exist_ok=True)
    doc_file_paths = []
    doc_folder_paths = []

    if doc_zip and doc_zip.filename:
        zip_path = os.path.join(org_dir, 'doc_cad.zip')
        doc_zip.save(zip_path)
        files, folders = extract_zip_contents(zip_path, doc_dir)
        doc_file_paths.extend(files)
        doc_folder_paths.extend(folders)
        os.remove(zip_path)
    elif doc_files and any(f.filename for f in doc_files):
        for f in doc_files:
            if f.filename:
                path = os.path.join(doc_dir, f.filename)
                f.save(path)
                doc_file_paths.append(path)

    if not hpv_paths and not doc_file_paths and not doc_folder_paths:
        flash('Incarcati cel putin fisiere Hot/PV sau documentatii cadastrale!', 'error')
        return redirect(url_for('org.index'))

    # Save paths in session
    session['org_borderou_paths'] = borderou_paths
    session['org_hpv_paths'] = hpv_paths
    session['org_doc_file_paths'] = doc_file_paths
    session['org_doc_folder_paths'] = doc_folder_paths

    return redirect(url_for('org.preview'))


@bp.route('/preview')
def preview():
    if 'org_borderou_paths' not in session:
        flash('Incarcati fisierele mai intai.', 'error')
        return redirect(url_for('org.index'))

    try:
        borderou_paths = session['org_borderou_paths']
        hpv_paths = session.get('org_hpv_paths', [])
        doc_file_paths = session.get('org_doc_file_paths', [])
        doc_folder_paths = session.get('org_doc_folder_paths', [])

        # Parse all borderourile
        borderou_data_list = [parse_borderou(p) for p in borderou_paths]

        # Build lookups
        hpv_lookup = scan_hpv_files(hpv_paths)
        doc_cad_lookup = scan_doc_cadastrale(doc_file_paths, doc_folder_paths)

        # Build preview
        preview_data = build_matching_preview(borderou_data_list, hpv_lookup, doc_cad_lookup)

        # Stats
        total_hpv_files = len(hpv_lookup)
        total_doc_files = len(doc_cad_lookup)

    except Exception as e:
        flash(f'Eroare la analiza fisierelor: {str(e)}', 'error')
        return redirect(url_for('org.index'))

    return render_template('org/preview.html',
        preview=preview_data,
        total_borderourile=len(borderou_paths),
        total_hpv_files=total_hpv_files,
        total_doc_files=total_doc_files)


@bp.route('/organize', methods=['POST'])
def organize():
    if 'org_borderou_paths' not in session:
        flash('Sesiune expirata. Reincarcati fisierele.', 'error')
        return redirect(url_for('org.index'))

    try:
        borderou_paths = session['org_borderou_paths']
        hpv_paths = session.get('org_hpv_paths', [])
        doc_file_paths = session.get('org_doc_file_paths', [])
        doc_folder_paths = session.get('org_doc_folder_paths', [])

        borderou_data_list = [parse_borderou(p) for p in borderou_paths]
        hpv_lookup = scan_hpv_files(hpv_paths)
        doc_cad_lookup = scan_doc_cadastrale(doc_file_paths, doc_folder_paths)

        output_dir = get_session_dir('output')
        org_output = os.path.join(output_dir, 'dosare')
        if os.path.exists(org_output):
            shutil.rmtree(org_output)
        os.makedirs(org_output, exist_ok=True)

        results = organize_files(borderou_data_list, hpv_lookup, doc_cad_lookup, borderou_paths, org_output)
        session['org_results'] = results
        session['org_output_dir'] = org_output

    except Exception as e:
        flash(f'Eroare la organizare: {str(e)}', 'error')
        return redirect(url_for('org.preview'))

    return redirect(url_for('org.results'))


@bp.route('/results')
def results():
    if 'org_results' not in session:
        flash('Nu exista rezultate.', 'error')
        return redirect(url_for('org.index'))

    results = session['org_results']
    total_folders = len(results)
    total_hpv = sum(r['copied_hpv'] for r in results)
    total_doc = sum(r['copied_doc'] for r in results)
    total_missing_hpv = sum(r['missing_hpv'] for r in results)
    total_missing_doc = sum(r['missing_doc'] for r in results)

    return render_template('org/results.html',
        results=results,
        total_folders=total_folders,
        total_hpv=total_hpv,
        total_doc=total_doc,
        total_missing_hpv=total_missing_hpv,
        total_missing_doc=total_missing_doc)


@bp.route('/download-all')
def download_all():
    if 'org_output_dir' not in session:
        flash('Nu exista fisiere de descarcat.', 'error')
        return redirect(url_for('org.index'))

    org_output = session['org_output_dir']
    output_dir = os.path.dirname(org_output)
    zip_name = 'Dosare_organizate.zip'
    create_output_zip(org_output, zip_name)
    return send_from_directory(org_output, zip_name, as_attachment=True)
