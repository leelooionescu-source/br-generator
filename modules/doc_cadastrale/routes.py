import os
import uuid
import shutil
import json
import logging
from flask import Blueprint, render_template, request, redirect, url_for, session, send_from_directory, flash

logger = logging.getLogger(__name__)
from modules.doc_cadastrale.process_doc_cad import (
    extract_zip, scan_and_extract, build_preview, write_to_excel, detect_doc_type
)

bp = Blueprint('doccad', __name__, url_prefix='/doccad', template_folder='../../templates/doccad')

BASE_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


def _log_dir_tree(path, prefix="", max_depth=4, current_depth=0):
    """Log directory tree for debugging."""
    if current_depth > max_depth:
        return
    try:
        items = sorted(os.listdir(path))
        for item in items[:50]:  # limit to 50 items per dir
            full = os.path.join(path, item)
            if os.path.isdir(full):
                logger.info(f"{prefix}[DIR] {item}/")
                _log_dir_tree(full, prefix + "  ", max_depth, current_depth + 1)
            else:
                size = os.path.getsize(full)
                logger.info(f"{prefix}{item} ({size} bytes)")
    except Exception as e:
        logger.error(f"Error listing {path}: {e}")


def get_session_dir(subdir):
    if 'session_id' not in session:
        session['session_id'] = str(uuid.uuid4())
    path = os.path.join(BASE_DIR, subdir, session['session_id'])
    os.makedirs(path, exist_ok=True)
    return path


@bp.route('/')
def index():
    return render_template('doccad/index.html')


@bp.route('/upload', methods=['POST'])
def upload():
    upload_dir = get_session_dir('uploads')
    doccad_dir = os.path.join(upload_dir, 'doccad')
    if os.path.exists(doccad_dir):
        shutil.rmtree(doccad_dir)
    os.makedirs(doccad_dir, exist_ok=True)

    # --- SITUATIE Excel ---
    situatie_file = request.files.get('situatie')
    logger.info(f"Upload request - situatie: {situatie_file.filename if situatie_file else 'None'}")
    if not situatie_file or not situatie_file.filename:
        flash('Incarcati fisierul SITUATIE Excel!', 'error')
        return redirect(url_for('doccad.index'))

    situatie_path = os.path.join(doccad_dir, situatie_file.filename)
    situatie_file.save(situatie_path)
    session['doccad_situatie_path'] = situatie_path
    session['doccad_situatie_name'] = situatie_file.filename

    # --- Documentatii cadastrale ---
    doc_zip = request.files.get('doc_zip')
    doc_files = request.files.getlist('doc_files')

    extract_dir = os.path.join(doccad_dir, 'documentatii')
    os.makedirs(extract_dir, exist_ok=True)

    if doc_zip and doc_zip.filename:
        zip_path = os.path.join(doccad_dir, 'doc_cad.zip')
        doc_zip.save(zip_path)
        zip_size = os.path.getsize(zip_path)
        logger.info(f"ZIP saved: {zip_path} ({zip_size} bytes)")
        extract_zip(zip_path, extract_dir)
        os.remove(zip_path)
        # Log extracted contents
        _log_dir_tree(extract_dir, max_depth=4)
    elif doc_files and any(f.filename for f in doc_files):
        for f in doc_files:
            if not f.filename:
                continue
            # Preserve folder structure from webkitdirectory
            rel_path = f.filename.replace('\\', '/')
            full_path = os.path.join(extract_dir, rel_path)
            os.makedirs(os.path.dirname(full_path), exist_ok=True)
            f.save(full_path)
        _log_dir_tree(extract_dir, max_depth=4)
    else:
        flash('Incarcati documentatiile cadastrale (ZIP sau folder)!', 'error')
        return redirect(url_for('doccad.index'))

    session['doccad_extract_dir'] = extract_dir

    return redirect(url_for('doccad.analyze'))


@bp.route('/analyze')
def analyze():
    extract_dir = session.get('doccad_extract_dir')
    situatie_path = session.get('doccad_situatie_path')

    if not extract_dir or not situatie_path:
        flash('Sesiune expirata. Reincarcati fisierele.', 'error')
        return redirect(url_for('doccad.index'))

    try:
        # Log what we're working with
        logger.info(f"Analyzing extract_dir: {extract_dir}")
        _log_dir_tree(extract_dir, max_depth=3)

        doc_type, extracted = scan_and_extract(extract_dir)
        logger.info(f"detect result: type={doc_type}, items={len(extracted)}")
        if not extracted:
            # Build debug info about what was found
            debug_items = []
            try:
                for root, dirs, files in os.walk(extract_dir):
                    rel = os.path.relpath(root, extract_dir)
                    for f in files[:10]:
                        debug_items.append(f"{rel}/{f}" if rel != "." else f)
                    if len(debug_items) > 20:
                        break
            except Exception:
                pass
            debug_str = ", ".join(debug_items[:15]) if debug_items else "director gol"
            flash(f'Nu s-au gasit documentatii cadastrale! Fisiere gasite: {debug_str}', 'error')
            return redirect(url_for('doccad.index'))

        # Save extracted data to session (serialize for JSON)
        session['doccad_doc_type'] = doc_type
        extracted_json = os.path.join(os.path.dirname(extract_dir), 'extracted.json')
        # Save to file since session can't hold complex data well
        serializable = []
        for item in extracted:
            s_item = {
                'pozitie': item['pozitie'],
                'secondary_positions': item['secondary_positions'],
                'folder_name': item['folder_name'],
                'doc_type': item['doc_type'],
                'is_extra_hg': item.get('is_extra_hg', False),
                'error': item['error'],
            }
            if item['data']:
                s_item['data'] = {
                    'cf_data': item['data'].get('cf_data', {}),
                    'pad_data': item['data'].get('pad_data', {}),
                    'memoriu_data': item['data'].get('memoriu_data', {}),
                    'incheiere_resp_data': item['data'].get('incheiere_resp_data', {}),
                    'admitere_status': item['data'].get('admitere_status'),
                }
            else:
                s_item['data'] = None
            serializable.append(s_item)

        with open(extracted_json, 'w', encoding='utf-8') as f:
            json.dump(serializable, f, ensure_ascii=False, default=str)
        session['doccad_extracted_json'] = extracted_json

        # Build preview
        preview = build_preview(extracted, situatie_path)

        # Stats
        total = len(preview)
        ok_count = sum(1 for p in preview if p.get('is_ok'))
        diff_count = sum(1 for p in preview if not p.get('is_ok') and not p.get('error') and not p.get('already_filled') and not p.get('not_found'))
        error_count = sum(1 for p in preview if p.get('error'))
        filled_count = sum(1 for p in preview if p.get('already_filled'))
        not_found_count = sum(1 for p in preview if p.get('not_found'))
        to_process = sum(1 for p in preview if not p.get('error') and not p.get('already_filled') and not p.get('not_found'))

    except Exception as e:
        flash(f'Eroare la analiza documentatiilor: {str(e)}', 'error')
        import traceback
        traceback.print_exc()
        return redirect(url_for('doccad.index'))

    type_label = 'Scan complet PDF' if doc_type == 'scan_complet' else 'Foldere EXPROPRIAT / RAMAS'

    return render_template('doccad/preview.html',
        preview=preview,
        doc_type=doc_type,
        type_label=type_label,
        total=total,
        ok_count=ok_count,
        diff_count=diff_count,
        error_count=error_count,
        filled_count=filled_count,
        not_found_count=not_found_count,
        to_process=to_process,
        situatie_name=session.get('doccad_situatie_name', ''))


@bp.route('/generate', methods=['POST'])
def generate():
    extracted_json = session.get('doccad_extracted_json')
    situatie_path = session.get('doccad_situatie_path')

    if not extracted_json or not situatie_path:
        flash('Sesiune expirata. Reincarcati fisierele.', 'error')
        return redirect(url_for('doccad.index'))

    try:
        with open(extracted_json, 'r', encoding='utf-8') as f:
            extracted = json.load(f)

        output_dir = get_session_dir('output')
        output_name = session.get('doccad_situatie_name', 'SITUATIE_actualizata.xlsx')
        output_path = os.path.join(output_dir, output_name)

        summary = write_to_excel(extracted, situatie_path, output_path)
        session['doccad_output_path'] = output_path
        session['doccad_output_name'] = output_name
        session['doccad_summary'] = summary

    except Exception as e:
        flash(f'Eroare la generare: {str(e)}', 'error')
        import traceback
        traceback.print_exc()
        return redirect(url_for('doccad.analyze'))

    return redirect(url_for('doccad.results'))


@bp.route('/results')
def results():
    summary = session.get('doccad_summary')
    if not summary:
        flash('Nu exista rezultate.', 'error')
        return redirect(url_for('doccad.index'))

    return render_template('doccad/results.html',
        summary=summary,
        output_name=session.get('doccad_output_name', ''))


@bp.route('/download')
def download():
    output_path = session.get('doccad_output_path')
    if not output_path or not os.path.exists(output_path):
        flash('Fisierul nu mai exista. Reincarcati.', 'error')
        return redirect(url_for('doccad.index'))

    return send_from_directory(
        os.path.dirname(output_path),
        os.path.basename(output_path),
        as_attachment=True
    )
