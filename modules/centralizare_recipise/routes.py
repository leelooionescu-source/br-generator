import os
import tempfile
from flask import Blueprint, render_template, request, redirect, url_for, flash, send_file
from modules.centralizare_recipise.database import (
    get_all_emails, get_email_by_id, get_attachments_for_email,
    get_hg_list, get_stats, update_email_fields
)
from modules.centralizare_recipise.process_emails import sync_from_outlook
from modules.centralizare_recipise.export_excel import generate_excel_report

bp = Blueprint('recipise', __name__, url_prefix='/recipise', template_folder='../../templates/recipise')


@bp.route('/')
def index():
    filters = {
        'hg': request.args.get('hg', ''),
        'br': request.args.get('br', ''),
        'sender': request.args.get('sender', ''),
    }
    emails = get_all_emails(
        hg_filter=filters['hg'] or None,
        br_filter=filters['br'] or None,
        sender_filter=filters['sender'] or None,
    )
    stats = get_stats()
    hg_list = get_hg_list()
    sync_result = request.args.get('sync_result')

    # Decode sync_result from query params
    sr = None
    if sync_result:
        try:
            parts = sync_result.split(',')
            sr = {
                'new': int(parts[0]),
                'skipped': int(parts[1]),
                'total_attachments': int(parts[2]),
                'errors': int(parts[3]),
            }
        except (ValueError, IndexError):
            pass

    return render_template('recipise/index.html',
        emails=emails, stats=stats, hg_list=hg_list,
        filters=filters, sync_result=sr)


@bp.route('/sync', methods=['POST'])
def sync():
    try:
        result = sync_from_outlook(folder_name='A.N.D.')
        sr = f"{result['new']},{result['skipped']},{result['total_attachments']},{result['errors']}"
        return redirect(url_for('recipise.index', sync_result=sr))
    except Exception as e:
        flash(f'Eroare la sincronizare: {str(e)}', 'error')
        return redirect(url_for('recipise.index'))


@bp.route('/detail/<int:email_id>')
def detail(email_id):
    email = get_email_by_id(email_id)
    if not email:
        flash('Emailul nu a fost gasit.', 'error')
        return redirect(url_for('recipise.index'))
    attachments = get_attachments_for_email(email_id)
    return render_template('recipise/detail.html', email=email, attachments=attachments)


@bp.route('/update/<int:email_id>', methods=['POST'])
def update_fields(email_id):
    hg = request.form.get('hg_number', '').strip()
    br = request.form.get('br_number', '').strip()
    br_date = request.form.get('br_date', '').strip()
    update_email_fields(email_id, hg_number=hg, br_number=br, br_date=br_date)
    flash('Campuri actualizate cu succes.', 'success')
    return redirect(url_for('recipise.detail', email_id=email_id))


@bp.route('/attachment/<int:att_id>')
def download_attachment(att_id):
    from modules.centralizare_recipise.database import db_session
    with db_session() as conn:
        row = conn.execute("SELECT * FROM attachments WHERE id = ?", (att_id,)).fetchone()
    if not row:
        flash('Atasamentul nu a fost gasit.', 'error')
        return redirect(url_for('recipise.index'))
    file_path = row['file_path']
    if not os.path.exists(file_path):
        flash('Fisierul atasament nu exista pe disk.', 'error')
        return redirect(url_for('recipise.index'))
    return send_file(file_path, as_attachment=True, download_name=row['filename'])


@bp.route('/export-excel')
def export_excel():
    try:
        filepath = generate_excel_report()
        return send_file(filepath, as_attachment=True,
                         download_name='Centralizare Recipise.xlsx')
    except Exception as e:
        flash(f'Eroare la export Excel: {str(e)}', 'error')
        return redirect(url_for('recipise.index'))
