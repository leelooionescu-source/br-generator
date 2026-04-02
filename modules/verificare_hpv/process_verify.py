import pandas as pd
import re
import os
from docx import Document
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side


def read_master(master_path):
    """Citeste MASTER si returneaza lista de dictionare cu datele per pozitie."""
    xls = pd.ExcelFile(master_path)
    all_entries = []
    for sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet_name, header=0)

        for i in range(len(df)):
            row = df.iloc[i]
            nr_pv = row.iloc[0]
            if pd.isna(nr_pv):
                continue
            try:
                nr_pv_int = int(float(nr_pv))
            except (ValueError, TypeError):
                continue

            def get_val(col_idx, default=''):
                v = row.iloc[col_idx] if col_idx < len(row) else default
                return str(v).strip() if pd.notna(v) else default

            entry = {
                'sheet': sheet_name,
                'nr_pv': nr_pv_int,
                'pozitie_hg': get_val(1),
                'judet': get_val(2),
                'uat': get_val(3),
                'nume1': get_val(5),
                'adresa1': get_val(6),
                'tarla': get_val(11),
                'parcela': get_val(12),
                'nr_cadastral': get_val(13),
                'nr_cf': get_val(14),
                'categorie': get_val(15),
                'extravilan_intravilan': get_val(16),
                'suprafata_totala': get_val(17),
                'suprafata_exp1': get_val(18),
                'valoare1': get_val(19),
                'suprafata_exp2': get_val(20),
                'valoare2': get_val(21),
                'hg': get_val(28),
                'decizie_expropriere': get_val(31),
                'decizie_comisie': get_val(32),
                'membru1': get_val(33),
                'membru2': get_val(34),
                'membru3': get_val(35),
                'membru4': get_val(36),
                'membru5': get_val(37),
                'rezulta': get_val(39),
                'data': get_val(40),
                'ora': get_val(41),
            }
            all_entries.append(entry)
    return all_entries


def read_docx(docx_path):
    """Extrage text complet din fisier .docx."""
    doc = Document(docx_path)
    paragraphs = [p.text for p in doc.paragraphs]
    # Also extract from tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                paragraphs.append(cell.text)
    return '\n'.join(paragraphs)


def split_sections(text):
    """Separa textul in sectiuni individuale H+PV (cate una per numar)."""
    # Split by "CONSILIUL LOCAL" - each H+PV starts with this
    parts = re.split(r'(?=CONSILIUL\s+LOCAL)', text)
    parts = [p for p in parts if p.strip() and len(p) > 500]

    # Group by nr - Hotararea si PV au acelasi numar
    sections = {}
    for p in parts:
        m = re.search(r'nr\.\s*(\d+)\s*din\s*([\d.]+)', p[:500])
        if m:
            nr = int(m.group(1))
            if nr not in sections:
                sections[nr] = ''
            sections[nr] += '\n' + p

    return sections


def extract_data_from_text(text, nr_pv, source_file):
    """Extrage date structurate din textul unui H si PV."""
    data = {'source_file': source_file, 'nr_pv': nr_pv}

    # Data
    m = re.search(r'nr\.\s*\d+\s*din\s*([\d.]+\.\d{4})', text)
    if m:
        data['data'] = m.group(1)

    # UAT si Judet
    m = re.search(r'CONSILIUL\s+LOCAL\s+(.+?)\s*,\s*JUDE[ȚŢ]UL\s+(\w+)', text)
    if m:
        data['uat'] = m.group(1).strip()
        data['judet'] = m.group(2).strip()

    # Pozitie HG
    m = re.search(r'pozi[țţ]i[ae]\s*(?:nr\.?)?\s*(\d+)\s*din\s*Anex', text, re.IGNORECASE)
    if m:
        data['pozitie_hg'] = m.group(1)

    # Nr cadastral
    m = re.search(r'num[aă]r\s+cadastral\s*/?\s*(?:nr\.?)?\s*topo\s+(\d+)', text, re.IGNORECASE)
    if m:
        data['nr_cadastral'] = m.group(1)

    # Nr CF
    m = re.search(r'carte\s+funciar[aă]\s+(\d+)', text, re.IGNORECASE)
    if m:
        data['nr_cf'] = m.group(1)

    # Tarla si parcela
    m = re.search(r'tarla\s*(?:nr\.?)?\s*([\d/]+)\s*,?\s*parcela\s*(?:nr\.?)?\s*([\d/]+)', text, re.IGNORECASE)
    if m:
        data['tarla'] = m.group(1)
        data['parcela'] = m.group(2)

    # Suprafete expropriate
    surfaces = re.findall(r'Teren\d?\s+[iî]n\s+suprafa[tț][aă]\s+de\s+([\d.,]+)\s*mp', text, re.IGNORECASE)
    if surfaces:
        data['suprafete'] = [s.replace('.', '').replace(',', '.') for s in surfaces]

    # Valori despagubiri
    vals = re.findall(r'([\d.,]+)\s*LEI.*?(?:pentru\s+imobilul\s+teren|stabilit[aă]\s+cu\s+titlu)', text, re.IGNORECASE)
    if vals:
        data['valori'] = [v.replace('.', '').replace(',', '.') for v in vals]

    # Suma totala
    m = re.search(r'suma\s+de\s+([\d.,]+)\s*LEI', text, re.IGNORECASE)
    if m:
        data['suma_totala'] = m.group(1).replace('.', '').replace(',', '.')

    # Proprietar
    m = re.search(r'supus[eă]?\s+exproprierii.*?\n\s*1[\.\)]\s*(.+?)(?:,\s*cu\s+(?:domiciliul|sediul)|;\s*\n)', text, re.IGNORECASE | re.DOTALL)
    if m:
        name = re.sub(r'\s+', ' ', m.group(1)).strip()
        if not any(kw in name for kw in ['REPREZENTANT', 'Primaria', 'AV.', 'LEI', 'teren']):
            data['proprietar'] = name

    # Decizie expropriere
    m = re.search(r'[Dd]eciziei\s+de\s+expropriere\s+nr\.\s*(\d+)\s+din\s+([\d.]+)', text)
    if m:
        data['decizie_expropriere'] = f'nr. {m.group(1)} din {m.group(2)}'

    # Decizie comisie
    m = re.search(r'[Dd]eciziei\s+nr\.\s*(\d+)\s+din\s+([\d.]+)\s+emis', text)
    if m:
        data['decizie_comisie'] = f'nr. {m.group(1)} din {m.group(2)}'

    # HG
    m = re.search(r'(?:Guvernului|H\.?\s*G\.?)\s*nr\.?\s*(\d{2,4}\s*/\s*\d{4})', text)
    if m:
        data['hg'] = re.sub(r'\s', '', m.group(1))

    return data


def compare_entry(master, doc_data):
    """Compara datele din MASTER cu cele din document. Returneaza lista de neconcordante."""
    issues = []

    def normalize(s):
        s = str(s).strip().upper()
        for old, new in [('Ț', 'T'), ('Ş', 'S'), ('Ș', 'S'), ('Ă', 'A'), ('Â', 'A'), ('Î', 'I'), ('Ţ', 'T')]:
            s = s.replace(old, new)
        return re.sub(r'\s+', ' ', s)

    def check(field_name, master_val, doc_val):
        if not doc_val:
            return
        m = normalize(master_val)
        p = normalize(doc_val)
        if m and p and m != p:
            if m in p or p in m:
                return
            issues.append({'camp': field_name, 'master': str(master_val).strip(), 'document': str(doc_val).strip()})

    check('UAT', master.get('uat', ''), doc_data.get('uat', ''))
    check('Județ', master.get('judet', ''), doc_data.get('judet', ''))
    check('Poziție HG', master.get('pozitie_hg', ''), doc_data.get('pozitie_hg', ''))
    check('Nr. cadastral', master.get('nr_cadastral', ''), doc_data.get('nr_cadastral', ''))
    check('Nr. CF', master.get('nr_cf', ''), doc_data.get('nr_cf', ''))
    check('Tarla', master.get('tarla', ''), doc_data.get('tarla', ''))
    check('Parcela', master.get('parcela', ''), doc_data.get('parcela', ''))
    check('Data', master.get('data', ''), doc_data.get('data', ''))
    check('HG', master.get('hg', ''), doc_data.get('hg', ''))

    # Proprietar
    if doc_data.get('proprietar'):
        m_name = normalize(master.get('nume1', ''))
        p_name = normalize(doc_data.get('proprietar', ''))
        if m_name and p_name:
            m_parts = set(m_name.split())
            p_parts = set(p_name.split())
            common = m_parts & p_parts
            if len(common) < min(2, len(m_parts)):
                issues.append({'camp': 'Proprietar', 'master': master.get('nume1', ''), 'document': doc_data.get('proprietar', '')})

    # Suprafete
    if doc_data.get('suprafete'):
        master_sups = []
        for key in ['suprafata_exp1', 'suprafata_exp2']:
            v = master.get(key, '')
            if v and v not in ('', '0', '-'):
                try:
                    master_sups.append(float(str(v).replace(',', '.')))
                except ValueError:
                    pass
        doc_sups = []
        for s in doc_data['suprafete']:
            try:
                doc_sups.append(float(s))
            except ValueError:
                pass
        if master_sups and doc_sups:
            for ms in master_sups:
                if not any(abs(ms - ds) < 1 for ds in doc_sups):
                    issues.append({
                        'camp': 'Suprafață expropriată',
                        'master': ', '.join(str(int(s)) for s in master_sups),
                        'document': ', '.join(str(int(s)) for s in doc_sups),
                    })
                    break

    # Valori
    if doc_data.get('suma_totala'):
        master_vals = []
        for key in ['valoare1', 'valoare2']:
            v = master.get(key, '')
            if v and v not in ('', '0', '-'):
                try:
                    master_vals.append(float(str(v).replace(',', '.')))
                except ValueError:
                    pass
        master_total = sum(master_vals)
        try:
            doc_total = float(doc_data['suma_totala'])
        except ValueError:
            doc_total = None
        if master_total > 0 and doc_total and abs(master_total - doc_total) > 1:
            issues.append({
                'camp': 'Valoare despăgubiri',
                'master': f'{master_total:,.2f}'.replace(',', '.'),
                'document': f'{doc_total:,.2f}'.replace(',', '.'),
            })

    # Decizie expropriere
    if doc_data.get('decizie_expropriere'):
        m_nums = re.findall(r'\d+', normalize(master.get('decizie_expropriere', '')))
        d_nums = re.findall(r'\d+', normalize(doc_data.get('decizie_expropriere', '')))
        if m_nums and d_nums and m_nums != d_nums:
            issues.append({'camp': 'Decizie expropriere', 'master': master.get('decizie_expropriere', ''), 'document': doc_data.get('decizie_expropriere', '')})

    return issues


def process_all(master_path, docx_files, progress_callback=None):
    """Proceseaza MASTER + toate fisierele .docx. Returneaza rezultate."""
    master_entries = read_master(master_path)
    master_by_pv = {e['nr_pv']: e for e in master_entries}

    results = []
    total_sections = 0

    # First pass: count total sections
    file_sections = []
    for docx_path in docx_files:
        filename = os.path.basename(docx_path)
        text = read_docx(docx_path)
        sections = split_sections(text)
        file_sections.append((filename, sections))
        total_sections += len(sections)

    processed = 0
    for filename, sections in file_sections:
        for nr_pv, section_text in sorted(sections.items()):
            processed += 1
            if progress_callback:
                progress_callback(processed, total_sections, f'{filename} - H/PV nr. {nr_pv}')

            doc_data = extract_data_from_text(section_text, nr_pv, filename)

            if nr_pv in master_by_pv:
                master_entry = master_by_pv[nr_pv]
                issues = compare_entry(master_entry, doc_data)
                results.append({
                    'source_file': filename,
                    'nr_pv': nr_pv,
                    'uat': master_entry.get('uat', ''),
                    'pozitie_hg': master_entry.get('pozitie_hg', ''),
                    'proprietar': master_entry.get('nume1', ''),
                    'issues': issues,
                    'status': 'NECONCORDANȚE' if issues else 'OK',
                })
            else:
                results.append({
                    'source_file': filename,
                    'nr_pv': nr_pv,
                    'uat': doc_data.get('uat', ''),
                    'pozitie_hg': doc_data.get('pozitie_hg', ''),
                    'proprietar': doc_data.get('proprietar', ''),
                    'issues': [{'camp': 'Nr. PV', 'master': 'NEGĂSIT', 'document': str(nr_pv)}],
                    'status': 'NEGĂSIT ÎN MASTER',
                })

    return results, master_entries


def generate_report(results, output_path):
    """Genereaza raport Excel cu neconcordantele."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Neconcordanțe"

    hf = Font(name='Arial', bold=True, size=11, color='FFFFFF')
    hfill = PatternFill('solid', fgColor='C00000')
    df = Font(name='Arial', size=10)
    tb = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    ca = Alignment(horizontal='center', vertical='center', wrap_text=True)

    ws.merge_cells('A1:H1')
    ws['A1'] = 'RAPORT NECONCORDANȚE - VERIFICARE H SI PV vs MASTER'
    ws['A1'].font = Font(name='Arial', bold=True, size=14, color='C00000')
    ws['A1'].alignment = Alignment(horizontal='center')

    headers = ['Nr.\ncrt.', 'Fișier sursă', 'Nr.\nPV/HSD', 'UAT', 'Poziție\nHG', 'Câmp cu\nneconcordanță', 'Valoare\nMASTER', 'Valoare\nDocument']
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=3, column=col, value=h)
        cell.font = hf
        cell.fill = hfill
        cell.alignment = ca
        cell.border = tb

    row = 4
    crt = 0
    for r in results:
        if not r['issues']:
            continue
        for issue in r['issues']:
            crt += 1
            vals = [crt, r['source_file'], r['nr_pv'], r.get('uat', ''),
                    r.get('pozitie_hg', ''), issue['camp'], issue['master'], issue['document']]
            for col, v in enumerate(vals, 1):
                cell = ws.cell(row=row, column=col, value=v)
                cell.font = df
                cell.alignment = ca if col <= 5 else Alignment(horizontal='left', vertical='center', wrap_text=True)
                cell.border = tb
            row += 1

    # Summary sheet
    ws2 = wb.create_sheet("Sumar")
    ws2['A1'] = 'SUMAR VERIFICARE'
    ws2['A1'].font = Font(name='Arial', bold=True, size=14)
    ws2['A3'] = 'Total H/PV verificate:'
    ws2['B3'] = len(results)
    ws2['A4'] = 'OK (fără neconcordanțe):'
    ws2['B4'] = sum(1 for r in results if not r['issues'])
    ws2['A5'] = 'Cu neconcordanțe:'
    ws2['B5'] = sum(1 for r in results if r['issues'])
    ws2['A5'].font = Font(name='Arial', bold=True, color='FF0000')
    ws2['B5'].font = Font(name='Arial', bold=True, color='FF0000')
    ws2['A6'] = 'Total neconcordanțe găsite:'
    ws2['B6'] = sum(len(r['issues']) for r in results)

    for c in ['A', 'B']:
        ws2.column_dimensions[c].width = 35

    ws.column_dimensions['A'].width = 8
    ws.column_dimensions['B'].width = 25
    ws.column_dimensions['C'].width = 10
    ws.column_dimensions['D'].width = 15
    ws.column_dimensions['E'].width = 10
    ws.column_dimensions['F'].width = 22
    ws.column_dimensions['G'].width = 30
    ws.column_dimensions['H'].width = 30
    ws.row_dimensions[3].height = 40

    wb.save(output_path)
    return output_path
