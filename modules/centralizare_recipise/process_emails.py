"""
process_emails.py — Citire emailuri din Outlook + parsare subiecte
"""

import os
import re
from modules.centralizare_recipise.database import (
    email_exists, insert_email, insert_attachment, ATTACHMENTS_DIR
)

# Keywords pentru filtrare emailuri
KEYWORDS = ['recipis', 'reconsemn']

# Regex patterns pentru parsare subiect
HG_PATTERN = re.compile(r'[Hh][Gg]\s*(?:nr\.?\s*)?(\d+\s*[/-]\s*\d{4})', re.IGNORECASE)
BR_PATTERN = re.compile(r'(B[RP])\.?\s*(\d+(?:\.\d+)?)', re.IGNORECASE)
DATE_PATTERN = re.compile(r'(\d{2}\.\d{2}\.\d{4})')


def parse_subject(subject):
    """Parseaza subiectul emailului si extrage HG, BR/BP, data."""
    if not subject:
        return {'hg_number': '', 'br_number': '', 'br_date': ''}

    hg_match = HG_PATTERN.search(subject)
    hg_number = hg_match.group(1).replace(' ', '') if hg_match else ''

    br_match = BR_PATTERN.search(subject)
    br_number = f"{br_match.group(1).upper()} {br_match.group(2)}" if br_match else ''

    date_match = DATE_PATTERN.search(subject)
    br_date = date_match.group(1) if date_match else ''

    return {
        'hg_number': hg_number,
        'br_number': br_number,
        'br_date': br_date,
    }


def _find_folder(folder, target_name):
    """Cauta recursiv un folder in Outlook dupa nume."""
    try:
        if folder.Name == target_name:
            return folder
        for i in range(1, folder.Folders.Count + 1):
            result = _find_folder(folder.Folders.Item(i), target_name)
            if result:
                return result
    except Exception:
        pass
    return None


def sync_from_outlook(folder_name='A.N.D.'):
    """Sincronizeaza emailurile din Outlook folderul specificat.
    Returneaza statistici: {new, skipped, errors, total_attachments}."""
    try:
        import win32com.client
        import pythoncom
    except ImportError:
        raise RuntimeError(
            'pywin32 nu este instalat. Modulul de sincronizare '
            'necesita Microsoft Outlook pe Windows.\n'
            'Instalati cu: pip install pywin32'
        )

    pythoncom.CoInitialize()
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        mapi = outlook.GetNamespace("MAPI")

        # Cauta folderul in toate store-urile
        target_folder = None
        for i in range(1, mapi.Stores.Count + 1):
            store = mapi.Stores.Item(i)
            try:
                root = store.GetRootFolder()
                target_folder = _find_folder(root, folder_name)
                if target_folder:
                    break
            except Exception:
                continue

        if not target_folder:
            raise ValueError(
                f'Folderul "{folder_name}" nu a fost gasit in Outlook. '
                f'Verificati ca exista si ca Outlook este deschis.'
            )

        stats = {'new': 0, 'skipped': 0, 'errors': 0, 'total_attachments': 0}
        items = target_folder.Items
        items.Sort("[ReceivedTime]", True)  # Cele mai recente primele

        for i in range(1, items.Count + 1):
            try:
                item = items.Item(i)
                # Verificam daca e un MailItem
                if item.Class != 43:  # 43 = olMail
                    continue

                subject = item.Subject or ''

                # Filtrare dupa keywords
                subject_lower = subject.lower()
                if not any(k in subject_lower for k in KEYWORDS):
                    continue

                # Deduplicare
                entry_id = item.EntryID
                if email_exists(entry_id):
                    stats['skipped'] += 1
                    continue

                # Parsare subiect
                parsed = parse_subject(subject)

                # Date email
                sender = ''
                try:
                    sender = item.SenderEmailAddress or ''
                    if '@' not in sender:
                        sender = item.Sender.GetExchangeUser().PrimarySmtpAddress
                except Exception:
                    try:
                        sender = item.SenderName or ''
                    except Exception:
                        pass

                received = ''
                try:
                    received = item.ReceivedTime.strftime('%Y-%m-%d %H:%M:%S')
                except Exception:
                    pass

                body = ''
                try:
                    body = item.Body or ''
                except Exception:
                    pass

                # Salveaza emailul in DB
                att_count = item.Attachments.Count
                email_id = insert_email(
                    entry_id=entry_id,
                    subject=subject,
                    sender=sender,
                    received_date=received,
                    body=body,
                    hg_number=parsed['hg_number'],
                    br_number=parsed['br_number'],
                    br_date=parsed['br_date'],
                    attachment_count=att_count,
                )

                if email_id and att_count > 0:
                    # Salveaza atasamentele
                    att_dir = os.path.join(ATTACHMENTS_DIR, str(email_id))
                    os.makedirs(att_dir, exist_ok=True)

                    for j in range(1, att_count + 1):
                        try:
                            att = item.Attachments.Item(j)
                            filename = att.FileName
                            if not filename:
                                continue
                            file_path = os.path.join(att_dir, filename)
                            att.SaveAsFile(file_path)
                            size = os.path.getsize(file_path) if os.path.exists(file_path) else 0
                            insert_attachment(email_id, filename, size, file_path)
                            stats['total_attachments'] += 1
                        except Exception:
                            stats['errors'] += 1

                stats['new'] += 1

            except Exception:
                stats['errors'] += 1

        return stats

    finally:
        pythoncom.CoUninitialize()
