#!/usr/bin/env python3
"""
fetch_policy_emails.py

功能：
 - 使用 win32com 从指定 Outlook 账户和文件夹抓取指定日期范围的邮件
 - 将邮件正文保存为文本文件；白名单附件单独保存（可选提取 PDF 文本并内联）
 - 将所有保存的文本聚合并通过 OpenAI 两步工作流（抽取→报告）生成周报（Markdown）

配置：在工作目录放置 `.env` 或在环境变量中设置 `OPENAI_API_KEY`
注意：需在 Windows 环境、Outlook 已配置且能访问目标邮箱时运行。
"""
import os
import sys
import re
import json
import hashlib
import logging
import pathlib
import argparse
import datetime
import asyncio
import threading
from typing import List
from email.utils import parsedate_to_datetime

# Optional deps
try:
    import win32com.client  # pywin32
except Exception:
    win32com = None
else:
    win32com = win32com

try:
    import pdfplumber
except Exception:
    pdfplumber = None

try:
    from openai import OpenAI
except Exception:
    OpenAI = None

from dotenv import load_dotenv
from prompts_extractor import EXTRACTOR_SYSTEM_PROMPT, EXTRACTOR_USER_TEMPLATE
from prompts_report_generator import REPORT_SYSTEM_PROMPT, REPORT_USER_TEMPLATE

load_dotenv()

USE_MOCK = os.getenv('USE_MOCK_EMAILS', '').lower() == 'true'

LOG = logging.getLogger("fetch_policy_emails")
LOG.setLevel(logging.INFO)
LOG.addHandler(logging.StreamHandler())


def _split_folder_path(s: str) -> List[str]:
    """Split a folder path string like '收件箱/2. Policy Update' into ['收件箱','2. Policy Update'].
    Accepts separators '/', '\\', '>', '|'.

    改进：修复因空字节导致的路径分割错误，简化分隔符处理逻辑
    """
    if not s:
        return []
    # normalize any of /, \, >, | (one or more) into '/'
    normalized = re.sub(r"[\\/>|]+", "/", s)
    parts = [p.strip() for p in normalized.split('/') if p.strip()]
    return parts


def list_outlook_mailboxes():
    if win32com is None:
        print('pywin32 not available')
        return
    ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    print('Mailboxes in current Outlook profile:')
    try:
        for i in range(ns.Folders.Count):
            root = ns.Folders.Item(i + 1)
            print(f"- {getattr(root, 'Name', '(unknown)')}")
    except Exception:
        # fallback best-effort
        try:
            for root in ns.Folders:
                print(f"- {getattr(root, 'Name', '(unknown)')}")
        except Exception:
            pass


def _print_folder_tree(folder, prefix: str, level: int, max_depth: int):
    name = getattr(folder, 'Name', '(unknown)')
    print(prefix + name)
    if level >= max_depth:
        return
    try:
        subs = folder.Folders
        cnt = getattr(subs, 'Count', 0) or 0
        for i in range(1, cnt + 1):
            sub = subs.Item(i)
            _print_folder_tree(sub, prefix + '  ', level + 1, max_depth)
    except Exception:
        return


def list_outlook_folders(mailbox_name: str, max_depth: int = 2):
    if win32com is None:
        print('pywin32 not available')
        return
    try:
        root = get_outlook_folder(mailbox_name, [])
    except Exception:
        # get mailbox root directly via Namespace
        ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        root = None
        for i in range(ns.Folders.Count):
            r = ns.Folders.Item(i + 1)
            if getattr(r, 'Name', '').lower() == mailbox_name.lower():
                root = r
                break
        if root is None:
            print(f"Mailbox not found: {mailbox_name}")
            return
    print(f"Folders under mailbox: {mailbox_name}")
    _print_folder_tree(root, prefix='', level=0, max_depth=max_depth)


def get_outlook_folder(mailbox_name: str, folder_path: List[str]):
    """Return an Outlook MAPIFolder for mailbox and nested folder path."""
    if win32com is None:
        raise RuntimeError("pywin32 not installed or not available")

    ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

    # Get top-level mailbox
    acct = None
    for i in range(ns.Folders.Count):
        root = ns.Folders.Item(i + 1)
        if getattr(root, 'Name', '').lower() == mailbox_name.lower():
            acct = root
            break
    if acct is None:
        try:
            acct = ns.Folders[mailbox_name]
        except Exception:
            raise FileNotFoundError(f"Mailbox '{mailbox_name}' not found in Outlook profile")

    folder = acct
    for part in folder_path:
        try:
            folder = folder.Folders[part]
        except Exception:
            raise FileNotFoundError(f"Folder part '{part}' not found under mailbox '{mailbox_name}'")
    return folder


def to_naive_local(dt: datetime.datetime) -> datetime.datetime:
    """Convert aware datetime to naive local datetime. If dt is already naive, return as-is."""
    if dt is None:
        return None
    if dt.tzinfo is None:
        return dt
    try:
        return dt.astimezone().replace(tzinfo=None)
    except Exception:
        return dt.replace(tzinfo=None)


def restrict_items_range(folder, since_dt: datetime.datetime, until_dt: datetime.datetime):
    """Try a robust DASL (@SQL) filter first; fallback to MAPI format if needed.
    Returns an Items collection (possibly already filtered).
    """
    items = getattr(folder, 'Items', None)
    if items is None:
        return None

    # Include recurrences and sort for deterministic iteration
    try:
        items.IncludeRecurrences = True
    except Exception:
        pass
    try:
        items.Sort("[ReceivedTime]", True)  # descending
    except Exception:
        pass

    # 1) Try DASL (@SQL) with ISO-like timestamp
    dasl_since = since_dt.strftime('%Y-%m-%d %H:%M:%S')
    dasl_until = until_dt.strftime('%Y-%m-%d %H:%M:%S')
    dasl = (
        f"@SQL=\"urn:schemas:httpmail:datereceived\" >= '{dasl_since}' "
        f"AND \"urn:schemas:httpmail:datereceived\" <= '{dasl_until}'"
    )
    try:
        filtered = items.Restrict(dasl)
        LOG.info('Applied Outlook DASL restriction: %s', dasl)
        # If we get any items, return
        try:
            if getattr(filtered, 'Count', 0) > 0:
                return filtered
        except Exception:
            return filtered
    except Exception:
        LOG.warning('DASL Restrict failed; will try MAPI format next')

    # 2) Fallback to MAPI format (US MM/DD/YYYY hh:mm AM/PM)
    try:
        since_str = since_dt.strftime('%m/%d/%Y %I:%M %p')
        until_str = until_dt.strftime('%m/%d/%Y %I:%M %p')
    except Exception:
        since_str = since_dt.strftime('%m/%d/%Y')
        until_str = until_dt.strftime('%m/%d/%Y')

    mapi_restrict = f"[ReceivedTime] >= '{since_str}' AND [ReceivedTime] <= '{until_str}'"
    try:
        filtered = items.Restrict(mapi_restrict)
        LOG.info('Applied Outlook restriction: %s', mapi_restrict)
        return filtered
    except Exception:
        LOG.warning('Outlook Restrict (MAPI) failed; returning unfiltered Items (client-side filter will apply)')
        return items


def _parse_header_date_from_raw_headers(raw_headers: str) -> datetime.datetime | None:
    """Parse 'Date:' header from raw headers and return naive local datetime."""
    if not raw_headers:
        return None
    date_line = None
    for line in raw_headers.splitlines():
        if line.lower().startswith('date:'):
            date_line = line.split(':', 1)[1].strip()
            break
    if not date_line:
        return None
    try:
        dt = parsedate_to_datetime(date_line)
        if dt is None:
            return None
        return to_naive_local(dt)
    except Exception:
        return None


def save_mail_and_attachments(msg, outdir: pathlib.Path):
    """Save email body as txt and attachments (document whitelist).
    Returns (txt_path, saved_attachments, received_local, message_id, chosen_date, date_source)
    """
    subj = getattr(msg, 'Subject', '') or 'no_subject'
    received = getattr(msg, 'ReceivedTime', datetime.datetime.now())
    received = to_naive_local(received)

    safe_subj = ''.join(c for c in subj if c.isalnum() or c in ' _-').strip()[:100]
    base_name = f"{received.strftime('%Y%m%d_%H%M%S')}_{safe_subj}"

    txt_path = outdir / (base_name + '.txt')
    body = getattr(msg, 'Body', '') or ''
    with txt_path.open('w', encoding='utf-8') as f:
        f.write(f"Subject: {subj}\n")
        f.write(f"From: {getattr(msg,'SenderName','')}\n")
        f.write(f"Received: {received}\n\n")
        f.write(body)

    # Build message id
    message_id = None
    try:
        message_id = getattr(msg, 'InternetMessageID', None)
    except Exception:
        pass
    if not message_id:
        try:
            message_id = getattr(msg, 'EntryID', None)
        except Exception:
            pass
    if not message_id:
        body_snip = (body or '')[:500]
        sender = getattr(msg, 'SenderEmailAddress', '') or getattr(msg, 'SenderName', '') or ''
        key = f"{subj}|{sender}|{received}|{body_snip}"
        message_id = hashlib.sha256(key.encode('utf-8')).hexdigest()

    # Parse header Date from raw headers if available
    parsed_header_date = None
    date_source = 'received_time'
    try:
        pa = getattr(msg, 'PropertyAccessor', None)
        raw_headers = None
        if pa is not None:
            try:
                raw_headers = pa.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x007D001E")
            except Exception:
                raw_headers = None
        if raw_headers:
            parsed_header_date = _parse_header_date_from_raw_headers(raw_headers)
            if parsed_header_date:
                date_source = 'header_date'
    except Exception:
        parsed_header_date = None
        date_source = 'received_time'

    chosen_date = parsed_header_date or received

    # Save attachments — whitelist
    attachments = getattr(msg, 'Attachments', None)
    saved_att = []
    allowed_exts = {'.pdf', '.docx', '.doc', '.xlsx', '.xls', '.pptx', '.ppt', '.txt', '.csv'}
    if attachments is not None:
        try:
            count = attachments.Count
        except Exception:
            count = 0
        for i in range(1, count + 1):
            att = attachments.Item(i)
            att_name = getattr(att, 'FileName', None)
            if not att_name:
                continue
            _, ext = os.path.splitext(att_name)
            if ext.lower() not in allowed_exts:
                LOG.info('skipping non-document attachment %s (ext=%s)', att_name, ext)
                continue
            att_path = outdir / (base_name + '_' + att_name)
            try:
                att.SaveAsFile(str(att_path))
                saved_att.append(str(att_path))
            except Exception as e:
                LOG.warning('failed to save attachment %s: %s', att_name, e)

    return str(txt_path), saved_att, received, message_id, chosen_date, date_source


def build_package_from_saved_emails(saved_emails: List[dict]) -> str:
    """Build a single text payload for OpenAI from saved email files and attachments."""
    parts = []
    for idx, rec in enumerate(saved_emails):
        txt_path = rec.get('txt')
        try:
            with open(txt_path, 'r', encoding='utf-8') as f:
                text = f.read()
        except Exception:
            LOG.warning('failed to read email text %s', txt_path)
            text = ''

        header = f"---EMAIL {idx+1}/{len(saved_emails)}---\n"
        part = header + text

        for a in rec.get('attachments', []):
            if a.lower().endswith('.pdf'):
                pdf_text = extract_text_from_pdf(a)
                if pdf_text.strip():
                    part += f"\n\n[PDF: {os.path.basename(a)}]\n{pdf_text}\n"

        parts.append(part)

    return '\n\n---EMAIL_BREAK---\n\n'.join(parts)


def extract_text_from_pdf(path: str) -> str:
    if pdfplumber is None:
        LOG.warning('pdfplumber not installed; skipping PDF text extraction for %s', path)
        return ''
    try:
        with pdfplumber.open(path) as pdf:
            extracted = []
            for page in pdf.pages:
                try:
                    extracted.append(page.extract_text() or '')
                except Exception:
                    extracted.append('')
            return '\n'.join(filter(None, extracted))
    except Exception:
        LOG.exception('pdfplumber failed for %s', path)
        return ''


def _extract_first_json_block(text: str) -> str:
    """Best-effort: extract the first top-level JSON object from text."""
    start = text.find('{')
    end = text.rfind('}')
    if start != -1 and end != -1 and end > start:
        candidate = text[start:end+1]
        try:
            json.loads(candidate)
            return candidate
        except Exception:
            pass
    m = re.search(r'\{.*\}', text, re.S)
    if m:
        try:
            json.loads(m.group(0))
            return m.group(0)
        except Exception:
            return m.group(0)
    return text


def call_openai_extract_updates(raw_text: str, file_label: str, openai_api_key: str, model: str = 'gpt-5-mini') -> str:
    """Step 1: Use extractor prompts to produce a single JSON string."""
    if OpenAI is None:
        raise RuntimeError('openai package not installed (requires openai>=1.0.0)')
    client = OpenAI(api_key=openai_api_key)

    raw_with_hint = f"[[FILE:{file_label}]]\n" + raw_text
    user_msg = EXTRACTOR_USER_TEMPLATE.format(RAW_TEXT=raw_with_hint)

    LOG.info('calling OpenAI extractor (size=%d chars)', len(user_msg))
    resp = client.chat.completions.create(
        model=model,
        messages=[
            {"role": "system", "content": EXTRACTOR_SYSTEM_PROMPT},
            {"role": "user", "content": user_msg},
        ],
    )
    content = resp.choices[0].message.content if getattr(resp, 'choices', None) else str(resp)
    json_text = content.strip()
    try:
        json.loads(json_text)
    except Exception:
        json_text = _extract_first_json_block(json_text)
    return json_text


def call_openai_generate_markdown(structured_json_text: str, period_str: str, openai_api_key: str, model: str = 'gpt-5-mini') -> str:
    """Step 2: Use report generator prompts to produce Markdown text."""
    if OpenAI is None:
        raise RuntimeError('openai package not installed (requires openai>=1.0.0)')
    client = OpenAI(api_key=openai_api_key)

    user_msg = REPORT_USER_TEMPLATE.format(STRUCTURED_JSON=structured_json_text, PERIOD=period_str)
    LOG.info('calling OpenAI report generator')
    resp = client.chat.completions.create(
        model=model,
        messages=[
            {"role": "system", "content": REPORT_SYSTEM_PROMPT},
            {"role": "user", "content": user_msg},
        ],
    )
    return resp.choices[0].message.content if getattr(resp, 'choices', None) else str(resp)


async def _wait_with_progress(coro_factory, interval: int = 3):
    """Run the provided coroutine factory, print progress periodically while awaiting."""
    if asyncio.iscoroutinefunction(coro_factory):
        task = asyncio.create_task(coro_factory())
    else:
        task = asyncio.create_task(coro_factory)
    while not task.done():
        LOG.info('waiting for OpenAI result...')
        try:
            await asyncio.wait_for(asyncio.shield(task), timeout=interval)
        except asyncio.TimeoutError:
            continue
    return task.result()


def main():
    # Defaults from environment (overridable via CLI)
    default_mailbox = os.getenv('OUTLOOK_MAILBOX', 'henry.wen@ocmortgage.com.au')
    default_folder_str = os.getenv('OUTLOOK_FOLDER_PATH', '收件箱/2. Policy Update')
    default_folder_path = _split_folder_path(default_folder_str)

    parser = argparse.ArgumentParser()
    parser.add_argument('--days', type=int, default=7, help='How many days back to fetch emails')
    parser.add_argument('--since', type=str, help='Start date YYYY-MM-DD (local)')
    parser.add_argument('--until', type=str, help='End date YYYY-MM-DD (local)')
    parser.add_argument('--dump-payload', type=str, help='Write merged payload to this file')
    parser.add_argument('--only-dump', action='store_true', help='Only dump merged payload and exit (skip OpenAI)')
    parser.add_argument('--only-extract', action='store_true', help='Only run extractor and write EXTRACT_*.json')
    parser.add_argument('--json-input', type=str, help='Skip extraction and generate report from this JSON file')
    parser.add_argument('--model-extract', type=str, default=os.getenv('OPENAI_MODEL_EXTRACT', 'gpt-5-mini'))
    parser.add_argument('--model-generate', type=str, default=os.getenv('OPENAI_MODEL_GENERATE', 'gpt-5-mini'))
    # New: mailbox/folder customization and discovery
    parser.add_argument('--mailbox', type=str, help='Outlook mailbox display name or address (overrides OUTLOOK_MAILBOX)')
    parser.add_argument('--folder', type=str, help='Folder path under mailbox, e.g. 收件箱/2. Policy Update (overrides OUTLOOK_FOLDER_PATH)')
    parser.add_argument('--list-mailboxes', action='store_true', help='List available Outlook mailboxes and exit')
    parser.add_argument('--list-folders', action='store_true', help='List folders under the mailbox and exit')
    parser.add_argument('--folders-depth', type=int, default=2, help='Depth for --list-folders (default: 2)')
    args = parser.parse_args()

    # Discovery commands
    if args.list_mailboxes:
        list_outlook_mailboxes()
        return
    if args.list_folders:
        mailbox = args.mailbox or default_mailbox
        list_outlook_folders(mailbox, max_depth=max(1, args.folders_depth))
        return

    # Resolve mailbox and folder path to use
    mailbox = args.mailbox or default_mailbox
    folder_path = _split_folder_path(args.folder) if args.folder else default_folder_path

    now = datetime.datetime.now()

    def parse_ymd(s: str) -> datetime.datetime:
        for fmt in ('%Y-%m-%d', '%Y/%m/%d', '%Y.%m.%d'):
            try:
                return datetime.datetime.strptime(s, fmt)
            except ValueError:
                continue
        LOG.error("invalid date format for '%s'; expected YYYY-MM-DD", s)
        sys.exit(1)

    if args.until:
        until = parse_ymd(args.until).replace(hour=23, minute=59, second=59, microsecond=999999)
    else:
        until = now

    if args.since:
        since = parse_ymd(args.since).replace(hour=0, minute=0, second=0, microsecond=0)
    else:
        since = until - datetime.timedelta(days=args.days)

    since = to_naive_local(since)
    until = to_naive_local(until)
    if since > until:
        LOG.warning('since > until; swapping values')
        since, until = until, since
    now = until

    LOG.info("fetching range: %s -> %s", since.isoformat(), until.isoformat())
    base_out = pathlib.Path(__file__).resolve().parent
    outdir = base_out / (now.strftime('%Y%m%d') + '_policy')
    outdir.mkdir(parents=True, exist_ok=True)
    LOG.info('saving files to %s', outdir)

    if USE_MOCK:
        sample = outdir / 'mock_email_1.txt'
        mid_dt = to_naive_local(since + (until - since) / 2)
        sample.write_text(
            f"Subject: MOCK\nFrom: tester\nReceived: {mid_dt}\n\nThis is mock policy email content for testing.",
            encoding='utf-8'
        )
        saved_emails = [{
            'txt': str(sample),
            'attachments': [],
            'received': mid_dt,
            'message_id': 'mock-1',
            'date_source': 'mock',
            'raw_received': mid_dt
        }]
        LOG.info('USE_MOCK_EMAILS enabled: created mock email %s', sample)
    else:
        if win32com is None:
            LOG.error('pywin32 (win32com) not available. Install pywin32 and run on Windows with Outlook configured.')
            sys.exit(2)
        try:
            folder = get_outlook_folder(mailbox, folder_path)
        except Exception as e:
            LOG.exception('failed to open outlook folder: %s', e)
            sys.exit(3)

        items = restrict_items_range(folder, since, now)
        try:
            LOG.info('Outlook items collection reports Count=%s', getattr(items, 'Count', 'N/A'))
        except Exception:
            pass

        saved_emails = []
        count = getattr(items, 'Count', 0) or 0
        for i in range(1, count + 1):
            try:
                msg = items.Item(i)
            except Exception:
                continue
            try:
                raw_received = to_naive_local(getattr(msg, 'ReceivedTime', None))
                if raw_received is None:
                    continue
                if raw_received < since or raw_received > now:
                    LOG.info('skipping message with ReceivedTime out of window: %s', raw_received)
                    continue
                txt, atts, rec_local, mid, chosen_date, date_src = save_mail_and_attachments(msg, outdir)
                saved_emails.append({
                    'txt': txt,
                    'attachments': atts,
                    'received': chosen_date,
                    'message_id': mid,
                    'date_source': date_src,
                    'raw_received': rec_local,
                    'chosen_date': chosen_date,
                })
                LOG.info('saved mail %s (canonical=%s raw_received=%s id=%s src=%s)', txt, chosen_date, rec_local, mid, date_src)
            except Exception:
                LOG.exception('failed to process an Outlook item')
                continue

        if not saved_emails:
            LOG.info('no messages found in the range %s to %s', since, now)
            return

    # Dedupe by message_id
    unique = {}
    deduped = []
    for rec in saved_emails:
        mid = rec.get('message_id')
        if not mid:
            deduped.append(rec)
            continue
        if mid in unique:
            LOG.info('dropping duplicate message id=%s', mid)
            continue
        unique[mid] = True
        deduped.append(rec)
    saved_emails = deduped

    # Sort by received
    saved_emails.sort(key=lambda r: r.get('received') or since)

    # Compute date range label
    dates = [r.get('received') for r in saved_emails if r.get('received')]
    if dates:
        start_dt = max(min(dates), since)
        end_dt = min(max(dates), until)
        start = start_dt.strftime('%y%m%d')
        end = end_dt.strftime('%y%m%d')
    else:
        start = since.strftime('%y%m%d')
        end = until.strftime('%y%m%d')

    report_basename = f"Weekly_report_{start}-{end}"
    report_path = outdir / (report_basename + '.md')

    # Build merged payload
    full_text = build_package_from_saved_emails(saved_emails)
    payload_default = outdir / f"ALL_{start}-{end}.txt"
    payload_default.write_text(full_text, encoding='utf-8')
    LOG.info('merged payload saved to %s', payload_default)

    # Optional dump
    dump_path = args.dump_payload or os.getenv('DUMP_PAYLOAD_PATH')
    if dump_path:
        dump_path = pathlib.Path(dump_path)
        dump_path.parent.mkdir(parents=True, exist_ok=True)
        dump_path.write_text(full_text, encoding='utf-8')
        LOG.info('merged payload saved to %s', dump_path)
        if args.only_dump:
            LOG.info('--only-dump specified; skipping OpenAI call')
            return

    # Skip OpenAI if requested
    if os.getenv('SKIP_OPENAI', '').lower() == 'true':
        LOG.info('SKIP_OPENAI=true: skipping OpenAI call and writing placeholder report')
        report_path.write_text('SKIPPED: OpenAI call was bypassed for testing.\n', encoding='utf-8')
        LOG.info('weekly report (placeholder) saved to %s', report_path)
        return

    openai_key = os.getenv('OPENAI_API_KEY')
    if not openai_key:
        LOG.error('OPENAI_API_KEY not found in environment. Put it in .env or env vars.')
        sys.exit(4)

    detach_openai = os.getenv('DETACH_OPENAI', '').lower() == 'true'
    period_label = f"{since.date()} — {until.date()}"
    extract_path = outdir / f"EXTRACT_{start}-{end}.json"

    # Step 2 only from existing JSON
    if args.json_input:
        try:
            structured_json_text = pathlib.Path(args.json_input).read_text(encoding='utf-8')
        except Exception:
            LOG.exception('failed to read --json-input')
            sys.exit(5)

        def _gen_only():
            md = call_openai_generate_markdown(structured_json_text, period_label, openai_key, model=args.model_generate)
            report_path.write_text(md, encoding='utf-8')
            LOG.info('weekly report saved to %s', report_path)

        if detach_openai:
            LOG.info('DETACH_OPENAI=true: background generating report from provided JSON')
            threading.Thread(target=_gen_only, daemon=True).start()
            return
        _gen_only()
        return

    # Step 1: Extract JSON
    try:
        structured_json_text = call_openai_extract_updates(full_text, file_label=payload_default.name, openai_api_key=openai_key, model=args.model_extract)
    except Exception:
        LOG.exception('extractor call failed')
        sys.exit(6)
    try:
        parsed = json.loads(structured_json_text)
        extract_path.write_text(json.dumps(parsed, ensure_ascii=False, indent=2), encoding='utf-8')
        LOG.info('extracted JSON saved to %s', extract_path)
    except Exception:
        extract_path.write_text(structured_json_text, encoding='utf-8')
        LOG.warning('extracted content is not valid JSON; raw saved to %s', extract_path)

    if args.only_extract:
        LOG.info('--only-extract specified; stopping after Step 1')
        return

    # Step 2: Generate Markdown report
    def _run_two_step_generate_sync():
        try:
            sj = extract_path.read_text(encoding='utf-8')
            md = call_openai_generate_markdown(sj, period_label, openai_key, model=args.model_generate)
            report_path.write_text(md, encoding='utf-8')
            LOG.info('weekly report (Markdown) saved to %s', report_path)
        except Exception:
            LOG.exception('report generation failed')

    if detach_openai:
        LOG.info('DETACH_OPENAI=true: spawning background report generation (two-step)')
        threading.Thread(target=_run_two_step_generate_sync, daemon=True).start()
        return

    async def _run_two_step_generate_async():
        return await asyncio.to_thread(_run_two_step_generate_sync)

    try:
        asyncio.run(_wait_with_progress(_run_two_step_generate_async, interval=3))
    except Exception:
        LOG.exception('failed while generating report')
        sys.exit(7)


if __name__ == '__main__':
    main()
