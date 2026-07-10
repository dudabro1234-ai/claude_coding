# -*- coding: utf-8 -*-
"""① Outlook 메일 수집기.

작업지시서 §3 ①:
- Outlook COM(pywin32)으로 지정 폴더의 미처리 메일을 추출 (본문 + 첨부 메타)
- 원본 절대 미변경: 읽음처리 X, 이동 X, 삭제 X (C3)
- 출력: work/inbox/{mail_id}.json

주의: 이 모듈은 Windows + Outlook 데스크톱 앱 환경에서만 동작한다.
"""
import hashlib
import json
import logging
import os

log = logging.getLogger(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INBOX_DIR = os.path.join(BASE_DIR, "work", "inbox")


def mail_id_from_entry_id(entry_id):
    """Outlook EntryID로부터 파일명으로 안전한 짧은 고유 ID를 만든다."""
    return hashlib.sha1(entry_id.encode("utf-8")).hexdigest()[:16]


def _resolve_folder(namespace, mailbox_name, folder_path):
    """'받은편지함/ESG요청' 형식의 경로를 Outlook 폴더 객체로 해석한다."""
    import win32com.client  # noqa: F401  (Windows 전용)
    parts = [p for p in folder_path.replace("\\", "/").split("/") if p]
    if not parts:
        raise ValueError("config.json의 outlook.folder_path가 비어 있습니다.")

    if mailbox_name:
        root = namespace.Folders[mailbox_name]
    else:
        # 기본 계정의 받은편지함 기준 (6 = olFolderInbox)
        inbox = namespace.GetDefaultFolder(6)
        if parts[0] in ("받은편지함", "Inbox", inbox.Name):
            parts = parts[1:]
        folder = inbox
        for name in parts:
            folder = folder.Folders[name]
        return folder

    folder = root
    for name in parts:
        folder = folder.Folders[name]
    return folder


def _sender_address(item):
    """SMTP 발신 주소를 최대한 정확히 얻는다 (Exchange 계정 대응)."""
    try:
        if item.SenderEmailType == "EX":
            exch = item.Sender.GetExchangeUser()
            if exch is not None:
                return exch.PrimarySmtpAddress
        return item.SenderEmailAddress or ""
    except Exception:
        try:
            return item.SenderEmailAddress or ""
        except Exception:
            return ""


def collect(config, processed_ids, limit=None):
    """지정 폴더에서 미처리 메일을 수집해 work/inbox/에 JSON으로 저장한다.

    반환: 새로 저장된 mail_id 목록.
    원본 메일은 어떤 속성도 변경하지 않는다 (UnRead 상태 유지, 이동/삭제 없음).
    """
    try:
        import win32com.client
    except ImportError:
        raise RuntimeError(
            "pywin32가 설치되어 있지 않습니다. Windows 환경에서 "
            "'pip install pywin32' 후 실행하세요.")

    os.makedirs(INBOX_DIR, exist_ok=True)
    ol_cfg = config.get("outlook", {})
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    folder = _resolve_folder(namespace, ol_cfg.get("mailbox", ""),
                             ol_cfg.get("folder_path", "받은편지함"))

    items = folder.Items
    items.Sort("[ReceivedTime]", True)  # 최신순 (읽기 전용 정렬 뷰)

    max_mails = limit or ol_cfg.get("max_mails_per_run", 100)
    new_ids = []
    for item in items:
        if len(new_ids) >= max_mails:
            break
        try:
            if getattr(item, "Class", None) != 43:  # 43 = olMail
                continue
            entry_id = item.EntryID
            mail_id = mail_id_from_entry_id(entry_id)
            if mail_id in processed_ids:
                continue

            attachments = []
            try:
                for att in item.Attachments:
                    name = att.FileName or ""
                    ext = os.path.splitext(name)[1].lstrip(".").lower()
                    attachments.append({"filename": name, "type": ext})
            except Exception:
                pass

            record = {
                "mail_id": mail_id,
                "entry_id": entry_id,
                "store_id": getattr(folder, "StoreID", ""),
                "received_at": item.ReceivedTime.strftime("%Y-%m-%dT%H:%M:%S"),
                "sender": _sender_address(item),
                "sender_name": getattr(item, "SenderName", ""),
                "subject": item.Subject or "",
                "body": item.Body or "",
                "attachments": attachments,
            }
            out_path = os.path.join(INBOX_DIR, f"{mail_id}.json")
            with open(out_path, "w", encoding="utf-8") as f:
                json.dump(record, f, ensure_ascii=False, indent=2)
            new_ids.append(mail_id)
            log.info("메일 수집: %s | %s", mail_id, record["subject"][:60])
        except Exception as e:
            log.warning("메일 1건 수집 실패(건너뜀): %s", e)
            continue

    log.info("수집 완료: 신규 %d건", len(new_ids))
    return new_ids
