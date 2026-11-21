import importlib.util
import base64
import hashlib
import json
import sqlite3
import sys
from email.message import EmailMessage
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
MAIN = ROOT / "scripts" / "faturas_shadow_77.py"


def load_main():
    spec = importlib.util.spec_from_file_location("faturas_shadow_77", MAIN)
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


def test_database_integrity_and_schema(tmp_path):
    main = load_main()
    db = tmp_path / "test.db"
    main.BASE_DIR = tmp_path
    main.DB_FILE = db
    main.inicializar_base_dados()
    conn = sqlite3.connect(db)
    try:
        assert conn.execute("PRAGMA integrity_check").fetchone()[0] == "ok"
        assert conn.execute("SELECT count(*) FROM calendar_events").fetchone()[0] == 0
    finally:
        conn.close()


def test_totp_is_valid_for_current_and_adjacent_windows():
    main = load_main()
    secret = "JBSWY3DPEHPK3PXP"
    code = main._totp(secret)
    assert main._totp_valido(secret, code)


def test_sha256_is_deterministic(tmp_path):
    main = load_main()
    target = tmp_path / "sample.txt"
    target.write_bytes(b"test data")
    first = main.sha256_ficheiro(target)
    second = main.sha256_ficheiro(target)
    assert len(first) == 64
    assert first == second


def test_email_trash_requires_a_backup_and_preserves_starred(tmp_path):
    main = load_main()
    assert main.APAGAR_EMAILS is True
    assert main.APAGAR_NEWSLETTERS is False
    main.EMAIL_BACKUP_DIR = tmp_path
    message_id = "message-id"
    backup = tmp_path / f"{message_id}.eml"
    backup.write_bytes(b"complete raw message")
    (tmp_path / f"{message_id}.json").write_text(
        json.dumps({"sha256": hashlib.sha256(backup.read_bytes()).hexdigest()}),
        encoding="utf-8",
    )

    class Request:
        def execute(self):
            return {"labelIds": ["INBOX", "STARRED"]}

    class Messages:
        def get(self, **kwargs):
            return Request()

        def users(self):
            return Users()

        def trash(self, **kwargs):
            raise AssertionError("Email com estrela não pode ser enviado para o lixo")

    class Users:
        def messages(self):
            return Messages()

    class Service:
        def users(self):
            return Users()

    assert main.enviar_email_para_lixo(Service(), message_id) is False


def test_email_cleanup_backs_up_starred_and_aborts_all_deletion_on_backup_error(tmp_path, monkeypatch):
    main = load_main()
    main.EMAIL_BACKUP_DIR = tmp_path
    content = b"starred full message"
    raw = base64.urlsafe_b64encode(content).decode("ascii").rstrip("=")
    trashed = []

    class Request:
        def __init__(self, value):
            self.value = value

        def execute(self):
            return self.value

    class Messages:
        def list(self, **kwargs):
            assert kwargs["q"] == "in:anywhere"
            return Request({"messages": [{"id": "starred"}, {"id": "broken"}]})

        def get(self, **kwargs):
            if kwargs["id"] == "starred":
                return Request({"raw": raw, "labelIds": ["STARRED"]})
            return Request({"labelIds": ["INBOX"]})

        def trash(self, **kwargs):
            trashed.append(kwargs["id"])
            return Request({})

    class Users:
        def messages(self):
            return Messages()

    class Service:
        def users(self):
            return Users()

    monkeypatch.setattr(main, "build", lambda *args, **kwargs: Service())
    monkeypatch.setattr(main, "info", lambda *args, **kwargs: None)
    monkeypatch.setattr(main, "aviso", lambda *args, **kwargs: None)
    monkeypatch.setattr(main, "progresso", lambda *args, **kwargs: None)
    monkeypatch.setattr(main, "registar_log", lambda *args, **kwargs: None)

    main.arquivar_e_apagar_emails_sem_estrela(object())

    backup = tmp_path / "starred.eml"
    metadata = json.loads((tmp_path / "starred.json").read_text(encoding="utf-8"))
    assert backup.read_bytes() == content
    assert metadata["sha256"] == hashlib.sha256(content).hexdigest()
    assert trashed == []


def test_payment_date_label_is_parsed():
    main = load_main()
    assert main.datas_de_transacao("Data/Hora 02-03-2025 12:47:49") == [(2025, 3, 2)]


def test_payment_date_does_not_depend_on_file_name():
    main = load_main()
    text = "Estado Executado\nData/Hora 02-03-2025 12:47:49\nEntidade 12345\nReferência 123 456 789"
    assert main.datas_de_transacao(text) == [(2025, 3, 2)]


def test_localized_menu_has_automatic_fallback():
    main = load_main()
    assert callable(main.t)
    assert main.t("processar")


def test_google_calendar_reminder_is_created_and_recorded(tmp_path, monkeypatch):
    main = load_main()
    main.BASE_DIR = tmp_path
    main.DB_FILE = tmp_path / "test.db"
    main.INTEGRACAO_FILE = tmp_path / "integracao.json"
    main.INTEGRACAO_FILE.write_text(
        json.dumps({"provider": "google", "email": "google-account"}),
        encoding="utf-8",
    )
    main.inicializar_base_dados()
    with sqlite3.connect(main.DB_FILE) as conn:
        conn.execute(
            "INSERT INTO faturas(emissor, tipo_doc, valor, data_limite, estado, ficheiro) "
            "VALUES (?, ?, ?, ?, ?, ?)",
            ("Fornecedor", "Fatura", 42.5, "2026-12-10", "Por pagar", "invoice.pdf"),
        )

    created = []

    class Request:
        def execute(self):
            return {"id": "event-123"}

    class Events:
        def insert(self, **kwargs):
            created.append(kwargs)
            return Request()

    class CalendarService:
        def events(self):
            return Events()

    monkeypatch.setattr(main, "build", lambda *args, **kwargs: CalendarService())
    monkeypatch.setattr(main, "info", lambda *args, **kwargs: None)
    monkeypatch.setattr(main, "aviso", lambda *args, **kwargs: None)

    main.sincronizar_lembretes_calendario(object())

    assert len(created) == 1
    assert created[0]["body"]["start"] == {"date": "2026-12-10"}
    with sqlite3.connect(main.DB_FILE) as conn:
        assert conn.execute(
            "SELECT provider, event_id FROM calendar_events WHERE invoice_id=1"
        ).fetchone() == ("google", "event-123")


def test_icloud_downloads_pdf_and_keeps_mail(tmp_path, monkeypatch):
    main = load_main()
    main.BASE_DIR = tmp_path
    main.DB_FILE = tmp_path / "test.db"
    main.ENTRADA_DIR = tmp_path / "entrada"
    main.ENTRADA_DIR.mkdir()
    main.inicializar_base_dados()
    message = EmailMessage()
    message["Date"] = "Mon, 02 Mar 2025 12:00:00 +0000"
    message["Subject"] = "Invoice"
    message.set_content("Attached")
    message.add_attachment(
        b"%PDF-1.4 test",
        maintype="application",
        subtype="pdf",
        filename="invoice.pdf",
    )
    raw_message = message.as_bytes()
    seen_flags = []

    class Mail:
        def __enter__(self):
            return self

        def __exit__(self, *args):
            return None

        def login(self, *_):
            return "OK", []

        def select(self, *_):
            return "OK", []

        def uid(self, command, *args):
            if command == "search":
                return "OK", [b"17"]
            if command == "fetch":
                return "OK", [(b"RFC822", raw_message)]
            if command == "store":
                seen_flags.append((args[0], args[-1]))
                return "OK", []
            raise AssertionError(f"Unexpected IMAP command: {command}")

        def logout(self):
            return "BYE", []

    monkeypatch.setattr(main, "ler_integracao", lambda: {"provider": "icloud", "email": "user@example.test"})
    monkeypatch.setattr(main.imaplib, "IMAP4_SSL", lambda *args, **kwargs: Mail())
    monkeypatch.setattr(main, "progresso", lambda *args, **kwargs: None)
    monkeypatch.setattr(main, "info", lambda *args, **kwargs: None)
    monkeypatch.setattr(main, "aviso", lambda *args, **kwargs: None)
    monkeypatch.setattr(main, "registar_log", lambda *args, **kwargs: None)
    monkeypatch.setattr(main, "stat", lambda *args, **kwargs: None)

    import keyring

    monkeypatch.setattr(keyring, "get_password", lambda *_: "app-password")
    main.descarregar_anexos_icloud()

    downloaded = list(main.ENTRADA_DIR.glob("*.pdf"))
    assert len(downloaded) == 1
    assert downloaded[0].read_bytes() == b"%PDF-1.4 test"
    assert seen_flags == [(b"17", "(\\Seen)")]
    with sqlite3.connect(main.DB_FILE) as conn:
        assert conn.execute(
            "SELECT msg_id, pdf, data_email FROM emails_pendentes"
        ).fetchone() == ("icloud:17", downloaded[0].name, "2025-03-02")


def test_icloud_calendar_creates_idempotent_invoice_reminder(tmp_path, monkeypatch):
    main = load_main()
    main.BASE_DIR = tmp_path
    main.DB_FILE = tmp_path / "test.db"
    main.INTEGRACAO_FILE = tmp_path / "integracao.json"
    main.INTEGRACAO_FILE.write_text(
        json.dumps({"provider": "icloud", "email": "user@example.test"}),
        encoding="utf-8",
    )
    main.inicializar_base_dados()
    with sqlite3.connect(main.DB_FILE) as conn:
        conn.execute(
            "INSERT INTO faturas(emissor, tipo_doc, valor, data_limite, estado, ficheiro) "
            "VALUES (?, ?, ?, ?, ?, ?)",
            ("Fornecedor", "Fatura", 42.5, "2026-12-10", "Por pagar", "invoice.pdf"),
        )

    saved_events = []

    class Calendar:
        def event_by_uid(self, _):
            from caldav.lib.error import NotFoundError

            raise NotFoundError("not found")

        def add_event(self, *, ical):
            saved_events.append(ical)

    calendar = Calendar()

    class Principal:
        def calendars(self):
            return [calendar]

    class Client:
        def __init__(self, **kwargs):
            assert kwargs["url"] == "https://caldav.icloud.com"

        def principal(self):
            return Principal()

    import caldav
    import keyring

    monkeypatch.setattr(caldav, "DAVClient", Client)
    monkeypatch.setattr(keyring, "get_password", lambda *_: "app-password")
    monkeypatch.setattr(main, "info", lambda *args, **kwargs: None)
    monkeypatch.setattr(main, "aviso", lambda *args, **kwargs: None)
    monkeypatch.setattr(main, "registar_log", lambda *args, **kwargs: None)

    main.sincronizar_lembretes_calendario(None)

    assert len(saved_events) == 1
    assert "faturas-shadow-77-1@local" in saved_events[0]
    assert "Fatura por pagar" in saved_events[0]
    with sqlite3.connect(main.DB_FILE) as conn:
        assert conn.execute(
            "SELECT provider, event_id FROM calendar_events WHERE invoice_id=1"
        ).fetchone() == ("icloud", "faturas-shadow-77-1@local")
