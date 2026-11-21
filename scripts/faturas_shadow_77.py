"""
FATURAS SHADOW_77 v9.7

Novo na v9.7:
  * Regras por nome de ficheiro (config/regras_nome.json, editável): "Detalhe_516671702_*" = Via Verde (fatura),
    "Extrato_MM_AAAA" = Cetelem (extrato), "COMUNICADO..." = informativo.
  * PDFs protegidos: as palavras-passe ficam em config/pdf_passwords.txt (menu Gerir registos, opção 8).
  * Gerir registos, opção 9: aplica as regras aos registos existentes e reprocessa a quarentena (2FA).
  * Faturas da Via Verde por pagar são confirmadas nos extratos do Millennium BCP (linha "Via Verde" + valor).
  * Datas no nome do ficheiro: AAAA_MM, MM_AAAA e "18 de Julho de 2026".

Novo na v9.6:
  * Separação por ANO (arquivo/<ano>/<mês>/<emissor>) pela data do documento: QR, "data de emissão", data no
    nome do ficheiro, data limite ou data do email.
  * Tudo o que for anterior a 2026 (ANO_VERIFICACAO) fica "Ok" e NÃO é verificado: sem quarentena, sem
    "por pagar", sem cruzamento. Só se verifica 2026 e os anos seguintes. Se houver QUALQUER data de 2026
    ou posterior no documento, ele é verificado normalmente.

Novo na v9.5:
  * Leitura do QR code das faturas portuguesas (AT): NIF do emissor, valor, data, nº do documento e ATCUD.
  * Comprovativos: Pagamento (entidade + referência + montante), Transferência (duas contas, sem entidade
    nem referência) e Pagamento ao Estado (só referência). Comprovativo incompleto vai para quarentena.
  * Imagens e vídeos: pré-processamento antes do OCR e, nos vídeos, votação entre vários frames.
  * Repetidos apanhados à entrada: mesmo conteúdo (hash), mesmo QR (NIF + ATCUD) ou mesma
    entidade/referência/valor com a mesma data.
  * Menu Gerir registos: importar as correções feitas no Excel para a base de dados (2FA).

Novo na v9.4:
  * Pasta whatsapp: imagens e vídeos são convertidos em PDF e analisados como as restantes faturas.
  * Cruzamento automático faturas <-> comprovativos (Pago / Pago (a confirmar) / Por pagar) + relatório.
  * Modo AUTO (menu opção 2, ou:  python faturas_shadow_77.py --auto [--cada MINUTOS]) sem qualquer intervenção.

Alterações da v9.3 face à v9.1/v9.2:
  * Emails só vão para o lixo DEPOIS de todos os PDFs desse email serem validados e arquivados.
  * Faturas nunca são tratadas como newsletter (a verificação de PDF vem primeiro).
  * 2FA real (PIN com hash + código TOTP de app autenticadora) ligada a alterar/apagar registos.
    O PIN já não está no código nem no ecrã.
  * Gmail com paginação e leitura recursiva de anexos.
  * Valores com milhares (5.801,49), deteção de emissores por palavra inteira,
    hiperlinks do Excel corrigidos e Excel reconstruído a partir da base de dados.
  * Menu principal + modo automático:  python faturas_shadow_77.py --auto
"""
import base64
import getpass
import hashlib
import hmac
import imaplib
import json
import logging
import locale
import os
import random
import re
import shutil
import sqlite3
import struct
import sys
import threading
import time
import urllib.parse
import warnings
from datetime import date, datetime, timedelta
from email import message_from_bytes
from email.policy import default as email_policy
from email.utils import parsedate_to_datetime
from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path

import openpyxl
import pdfplumber
import pytesseract
import qrcode
from openpyxl.styles import Alignment, Font, PatternFill
from PIL import Image, ImageFilter, ImageOps, ImageStat
from pypdf import PdfReader

try:
    import cv2  # vídeos e leitura do QR code (pip install opencv-python)
    import numpy as np
except ImportError:
    cv2 = None
    np = None

# Silenciar avisos técnicos
warnings.filterwarnings("ignore")
logging.getLogger("pypdf").setLevel(logging.ERROR)
logging.getLogger("pdfminer").setLevel(logging.ERROR)

from google.auth.exceptions import RefreshError
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

os.system('')  # ativa cores ANSI na consola do Windows
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass


def _idioma_windows():
    """Seleciona o idioma da interface a partir da linguagem da interface do Windows."""
    if os.name == "nt":
        try:
            import ctypes
            user_ui_language = ctypes.windll.kernel32.GetUserDefaultUILanguage()
            nome = locale.windows_locale.get(user_ui_language)
            if nome:
                return nome.split("_")[0].lower()
        except Exception:
            pass
    return (locale.getlocale()[0] or locale.getdefaultlocale()[0] or "pt").split("_")[0].lower()


IDIOMA = _idioma_windows()
IDIOMAS_BASE = {
    "pt": {
        "sair": "Sair",
        "processar": "Processar faturas (Gmail + WhatsApp)",
        "automatico": "AUTO - processo completo sem intervenção",
        "gestao": "Gerir registos",
        "pagamentos": "Estado dos pagamentos",
        "email_limpeza": "Fazer backup e mover emails sem estrela para o Lixo",
        "opcao": "Opção",
        "cancelado": "Operação cancelada; nenhum email foi alterado.",
        "confirmar_email": "Confirma? Primeiro será feito backup de todos os emails; os sem estrela vão para o Lixo.",
        "escreve": "Escreve",
        "design": "design by Shadow_77",
    },
    "en": {
        "sair": "Exit",
        "processar": "Process invoices (Gmail + WhatsApp)",
        "automatico": "AUTO - complete process without interaction",
        "gestao": "Manage records",
        "pagamentos": "Payment status",
        "email_limpeza": "Back up and move non-starred emails to Trash",
        "opcao": "Option",
        "cancelado": "Operation cancelled; no email was changed.",
        "confirmar_email": "Confirm? All emails will be backed up first; non-starred emails will go to Trash.",
        "escreve": "Type",
        "design": "design by Shadow_77",
    },
    "fr": {
        "sair": "Quitter",
        "processar": "Traiter les factures (Gmail + WhatsApp)",
        "automatico": "AUTO - processus complet sans intervention",
        "gestao": "Gérer les enregistrements",
        "pagamentos": "État des paiements",
        "email_limpeza": "Sauvegarder et mettre les e-mails sans étoile à la corbeille",
        "opcao": "Option",
        "cancelado": "Opération annulée; aucun e-mail n'a été modifié.",
        "confirmar_email": "Confirmer ? Tous les e-mails seront sauvegardés; ceux sans étoile iront à la corbeille.",
        "escreve": "Tapez",
        "design": "design by Shadow_77",
    },
    "es": {
        "sair": "Salir",
        "processar": "Procesar facturas (Gmail + WhatsApp)",
        "automatico": "AUTO - proceso completo sin intervención",
        "gestao": "Gestionar registros",
        "pagamentos": "Estado de pagos",
        "email_limpeza": "Copia y mueve al correo no destacado a la papelera",
        "opcao": "Opción",
        "cancelado": "Operación cancelada; no se modificó ningún correo.",
        "confirmar_email": "¿Confirmar? Se copiarán todos los correos; los no destacados irán a la papelera.",
        "escreve": "Escribe",
        "design": "design by Shadow_77",
    },
    "de": {
        "sair": "Beenden",
        "processar": "Rechnungen verarbeiten (Gmail + WhatsApp)",
        "automatico": "AUTO - vollständiger Prozess ohne Eingriff",
        "gestao": "Datensätze verwalten",
        "pagamentos": "Zahlungsstatus",
        "email_limpeza": "Sicherung erstellen und nicht markierte E-Mails in den Papierkorb verschieben",
        "opcao": "Option",
        "cancelado": "Vorgang abgebrochen; keine E-Mail wurde geändert.",
        "confirmar_email": "Bestätigen? Zuerst werden alle E-Mails gesichert; nicht markierte E-Mails gehen in den Papierkorb.",
        "escreve": "Eingeben",
        "design": "design by Shadow_77",
    },
}


def t(chave):
    return IDIOMAS_BASE.get(IDIOMA, IDIOMAS_BASE["en"]).get(
        chave, IDIOMAS_BASE["en"].get(chave, chave)
    )

# Cores ANSI
R = "\033[91m"
G = "\033[92m"
Y = "\033[93m"
B = "\033[94m"
C = "\033[96m"
W = "\033[97m"
GRAY = "\033[90m"
BOLD = "\033[1m"
RESET = "\033[0m"

# ---------------------------------------------------------------------------
# CONFIGURAÇÃO
# ---------------------------------------------------------------------------
TESSERACT_CMD = os.environ.get("TESSERACT_CMD", r"C:\Program Files\Tesseract-OCR\tesseract.exe")
if os.path.exists(TESSERACT_CMD):
    pytesseract.pytesseract.tesseract_cmd = TESSERACT_CMD


def verificar_tesseract():
    """Confirma que o OCR está disponível antes de iniciar a auditoria."""
    executavel = pytesseract.pytesseract.tesseract_cmd
    disponivel = os.path.isfile(executavel) or shutil.which(executavel) is not None
    if not disponivel:
        aviso(
            f"{R}      -> Tesseract não encontrado. Instala-o ou define TESSERACT_CMD "
            "para o executável real.{RESET}"
        )
        registar_log(
            f"Tesseract indisponível: executável='{executavel}'. "
            "Defina TESSERACT_CMD ou instale Tesseract-OCR."
        )
        return False
    return True

_DADOS_POR_UTILIZADOR = Path(
    os.environ.get(
        "LOCALAPPDATA",
        str(Path.home() / "AppData" / "Local"),
    )
) / "Faturas Shadow_77"
BASE_DIR = Path(os.environ.get("FATURAS_SHADOW_77_DIR", str(_DADOS_POR_UTILIZADOR)))
ENTRADA_DIR = BASE_DIR / "entrada"
WHATSAPP_DIR = BASE_DIR / "whatsapp"
ARQUIVO_DIR = BASE_DIR / "arquivo"
POR_IDENTIFICAR_DIR = BASE_DIR / "por identificar"
LOGS_DIR = BASE_DIR / "logs"
CONFIG_DIR = _DADOS_POR_UTILIZADOR / "config"
EXCEL_FILE = BASE_DIR / "Faturas_2026.xlsx"
DB_FILE = BASE_DIR / "Faturas Shadow_77.db"
SEGURANCA_FILE = CONFIG_DIR / "seguranca.json"
EMAIL_BACKUP_DIR = BASE_DIR / "emails_backup"
INTEGRACAO_FILE = CONFIG_DIR / "integracao.json"
KEYRING_SERVICE = "Faturas Shadow_77"

APAGAR_NEWSLETTERS = False  # Limpeza de newsletters não faz parte do processamento automático
APAGAR_EMAILS = True        # Autorizado apenas na ação manual, após cópia integral verificada


def enviar_email_para_lixo(service, msg_id):
    """Envia para o lixo apenas emails sem estrela que tenham uma cópia local verificada."""
    if not APAGAR_EMAILS:
        registar_log(f"Proteção ativa: email {msg_id} não foi enviado para o lixo.")
        return False
    copia = EMAIL_BACKUP_DIR / f"{msg_id}.eml"
    metadados = EMAIL_BACKUP_DIR / f"{msg_id}.json"
    if not copia.is_file() or not metadados.is_file():
        registar_log(f"Email {msg_id} mantido: cópia local integral não encontrada.")
        return False
    try:
        info_copia = json.loads(metadados.read_text(encoding="utf-8"))
        digest = hashlib.sha256(copia.read_bytes()).hexdigest()
    except (OSError, ValueError) as e:
        registar_log(f"Email {msg_id} mantido: erro ao verificar cópia local ({e}).")
        return False
    if digest != info_copia.get("sha256"):
        registar_log(f"Email {msg_id} mantido: SHA-256 da cópia local não corresponde.")
        return False
    msg = service.users().messages().get(userId="me", id=msg_id, format="minimal").execute()
    if "STARRED" in msg.get("labelIds", []):
        registar_log(f"Email {msg_id} protegido: tem estrela e não foi enviado para o lixo.")
        return False
    if "TRASHED" in msg.get("labelIds", []):
        return True
    service.users().messages().trash(userId='me', id=msg_id).execute()
    return True

SCOPES = [
    'https://www.googleapis.com/auth/gmail.modify',
    'https://www.googleapis.com/auth/calendar.events',
]

COLUNAS = [
    "Categoria", "Emissor", "Tipo Doc", "Entidade MB", "Referência MB",
    "Valor", "Data limite", "Estado", "Origem", "Ficheiro", "Notas", "Data doc"
]

ANO_VERIFICACAO = int(os.environ.get("FATURAS_SHADOW_77_ANO", "2026"))
ESTADOS_VALIDOS = ["Pago", "Pago (a confirmar)", "Por pagar", "Informativo", "Por Identificar", "Ordenado", "Ok"]
ESTADOS_PENDENTES = ("Por pagar", "Por Identificar", "Pago (a confirmar)")   # ainda exigem atenção
ESTADOS_PARA_OK = ESTADOS_PENDENTES + ("Informativo",)                    # nos anos antigos passam a "Ok" ("Pago" fica Pago)

EXT_IMAGENS = {".jpg", ".jpeg", ".png", ".webp", ".bmp", ".tif", ".tiff"}
EXT_VIDEOS = {".mp4", ".mov", ".avi", ".mkv", ".3gp"}
MIN_LADO_IMAGEM = 600   # imagens mais pequenas (stickers, ícones, miniaturas) são ignoradas

TIPOS_A_PAGAR = ("Fatura", "Aviso de Cobrança / Condomínio")
EMISSOR_COMPROVATIVO = "Comprovativo (por associar)"
EMISSORES_GENERICOS = {EMISSOR_COMPROVATIVO, "Desconhecido", "MB WAY / SIBS", "CGD", "Millennium BCP", "Novo Banco"}

FILL_VERDE = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
FONT_VERDE = Font(color="006100", bold=True)
FILL_VERMELHO = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
FONT_VERMELHO = Font(color="9C0006", bold=True)
FILL_AZUL = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
FONT_AZUL = Font(color="1F497D", bold=True)
FILL_AMARELO = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
FONT_AMARELO = Font(color="9C5700", bold=True)
FILL_LARANJA = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
FONT_LARANJA = Font(color="C65911", bold=True)
FILL_CINZA = PatternFill(start_color="EDEDED", end_color="EDEDED", fill_type="solid")
FONT_CINZA = Font(color="595959")


class SilenciadorStderr:
    def __enter__(self):
        self._original_stderr = sys.stderr
        sys.stderr = open(os.devnull, 'w')

    def __exit__(self, exc_type, exc_val, exc_tb):
        sys.stderr.close()
        sys.stderr = self._original_stderr


# ---------------------------------------------------------------------------
# LOG
# ---------------------------------------------------------------------------
def registar_log(mensagem):
    try:
        LOGS_DIR.mkdir(parents=True, exist_ok=True)
        log_file = LOGS_DIR / f"Faturas_Shadow_77_{datetime.now().strftime('%Y-%m')}.log"
        with open(log_file, "a", encoding="utf-8") as f:
            f.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {mensagem}\n")
    except Exception:
        pass


# ---------------------------------------------------------------------------
# ESTADO DA EXECUÇÃO: mensagens, estatísticas e progresso
# (durante a animação os passos ficam em segundo plano; só o resumo final aparece)
# ---------------------------------------------------------------------------
SILENCIOSO = False
AVISOS = []
STATS = {}
PESOS_FASES = {"gmail": 0.20, "media": 0.15, "auditoria": 0.45, "cruzamento": 0.05, "calendario": 0.05, "excel": 0.10}
ORDEM_FASES = list(PESOS_FASES)
PROG = {"pct": 0.0, "concluido": False}


def _sem_ansi(texto):
    return re.sub(r"\x1b\[[0-9;?]*[A-Za-z]", "", texto)


def info(msg=""):
    """Mensagem normal: só aparece fora do modo animado."""
    if not SILENCIOSO:
        print(msg)


def aviso(msg):
    """Aviso/erro: fica guardado, vai para o log e aparece no resumo final."""
    limpo = _sem_ansi(msg).strip()
    AVISOS.append(limpo)
    registar_log(f"AVISO: {limpo}")
    if not SILENCIOSO:
        print(msg)


def stat(chave, n=1):
    STATS[chave] = STATS.get(chave, 0) + n


def resetar_estado():
    AVISOS.clear()
    STATS.clear()
    PROG["pct"] = 0.0
    PROG["concluido"] = False


def progresso(fase, frac=1.0):
    antes = sum(PESOS_FASES[f] for f in ORDEM_FASES[:ORDEM_FASES.index(fase)])
    pct = 100 * (antes + PESOS_FASES[fase] * max(0.0, min(1.0, frac)))
    PROG["pct"] = max(PROG["pct"], pct)


# ---------------------------------------------------------------------------
# BASE DE DADOS
# ---------------------------------------------------------------------------
def ligar_bd():
    return sqlite3.connect(DB_FILE, timeout=30)


def inicializar_base_dados():
    BASE_DIR.mkdir(parents=True, exist_ok=True)
    conn = ligar_bd()
    try:
        conn.execute('''
            CREATE TABLE IF NOT EXISTS faturas (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                categoria TEXT,
                emissor TEXT,
                tipo_doc TEXT,
                entidade_mb TEXT,
                referencia_mb TEXT,
                valor REAL,
                data_limite TEXT,
                estado TEXT,
                origem TEXT,
                ficheiro TEXT UNIQUE,
                notas TEXT,
                data_registo TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        # Emails cujos PDFs foram descarregados mas ainda não validados.
        # O email só vai para o lixo quando TODOS os seus PDFs estiverem ok/duplicado.
        conn.execute('''
            CREATE TABLE IF NOT EXISTS emails_pendentes (
                msg_id TEXT,
                pdf TEXT,
                resultado TEXT,
                PRIMARY KEY (msg_id, pdf)
            )
        ''')
        conn.execute('''
            CREATE TABLE IF NOT EXISTS calendar_events (
                invoice_id INTEGER PRIMARY KEY,
                provider TEXT NOT NULL,
                event_id TEXT NOT NULL,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        colunas = [r[1] for r in conn.execute("PRAGMA table_info(faturas)")]
        for col, tipo in (("ligado_a", "INTEGER"),          # fatura <-> comprovativo
                          ("hash_ficheiro", "TEXT"),        # SHA-256 do PDF (repetidos à entrada)
                          ("nif_emissor", "TEXT"), ("atcud", "TEXT"), ("num_doc", "TEXT"),   # QR code da AT
                          ("data_doc", "TEXT")):            # data do documento (QR), AAAA-MM-DD
            if col not in colunas:
                conn.execute(f"ALTER TABLE faturas ADD COLUMN {col} {tipo}")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_faturas_hash ON faturas(hash_ficheiro)")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_faturas_estado ON faturas(estado)")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_faturas_emissor ON faturas(emissor)")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_faturas_referencia ON faturas(referencia_mb)")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_faturas_ligado_a ON faturas(ligado_a)")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_faturas_data_doc ON faturas(data_doc)")
        if "data_email" not in [r[1] for r in conn.execute("PRAGMA table_info(emails_pendentes)")]:
            conn.execute("ALTER TABLE emails_pendentes ADD COLUMN data_email TEXT")
        conn.commit()
    finally:
        conn.close()


def registar_bd(dados):
    conn = ligar_bd()
    try:
        conn.execute("""
            INSERT OR IGNORE INTO faturas
            (categoria, emissor, tipo_doc, entidade_mb, referencia_mb, valor, data_limite, estado, origem, ficheiro, notas,
             hash_ficheiro, nif_emissor, atcud, num_doc, data_doc)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            dados["Categoria"], dados["Emissor"], dados["Tipo Doc"],
            dados["Entidade MB"], dados["Referência MB"], dados["Valor"],
            dados["Data limite"], dados["Estado"], dados["Origem"],
            dados["Ficheiro"], dados["Notas"],
            dados.get("Hash"), dados.get("NIF emissor"), dados.get("ATCUD"), dados.get("Num doc"), dados.get("Data doc")
        ))
        conn.commit()
    except Exception as e:
        registar_log(f"Erro ao inserir na BD: {e}")
    finally:
        conn.close()


def sha256_ficheiro(caminho):
    h = hashlib.sha256()
    with open(caminho, "rb") as f:
        for bloco in iter(lambda: f.read(1024 * 1024), b""):
            h.update(bloco)
    return h.hexdigest()


def hash_ja_registado(h):
    conn = ligar_bd()
    try:
        r = conn.execute("SELECT id FROM faturas WHERE hash_ficheiro = ? LIMIT 1", (h,)).fetchone()
        return r[0] if r else None
    finally:
        conn.close()


def preencher_hashes_em_falta():
    """Calcula o hash dos ficheiros já registados (de versões antigas) para os repetidos serem apanhados à entrada."""
    conn = ligar_bd()
    feitos = 0
    try:
        for rid, rel in conn.execute("SELECT id, ficheiro FROM faturas WHERE hash_ficheiro IS NULL").fetchall():
            p = caminho_do_registo(rel or "")
            if p is not None and p.exists():
                try:
                    conn.execute("UPDATE faturas SET hash_ficheiro = ? WHERE id = ?", (sha256_ficheiro(p), rid))
                    feitos += 1
                except OSError:
                    pass
        conn.commit()
    finally:
        conn.close()
    return feitos


def registar_pendente(msg_id, pdf_nome, data_email=None):
    conn = ligar_bd()
    try:
        conn.execute("INSERT OR IGNORE INTO emails_pendentes (msg_id, pdf, resultado, data_email) VALUES (?, ?, NULL, ?)",
                     (msg_id, pdf_nome, data_email))
        conn.commit()
    finally:
        conn.close()


def data_email_do_pdf(pdf_nome):
    conn = ligar_bd()
    try:
        r = conn.execute("SELECT data_email FROM emails_pendentes WHERE pdf = ? LIMIT 1", (pdf_nome,)).fetchone()
        return r[0] if r else None
    finally:
        conn.close()


def registar_resultado_pdf(pdf_nome, resultado):
    conn = ligar_bd()
    try:
        conn.execute("UPDATE emails_pendentes SET resultado = ? WHERE pdf = ?", (resultado, pdf_nome))
        conn.commit()
    finally:
        conn.close()


# ---------------------------------------------------------------------------
# SEGURANÇA: PIN (hash) + TOTP (app autenticadora)
# ---------------------------------------------------------------------------
def _hash_pin(pin, salt):
    return hashlib.pbkdf2_hmac("sha256", pin.encode("utf-8"), salt, 200_000).hex()


def _totp(segredo_b32, momento=None, passo=30, digitos=6):
    chave = base64.b32decode(segredo_b32, casefold=True)
    contador = int((time.time() if momento is None else momento) // passo)
    h = hmac.new(chave, struct.pack(">Q", contador), hashlib.sha1).digest()
    o = h[-1] & 0x0F
    codigo = (struct.unpack(">I", h[o:o + 4])[0] & 0x7FFFFFFF) % (10 ** digitos)
    return str(codigo).zfill(digitos)


def _totp_valido(segredo_b32, codigo):
    codigo = (codigo or "").strip().replace(" ", "")
    agora = time.time()
    return any(hmac.compare_digest(_totp(segredo_b32, agora + d * 30), codigo) for d in (-1, 0, 1))


def _carregar_seguranca():
    if not SEGURANCA_FILE.exists():
        return None
    try:
        return json.loads(SEGURANCA_FILE.read_text(encoding="utf-8"))
    except Exception:
        return None


def configurar_seguranca():
    print(f"\n{C}{BOLD}=== CONFIGURAR SEGURANÇA (PIN + APP AUTENTICADORA) ==={RESET}")
    pin = getpass.getpass("Define o PIN (mínimo 4 dígitos): ").strip()
    if not pin.isdigit() or len(pin) < 6:
        print(f"{R}✗ O PIN tem de ter pelo menos 6 dígitos.{RESET}")
        return False
    if getpass.getpass("Repete o PIN: ").strip() != pin:
        print(f"{R}✗ Os PINs não coincidem.{RESET}")
        return False

    segredo = base64.b32encode(os.urandom(10)).decode()
    uri = f"otpauth://totp/Faturas%20Shadow_77:Shadow_77?secret={segredo}&issuer=Faturas%20Shadow_77"
    print(f"\n{W}Lê este QR code na tua app autenticadora (Google Authenticator, Authy, ...):{RESET}")
    qr = qrcode.QRCode(box_size=1, border=2)
    qr.add_data(uri)
    qr.print_ascii(invert=True)
    print(f"\nOu introduz a chave manualmente: {Y}{segredo}{RESET}\n")

    if not _totp_valido(segredo, input("Código de 6 dígitos que a app mostra agora: ")):
        print(f"{R}✗ Código inválido. Segurança NÃO configurada.{RESET}")
        return False

    salt = os.urandom(16)
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    SEGURANCA_FILE.write_text(json.dumps({
        "salt": salt.hex(), "pin": _hash_pin(pin, salt), "totp": segredo
    }), encoding="utf-8")
    try:
        os.chmod(SEGURANCA_FILE, 0o600)
    except Exception:
        pass
    registar_log("Segurança (PIN + TOTP) configurada.")
    print(f"{G}✓ Segurança configurada.{RESET}")
    return True


def verificar_dupla_autenticacao(acao_desc):
    print(f"\n{R}{BOLD}=== SEGURANÇA: DUPLA AUTENTICAÇÃO REQUERIDA ==={RESET}")
    print(f"Ação sensível: {W}{acao_desc}{RESET}")

    cfg = _carregar_seguranca()
    if cfg is None:
        print(f"{Y}A segurança ainda não está configurada. Vamos configurá-la agora.{RESET}")
        configurar_seguranca()
        print(f"{Y}Repete agora a operação.{RESET}")
        return False

    pin_ok = False
    for _ in range(3):
        pin = getpass.getpass("PIN: ").strip()
        if hmac.compare_digest(_hash_pin(pin, bytes.fromhex(cfg["salt"])), cfg["pin"]):
            pin_ok = True
            break
        print(f"{R}✗ PIN incorreto.{RESET}")
    if not pin_ok:
        registar_log(f"2FA falhou (PIN) para: {acao_desc}")
        print(f"{R}Operação cancelada.{RESET}")
        return False

    for _ in range(3):
        if _totp_valido(cfg["totp"], input("Código da app autenticadora (6 dígitos): ")):
            registar_log(f"2FA ok para: {acao_desc}")
            print(f"{G}✓ Dupla autenticação bem-sucedida.{RESET}\n")
            return True
        print(f"{R}✗ Código inválido.{RESET}")
    registar_log(f"2FA falhou (TOTP) para: {acao_desc}")
    print(f"{R}Operação cancelada.{RESET}")
    return False


# ---------------------------------------------------------------------------
# ANIMAÇÕES
# ---------------------------------------------------------------------------
def animacao_abertura(limpar=True):
    """Apresentação inicial com o logótipo integrado."""
    if limpar:
        os.system('cls' if os.name == 'nt' else 'clear')
    largura = shutil.get_terminal_size((80, 20)).columns
    try:
        app_dir = Path(
            getattr(sys, "_MEIPASS", Path(__file__).resolve().parents[1])
        )
        logo = app_dir / "assets" / "shadow_77_logo.jpg"
        if not logo.exists():
            logo = BASE_DIR / "assets" / "shadow_77_logo.jpg"
        if logo.exists():
            imagem = Image.open(logo).convert("RGB")
            imagem = ImageOps.fit(imagem, (min(48, largura - 10), min(24, imagem.height * 48 // imagem.width)), method=Image.Resampling.LANCZOS)
            altura = imagem.height
            if altura % 2:
                altura += 1
                imagem = ImageOps.pad(imagem, (imagem.width, altura), color=(8, 14, 34))
            for y in range(0, altura, 2):
                linha = " " * max(0, (largura - imagem.width) // 2)
                for x in range(imagem.width):
                    r1, g1, b1 = imagem.getpixel((x, y))
                    r2, g2, b2 = imagem.getpixel((x, y + 1))
                    linha += f"\033[38;2;{r1};{g1};{b1}m\033[48;2;{r2};{g2};{b2}m▀"
                print(linha + RESET)
            print()
    except Exception as e:
        registar_log(f"Logótipo de arranque não pôde ser mostrado: {e}")
    linhas = [
        ("=" * 40, C),
        ("FATURAS SHADOW_77 v9.7", C + BOLD),
        (t("design"), W),
        ("=" * 40, C),
    ]
    print()
    for texto, cor in linhas:
        print(" " * max(0, (largura - len(texto)) // 2) + f"{cor}{texto}{RESET}")
    print()



def _grelha(largura, altura, linhas):
    g = [["."] * largura for _ in range(altura)]
    for y, spans in linhas.items():
        for a, b in spans:
            for x in range(a, b + 1):
                g[y][x] = "#"
    return ["".join(r) for r in g]

def _meios_blocos(grelha):
    if len(grelha) % 2:
        grelha = grelha + ["." * len(grelha[0])]
    mapa = {(True, True): "█", (True, False): "▀", (False, True): "▄", (False, False): " "}
    return ["".join(mapa[(t == "#", b == "#")] for t, b in zip(grelha[i], grelha[i + 1]))
            for i in range(0, len(grelha), 2)]

_DINO_CORPO = {
    0: [(10, 17)], 1: [(9, 18)], 2: [(9, 10), (12, 18)], 3: [(9, 18)], 4: [(9, 18)],
    5: [(9, 13)], 6: [(9, 17)], 7: [(0, 0), (8, 18)], 8: [(0, 0), (7, 18)],
    9: [(0, 1), (6, 18)], 10: [(0, 2), (5, 18), (19, 20)], 11: [(1, 3), (4, 18), (20, 21)],
    12: [(3, 17)],
}
def _dino(perna_esq, perna_dir):
    linhas = dict(_DINO_CORPO)
    for y, spans in ((13, [(6, 9)]), (14, [(6, 8)]), (15, [(6, 10)])) if perna_esq == "p" else \
                    ((13, [(6, 9)]), (14, [(7, 9)])):
        linhas[y] = linhas.get(y, []) + spans
    for y, spans in ((13, [(12, 15)]), (14, [(12, 14)]), (15, [(12, 16)])) if perna_dir == "p" else \
                    ((13, [(12, 15)]), (14, [(13, 15)])):
        linhas[y] = linhas.get(y, []) + spans
    return _meios_blocos(_grelha(22, 16, linhas))

DINO = {"a": _dino("p", "l"), "b": _dino("l", "p"), "j": _dino("p", "p")}
DINO_L, DINO_A = 22, 8

def _cacto(altura):
    base = {0: [(3, 5)], 1: [(3, 5)], 2: [(0, 1), (3, 5), (7, 8)], 3: [(0, 1), (3, 5), (7, 8)],
            4: [(0, 1), (3, 5), (7, 8)], 5: [(0, 5), (7, 8)], 6: [(3, 8)]}
    linhas = {}
    extra = altura - 10
    for y, sp in base.items():
        linhas[y + extra] = sp
    for y in range(7 + extra, altura):
        linhas[y] = [(3, 5)]
    for y in range(0, extra):
        linhas[y] = [(3, 5)]
    return _meios_blocos(_grelha(9, altura, linhas))

CACTOS = [_cacto(10), _cacto(8), _cacto(12)]
def _duplo(a, b):
    ha, hb = len(a), len(b)
    h = max(ha, hb)
    return [(" " * 9 if h - ha > i else a[i - (h - ha)]) + " " + (" " * 9 if h - hb > i else b[i - (h - hb)])
            for i in range(h)]
CACTOS.append(_duplo(CACTOS[1], CACTOS[0]))

PASSARO = [
    ["  ▄▄    ", "▀▀████▄▄", "   ▀▀▀  "],
    ["   ▄▄   ", " ▄████▄▄", "▀▀  ▀▀  "],
]
NUVEM = [" ▄▄▄    ", "▄█████▄▄", "▀▀▀▀▀▀▀▀"]

# ---- cores
RST = RESET; VERDE = G; VERDE_ESC = "\033[32m"; CIANO = C; CINZA = GRAY; AMARELO = Y

class CenaDino:
    SALTO = 6
    VEL = 2.0

    def __init__(self, largura):
        self.W = largura
        self.H = DINO_A + self.SALTO + 2
        self.dx = 8
        self.tick = 0
        self.obst = []            # [x, sprite]
        self.passaros = []        # [x, y, tick0]
        self.nuvens = [[random.randint(0, largura), random.randint(0, 3)] for _ in range(3)]
        self.rng = random.Random()
        self.salto_t = None       # ticks desde o início do salto
        self.salto_T = 0
        self.prox_gap = self.rng.randint(30, 46)
        self.chao = "".join(self.rng.choice("▁▁▁▁▂▁▁▃▁▁▁▂▁▁▁▁▁▃▁▁") for _ in range(400))

    def passo(self):
        self.tick += 1
        for n in self.nuvens:
            n[0] -= 0.35
            if n[0] < -10:
                n[0] = self.W + self.rng.randint(2, 20); n[1] = self.rng.randint(0, 3)
        for o in self.obst:
            o[0] -= self.VEL
        self.obst = [o for o in self.obst if o[0] + len(o[1][0]) > -2]
        if not self.obst or self.W - (self.obst[-1][0] + len(self.obst[-1][1][0])) >= self.prox_gap:
            self.obst.append([float(self.W), self.rng.choice(CACTOS)])
            self.prox_gap = self.rng.randint(30, 46)
        for p in self.passaros:
            p[0] -= 2.6
        self.passaros = [p for p in self.passaros if p[0] > -10]
        if not self.passaros and self.rng.random() < 0.02:
            self.passaros.append([float(self.W), self.rng.randint(0, 2), self.tick])
        # salto automático
        if self.salto_t is None:
            for o in self.obst:
                frente = self.dx + DINO_L
                if o[0] + len(o[1][0]) > self.dx and o[0] - frente <= self.VEL * 3:
                    self.salto_T = int((DINO_L + len(o[1][0])) / self.VEL) + 7
                    self.salto_t = 0
                    break
        else:
            self.salto_t += 1
            if self.salto_t >= self.salto_T:
                self.salto_t = None

    def _altura(self):
        if self.salto_t is None:
            return 0
        t = self.salto_t / self.salto_T
        return round(self.SALTO * 4 * t * (1 - t))

    def linhas(self):
        W, H = self.W, self.H
        c = [[(" ", "") for _ in range(W)] for _ in range(H)]
        def blit(spr, x, y, cor):
            for r, linha in enumerate(spr):
                yy = y + r
                if not 0 <= yy < H: continue
                for k, ch in enumerate(linha):
                    xx = int(x) + k
                    if ch != " " and 0 <= xx < W:
                        c[yy][xx] = (ch, cor)
        for x, y in self.nuvens:
            blit(NUVEM, x, y, CINZA)
        for x, y, t0 in self.passaros:
            blit(PASSARO[((self.tick - t0) // 3) % 2], x, y, AMARELO)
        for x in range(W):                          # chão
            c[H - 2][x] = ("▔", CINZA)
            c[H - 1][x] = (self.chao[(x + int(self.tick * self.VEL)) % len(self.chao)], CINZA)
        for x, spr in self.obst:
            blit(spr, x, H - 2 - len(spr), VERDE_ESC)
        dy = self._altura()
        pose = "j" if dy > 0 else ("a" if (self.tick // 2) % 2 == 0 else "b")
        blit(DINO[pose], self.dx, H - 2 - DINO_A - dy, VERDE)
        out = []
        for linha in c:
            s, cor_atual = "", ""
            for ch, cor in linha:
                if cor != cor_atual:
                    s += (cor if cor else RST); cor_atual = cor
                s += ch
            out.append(s + RST)
        return out

def barra(pct, largura):
    frac = max(0.0, min(1.0, pct / 100))
    total = frac * largura
    cheio = int(total)
    parcial = " ▏▎▍▌▋▊▉"[int((total - cheio) * 8)] if cheio < largura else ""
    resto = largura - cheio - (1 if parcial.strip() else 0)
    cor = VERDE if pct >= 99.5 else CIANO
    return f"{cor}{'█' * cheio}{parcial if parcial.strip() else ''}{CINZA}{'░' * max(0, resto)}{RST}"


def animacao_dino(prog):
    """Dino do Chrome a correr e a saltar cactos enquanto o processamento corre em segundo plano."""
    colunas, linhas_term = shutil.get_terminal_size((100, 30))
    largura = max(40, min(colunas - 6, 78))
    cena = CenaDino(largura)
    n_linhas = cena.H + 2
    compacto = linhas_term < n_linhas + 8
    if compacto:
        n_linhas = 1
    pad = " " * ((colunas - largura) // 2)
    larg_barra = max(10, min(44, largura - 18))
    pad_barra = " " * max(0, (colunas - (larg_barra + 18)) // 2)

    mostrado, ticks = 0.0, 0
    sys.stdout.write("\033[?25l")
    try:
        while True:
            ticks += 1
            alvo = 100.0 if prog["concluido"] else min(99.0, prog["pct"])
            mostrado = min(alvo, mostrado + max(0.4, (alvo - mostrado) * 0.12))
            pct = int(mostrado)

            quadro = ""
            if not compacto:
                cena.passo()
                for linha in cena.linhas():
                    quadro += f"\033[2K{pad}{linha}\n"
                quadro += "\033[2K\n"
            quadro += f"\033[2K{pad_barra}{W}Progresso{RESET}  [{barra(mostrado, larg_barra)}]  {Y}{pct:3d}%{RESET}\n"
            sys.stdout.write(quadro)
            sys.stdout.flush()

            if prog["concluido"] and mostrado >= 100.0 and ticks >= 25:
                break
            sys.stdout.write(f"\033[{n_linhas}A\r")
            time.sleep(0.055)
    finally:
        sys.stdout.write("\033[?25h\n")
        sys.stdout.flush()


# ---------------------------------------------------------------------------
# GMAIL
# ---------------------------------------------------------------------------
# Se alguma destas palavras aparecer no assunto/remetente/excerto, o email NUNCA é tratado como newsletter.
PALAVRAS_PROTEGIDAS = [
    "fatura", "factura", "invoice", "recibo", "comprovativo", "pagamento", "referência", "referencia",
    "vencimento", "multibanco", "débito direto", "debito direto", "aviso de cobrança", "condomínio",
    "condominio", "apólice", "apolice", "seguro", "extrato", "extracto", "nota de crédito",
]


def _cabecalhos(payload):
    return {h.get('name', '').lower(): h.get('value', '').lower() for h in payload.get('headers', [])}


def e_newsletter(payload, snippet=""):
    """Só devolve True para emails SEM indícios de fatura. É chamado apenas quando o email não tem PDF."""
    h = _cabecalhos(payload)
    texto = f"{h.get('subject', '')} {h.get('from', '')} {snippet}".lower()
    if any(p in texto for p in PALAVRAS_PROTEGIDAS):
        return False
    return ("list-unsubscribe" in h
            or h.get("precedence", "") in ("bulk", "list")
            or "newsletter" in texto)


def _iterar_partes(parte):
    yield parte
    for sub in parte.get("parts", []) or []:
        yield from _iterar_partes(sub)


def _anexos_pdf(payload):
    return [p for p in _iterar_partes(payload)
            if p.get("filename") and p["filename"].lower().endswith(".pdf")
            and p.get("body", {}).get("attachmentId")]


def _nome_seguro(nome):
    return re.sub(r'[\\/*?:"<>|\r\n]', "_", nome).strip() or "documento.pdf"


def _listar_ids(service, query):
    ids, token = [], None
    while True:
        resp = service.users().messages().list(
            userId='me', q=query, pageToken=token, maxResults=100
        ).execute()
        ids.extend(m['id'] for m in resp.get('messages', []))
        token = resp.get('nextPageToken')
        if not token:
            return ids


def _listar_ids_nao_lidos(service):
    return _listar_ids(service, "is:unread")


def descarregar_anexos_icloud():
    """Descarrega anexos PDF não lidos de iCloud Mail, sem apagar mensagens."""
    conta = ler_integracao()
    if not conta or conta["provider"] != "icloud":
        return
    try:
        import keyring
        password = keyring.get_password(KEYRING_SERVICE, conta["email"])
        if not password:
            raise RuntimeError("A palavra-passe iCloud não está disponível no gestor seguro do Windows.")
        with imaplib.IMAP4_SSL("imap.mail.me.com", 993, timeout=30) as mail:
            mail.login(conta["email"], password)
            status, _ = mail.select("INBOX")
            if status != "OK":
                raise RuntimeError("Não foi possível abrir a caixa de entrada iCloud.")
            status, resultado = mail.uid("search", None, "UNSEEN")
            if status != "OK":
                raise RuntimeError("Não foi possível pesquisar mensagens iCloud não lidas.")
            uids = resultado[0].split()
            descarregados = 0
            for i, uid_bytes in enumerate(uids, 1):
                uid = uid_bytes.decode("ascii")
                msg_id = f"icloud:{uid}"
                try:
                    status, partes = mail.uid("fetch", uid_bytes, "(RFC822)")
                    if status != "OK":
                        raise RuntimeError("Não foi possível descarregar a mensagem.")
                    raw = next(item[1] for item in partes if isinstance(item, tuple))
                    mensagem = message_from_bytes(raw, policy=email_policy)
                    try:
                        data_email = parsedate_to_datetime(mensagem["Date"]).date().isoformat()
                    except (TypeError, ValueError, OverflowError):
                        data_email = None
                    guardados = []
                    for parte in mensagem.iter_attachments():
                        nome = parte.get_filename() or ""
                        if not nome.lower().endswith(".pdf"):
                            continue
                        conteudo = parte.get_payload(decode=True)
                        if not conteudo:
                            raise ValueError(f"Anexo vazio ou ilegível: {nome}")
                        nome_local = _nome_seguro(f"{uid}_{nome}")
                        destino = ENTRADA_DIR / nome_local
                        if destino.exists() and hashlib.sha256(destino.read_bytes()).digest() != hashlib.sha256(conteudo).digest():
                            nome_local = _nome_seguro(f"{uid}_{random.randrange(100000, 999999)}_{nome}")
                            destino = ENTRADA_DIR / nome_local
                        destino.write_bytes(conteudo)
                        guardados.append(nome_local)
                    for nome_local in guardados:
                        registar_pendente(msg_id, nome_local, data_email)
                    if guardados:
                        mail.uid("store", uid_bytes, "+FLAGS", "(\\Seen)")
                        descarregados += len(guardados)
                except Exception as e:
                    registar_log(f"Erro na mensagem iCloud {uid}: {type(e).__name__}: {e}")
                progresso("gmail", i / max(1, len(uids)))
            mail.logout()
        stat("pdfs_gmail", descarregados)
        info(f"      {G}-> {descarregados} PDF(s) descarregado(s) do iCloud; emails mantidos.{RESET}")
    except Exception as e:
        registar_log(f"Erro no iCloud Mail: {type(e).__name__}: {e}")
        aviso(f"{R}      -> Erro ao aceder ao iCloud Mail (ver log).{RESET}")


def arquivar_e_apagar_emails_sem_estrela(creds):
    """Guarda cada email completo localmente e só depois envia os não estrelados para o lixo."""
    if not creds:
        info(f"{Y}      -> Gmail ignorado (sem autenticação).{RESET}")
        return

    EMAIL_BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    service = build("gmail", "v1", credentials=creds)
    ids = _listar_ids(service, "in:anywhere")
    arquivados, protegidos, erros = [], set(), []

    for i, msg_id in enumerate(ids, 1):
        try:
            msg = service.users().messages().get(
                userId="me", id=msg_id, format="raw"
            ).execute()
            if "STARRED" in msg.get("labelIds", []):
                protegidos.add(msg_id)
            raw = msg.get("raw")
            if not raw:
                raise ValueError("Gmail não devolveu a mensagem completa em formato RAW")
            conteudo = base64.urlsafe_b64decode(raw + "=" * (-len(raw) % 4))
            caminho = EMAIL_BACKUP_DIR / f"{msg_id}.eml"
            meta_path = EMAIL_BACKUP_DIR / f"{msg_id}.json"
            digest = hashlib.sha256(conteudo).hexdigest()

            if caminho.exists():
                existente = hashlib.sha256(caminho.read_bytes()).hexdigest()
                if existente != digest:
                    raise OSError("já existe uma cópia local diferente para este ID")
            else:
                temporario = caminho.with_suffix(".eml.tmp")
                with temporario.open("wb") as f:
                    f.write(conteudo)
                    f.flush()
                    os.fsync(f.fileno())
                os.replace(temporario, caminho)

            metadata = {
                "message_id": msg_id,
                "thread_id": msg.get("threadId"),
                "internal_date": msg.get("internalDate"),
                "labels": msg.get("labelIds", []),
                "size_estimate": msg.get("sizeEstimate"),
                "sha256": digest,
                "archived_at": datetime.now().isoformat(timespec="seconds"),
            }
            meta_tmp = meta_path.with_suffix(".json.tmp")
            meta_tmp.write_text(json.dumps(metadata, ensure_ascii=False), encoding="utf-8")
            os.replace(meta_tmp, meta_path)
            if hashlib.sha256(caminho.read_bytes()).hexdigest() != digest:
                raise OSError("a verificação SHA-256 da cópia local falhou")
            arquivados.append(msg_id)
        except Exception as e:
            erros.append(msg_id)
            registar_log(f"Erro ao copiar email {msg_id}; mantido no Gmail: {type(e).__name__}: {e}")
        progresso("emails", i / max(1, len(ids)))

    if erros:
        aviso(
            f"{R}      -> {len(erros)} email(s) não foram copiados; nenhum email será enviado "
            f"para o lixo nesta execução.{RESET}"
        )
        return

    apagados = 0
    for msg_id in arquivados:
        if msg_id in protegidos:
            continue
        try:
            if enviar_email_para_lixo(service, msg_id):
                apagados += 1
                conn = ligar_bd()
                try:
                    conn.execute("DELETE FROM emails_pendentes WHERE msg_id = ?", (msg_id,))
                    conn.commit()
                finally:
                    conn.close()
        except Exception as e:
            registar_log(f"Erro ao enviar email {msg_id} para o lixo: {type(e).__name__}: {e}")

    stat("emails_apagados", apagados)
    stat("emails_retidos", len(ids) - apagados)
    info(f"      {G}• Emails copiados e enviados para o lixo: {apagados}{RESET}")
    info(f"      {Y}• Emails com estrela preservados: {len(protegidos)}{RESET}")


def descarregar_e_gerir_emails_gmail(creds):
    conta = ler_integracao()
    if conta and conta["provider"] == "icloud":
        descarregar_anexos_icloud()
        return
    if not creds:
        info(f"{Y}[2/6] Gmail ignorado (sem autenticação).{RESET}")
        return

    info(f"{B}[2/6] {W}A varrer o Gmail (descarrega PDFs; o email só é apagado depois de validado)...{RESET}")
    ENTRADA_DIR.mkdir(parents=True, exist_ok=True)
    descarregados = eliminados_news = erros = 0

    try:
        service = build('gmail', 'v1', credentials=creds)
        ids = _listar_ids_nao_lidos(service)
        if not ids:
            info(f"      {W}-> Nenhum email por ler.{RESET}")
            return

        for i, msg_id in enumerate(ids, 1):
            try:
                msg = service.users().messages().get(userId='me', id=msg_id).execute()
                payload = msg.get('payload', {})
                anexos = _anexos_pdf(payload)

                if anexos:
                    guardados = []
                    for part in anexos:
                        anexo = service.users().messages().attachments().get(
                            userId='me', messageId=msg_id, id=part['body']['attachmentId']).execute()
                        dados = base64.urlsafe_b64decode(anexo['data'].encode('UTF-8'))
                        nome = f"{msg_id[:10]}_{_nome_seguro(part['filename'])}"
                        (ENTRADA_DIR / nome).write_bytes(dados)
                        guardados.append(nome)
                    try:
                        data_email = datetime.fromtimestamp(int(msg["internalDate"]) / 1000).date().isoformat()
                    except (KeyError, ValueError, TypeError, OSError):
                        data_email = None
                    for nome in guardados:
                        registar_pendente(msg_id, nome, data_email)
                    descarregados += len(guardados)
                    # Marca como lido para não voltar a descarregar. NÃO apaga o email aqui.
                    service.users().messages().modify(
                        userId='me', id=msg_id, body={'removeLabelIds': ['UNREAD']}).execute()

                elif APAGAR_NEWSLETTERS and e_newsletter(payload, msg.get('snippet', '')):
                    if enviar_email_para_lixo(service, msg_id):
                        eliminados_news += 1
                # Qualquer outro email sem PDF fica intacto.
            except HttpError as e:
                erros += 1
                registar_log(f"Erro no email {msg_id}: {e}")
            progresso("gmail", i / len(ids))

        stat("pdfs_gmail", descarregados)
        stat("newsletters", eliminados_news)
        info(f"      {G}-> {descarregados} PDF(s) descarregado(s).{RESET}")
        info(f"      {Y}-> {eliminados_news} newsletter(s) sem PDF enviada(s) para o lixo.{RESET}")
        if erros:
            aviso(f"{R}      -> {erros} email(s) do Gmail com erro (ver log).{RESET}")
    except HttpError as error:
        registar_log(f"Erro na API do Gmail: {error}")
        aviso(f"{R}      -> Erro na API do Gmail (ver log).{RESET}")


def finalizar_emails_gmail(creds):
    """Faz backup integral e só depois envia para o lixo os emails sem estrela."""
    arquivar_e_apagar_emails_sem_estrela(creds)


def ler_integracao():
    if not INTEGRACAO_FILE.exists():
        return None
    try:
        dados = json.loads(INTEGRACAO_FILE.read_text(encoding="utf-8"))
    except (OSError, ValueError) as e:
        raise RuntimeError(f"Não foi possível ler a configuração de conta {INTEGRACAO_FILE}: {e}") from e
    if dados.get("provider") not in ("google", "icloud") or not dados.get("email"):
        raise RuntimeError(f"Configuração de conta inválida: {INTEGRACAO_FILE}")
    return dados


def _guardar_integracao(provider, email):
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    temporario = INTEGRACAO_FILE.with_suffix(".json.tmp")
    temporario.write_text(
        json.dumps({"provider": provider, "email": email}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    os.replace(temporario, INTEGRACAO_FILE)


def configurar_integracao():
    """Configura a conta individual sem gravar palavras-passe na pasta do programa."""
    print("\n" + "=" * 58)
    print(" CONFIGURAÇÃO INICIAL — LIGA A TUA CONTA")
    print(" As credenciais ficam no perfil deste utilizador do Windows.")
    print("  1) Google (Gmail + Google Calendar)")
    print("  2) Apple iCloud (Mail + Calendar)")
    escolha = input("Escolhe 1 ou 2: ").strip()
    if escolha == "1":
        cred_path = Path(
            os.environ.get("FATURAS_GOOGLE_CREDENTIALS", str(CONFIG_DIR / "credentials.json"))
        )
        if not cred_path.is_file():
            aviso(
                f"{Y}      -> Para Google, coloca o ficheiro OAuth do tipo Desktop app em "
                f"{cred_path}. Não partilhes esse ficheiro.{RESET}"
            )
            return None
        creds = obter_credenciais_google(interativo=True)
        if creds:
            try:
                profile = build("gmail", "v1", credentials=creds).users().getProfile(
                    userId="me"
                ).execute()
            except HttpError as e:
                registar_log(f"Não foi possível confirmar a conta Google autenticada: {e}")
                aviso(f"{R}      -> O Google não confirmou a conta. Verifica se a Gmail API está ativa.{RESET}")
                return None
            email = profile.get("emailAddress")
            if not email:
                aviso(f"{R}      -> O Google não devolveu o endereço da conta autenticada.{RESET}")
                return None
            _guardar_integracao("google", email)
            info(f"{G}      -> Conta Google ligada: {email}.{RESET}")
            return "google"
        return None

    if escolha == "2":
        email = input("Apple ID (endereço de email): ").strip()
        if not email or "@" not in email:
            aviso(f"{R}      -> Endereço Apple ID inválido.{RESET}")
            return None
        palavra_passe = getpass.getpass(
            "Palavra-passe específica de app Apple (não uses a palavra-passe normal): "
        ).strip()
        if not palavra_passe:
            aviso(f"{R}      -> Palavra-passe específica de app vazia; configuração cancelada.{RESET}")
            return None
        try:
            with imaplib.IMAP4_SSL("imap.mail.me.com", 993, timeout=20) as mail:
                mail.login(email, palavra_passe)
                mail.logout()
            from caldav import DAVClient
            client = DAVClient(
                url="https://caldav.icloud.com",
                username=email,
                password=palavra_passe,
                timeout=20,
            )
            calendarios = client.principal().calendars()
            if not calendarios:
                raise RuntimeError("A conta iCloud não disponibilizou um calendário.")
            import keyring
            keyring.set_password(KEYRING_SERVICE, email, palavra_passe)
            _guardar_integracao("icloud", email)
        except Exception as e:
            aviso(f"{R}      -> Não foi possível validar Mail/Calendar iCloud: {e}{RESET}")
            return None
        info(f"{G}      -> Conta iCloud ligada: {email}.{RESET}")
        return "icloud"

    aviso(f"{Y}      -> Escolha inválida; a configuração inicial não foi concluída.{RESET}")
    return None


def obter_credenciais_google(interativo=True):
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    creds = None
    token_path = CONFIG_DIR / 'token.json'
    cred_path = Path(
        os.environ.get("FATURAS_GOOGLE_CREDENTIALS", str(CONFIG_DIR / "credentials.json"))
    )

    info(f"{B}[1/6] {W}A verificar subsistema de credenciais Google...{RESET}")
    if token_path.exists():
        try:
            creds = Credentials.from_authorized_user_file(str(token_path), SCOPES)
            if creds and not creds.has_scopes(SCOPES):
                creds = None
        except Exception as e:
            registar_log(f"token.json inválido: {e}")
            creds = None

    if creds and not creds.valid:
        if creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
                token_path.write_text(creds.to_json())
            except RefreshError:
                info(f"{Y}      -> Token expirado ou revogado. É preciso autenticar de novo.{RESET}")
                creds = None
        else:
            creds = None

    if creds and creds.valid:
        return creds

    if not interativo:
        aviso(f"{Y}      -> Gmail ignorado: é preciso autenticar de novo (corre a opção 1 do menu uma vez).{RESET}")
        registar_log("AUTO: Gmail ignorado (token inválido/expirado).")
        return None

    if not cred_path.exists():
        aviso(f"{R}      -> Falta o ficheiro {cred_path}.{RESET}")
        return None

    flow = InstalledAppFlow.from_client_secrets_file(str(cred_path), SCOPES)
    porta = 8080
    flow.redirect_uri = f"http://localhost:{porta}/"
    estado_oauth = base64.urlsafe_b64encode(os.urandom(24)).decode("ascii").rstrip("=")
    auth_url, _ = flow.authorization_url(
        prompt='select_account consent',
        access_type='offline',
        include_granted_scopes='true',
        state=estado_oauth,
    )

    print("\n" + "=" * 60)
    print(" LEIA ESTE QR CODE COM A CÂMARA DO TELEMÓVEL")
    print("=" * 60)
    qr = qrcode.QRCode(box_size=1, border=2)
    qr.add_data(auth_url)
    qr.print_ascii(invert=True)
    print(f"\nLink direto: {auth_url}\n" + "=" * 60 + "\n")

    codigo_recebido = {}

    class OAuthHandler(BaseHTTPRequestHandler):
        def do_GET(self):
            params = urllib.parse.parse_qs(urllib.parse.urlparse(self.path).query)
            if params.get('state') != [estado_oauth]:
                self.send_response(400)
                self.end_headers()
                return
            if 'code' in params:
                codigo_recebido['code'] = params['code'][0]
                self.send_response(200)
                self.send_header('Content-type', 'text/html; charset=utf-8')
                self.end_headers()
                self.wfile.write(b"<h1>Autenticacao concluida com sucesso!</h1>")
            else:
                self.send_response(400)
                self.end_headers()

        def log_message(self, format, *args):
            pass

    servidor = HTTPServer(('localhost', porta), OAuthHandler)
    servidor.timeout = 120
    try:
        for _ in range(5):
            servidor.handle_request()
            if 'code' in codigo_recebido:
                break
    finally:
        servidor.server_close()

    if 'code' not in codigo_recebido:
        aviso(f"{R}      -> Autenticação não concluída.{RESET}")
        return None

    flow.fetch_token(code=codigo_recebido['code'])
    creds = flow.credentials
    token_path.write_text(creds.to_json())
    return creds


# ---------------------------------------------------------------------------
# ANÁLISE DE DOCUMENTOS
# ---------------------------------------------------------------------------
def limpar_ref(ref):
    if not ref:
        return ""
    return re.sub(r'[\s\-\.]+', '', str(ref)).strip()


def contem(texto, chave):
    """Procura a chave como palavra inteira (evita 'digi' em 'digital', 'edp' em 'wedpoint', etc.)."""
    return re.search(r'(?<![a-z0-9à-ú])' + re.escape(chave) + r'(?![a-z0-9à-ú])', texto) is not None


RE_VALOR = re.compile(
    r"(?:Total|Valor|Montante|Importância|A Pagar|Valor total trimestre)[^\d]*"
    r"(\d{1,3}(?:[. ]\d{3})+[.,]\d{2}|\d+[.,]\d{2})\s*€?",
    re.IGNORECASE
)


def parse_valor(texto_valor):
    s = texto_valor.replace(" ", "")
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    else:
        s = s.replace(",", ".")
    return round(float(s), 2)


EMISSORES_CONHECIDOS = [
    (["condominio", "condomínio", "administração do condomínio"], "Condomínio", "Habitação / Condomínio"),
    (["ocidental", "ageas", "503454109"], "Ocidental Seguros", "Seguros"),
    (["domestic & general", "domestic and general"], "Domestic & General", "Seguros"),
    (["iberdrola", "502124083", "iberdrola clientes", "a78213805"], "Iberdrola", "Energia"),
    (["edp comercial", "edp s.a.", "edp", "503504145", "edp distribuicao"], "EDP Comercial", "Energia"),
    (["endesa", "508195840"], "Endesa", "Energia"),
    (["galp energia", "galp power", "galp", "500084478"], "Galp", "Energia / Combustível"),
    (["su eletricidade"], "SU Eletricidade", "Energia"),
    (["vodafone", "501413203", "10297"], "Vodafone", "Telecomunicações"),
    (["meo", "altice", "500079417", "pt com"], "MEO", "Telecomunicações"),
    (["nos comunicações", "nos lusomundo", "502784897", "zon", "21690"], "NOS", "Telecomunicações"),
    (["nowo", "cabovisao", "504381831"], "NOWO", "Telecomunicações"),
    (["digi", "517351608"], "DIGI", "Telecomunicações"),
    (["auchan", "sonae", "continente", "modelo continente", "502011475"], "Continente / Auchan", "Supermercado"),
    (["pingo doce", "jeronimo martins", "500100144"], "Pingo Doce", "Supermercado"),
    (["mercadona", "509173456"], "Mercadona", "Supermercado"),
    (["leroy merlin"], "Leroy Merlin", "Bricolage"),
    (["cgd", "caixa geral de depósitos", "500969643"], "CGD", "Banca / Serviços"),
    (["mb way", "sibs", "multibanco", "501579730"], "MB WAY / SIBS", "Pagamentos"),
    (["millennium bcp", "bcp", "501216664"], "Millennium BCP", "Banca / Serviços"),
    (["novo banco", "502398555"], "Novo Banco", "Banca / Serviços"),
    (["fidelidade", "500918880"], "Fidelidade", "Seguros"),
    (["autoridade tributária", "autoridade tributaria", "finanças", "republica portuguesa"], "Estado", "Estado / Impostos"),
    (["segurança social"], "Estado", "Estado / Segurança Social"),
    (["epal", "501062973"], "Águas de Portugal / EPAL", "Água"),
    (["via verde", "isps"], "Via Verde", "Transportes / Portagens"),
]


# ---------------------------------------------------------------------------
# REGRAS POR NOME DE FICHEIRO  (config/regras_nome.json: cria-se sozinho; podes acrescentar regras)
# ---------------------------------------------------------------------------
REGRAS_NOME_PADRAO = [
    {"padrao": r"^detalhe_516671702", "emissor": "Via Verde", "categoria": "Transportes / Portagens", "tipo": "Fatura"},
    {"padrao": r"^extrato_\d{2}_\d{4}", "emissor": "Cetelem", "categoria": "Crédito / Cartão", "tipo": "Extrato",
     "estado": "Informativo"},
    {"padrao": r"^comunicado", "tipo": "Diversos", "estado": "Informativo"},
]
_REGRAS_CACHE = []


def carregar_regras_nome(forcar=False):
    """Lista de (regex, regra). 'padrao' é uma expressão regular aplicada ao nome sem prefixos de hora/sequência."""
    if _REGRAS_CACHE and not forcar:
        return _REGRAS_CACHE[0]
    ficheiro = CONFIG_DIR / "regras_nome.json"
    regras = REGRAS_NOME_PADRAO
    try:
        if ficheiro.exists():
            lido = json.loads(ficheiro.read_text(encoding="utf-8"))
            if isinstance(lido, list):
                regras = lido
        else:
            CONFIG_DIR.mkdir(parents=True, exist_ok=True)
            ficheiro.write_text(json.dumps(REGRAS_NOME_PADRAO, ensure_ascii=False, indent=1), encoding="utf-8")
    except Exception as e:
        registar_log(f"regras_nome.json inválido: {e}")
    compiladas = []
    for r in regras:
        try:
            compiladas.append((re.compile(r["padrao"], re.IGNORECASE), r))
        except (re.error, KeyError, TypeError) as e:
            registar_log(f"Regra ignorada ({e}): {r}")
    _REGRAS_CACHE[:] = [compiladas]
    return compiladas


def nome_base_para_regras(nome):
    n = Path(str(nome)).name
    n = re.sub(r"^\d{6,12}_", "", n)       # prefixo de hora que o script junta quando o nome já existia
    n = re.sub(r"^\d{8}-", "", n)          # nº de sequência de importações antigas
    return n.lower()


def regra_para_nome(nome):
    base = nome_base_para_regras(nome)
    for rx, regra in carregar_regras_nome():
        if rx.search(base):
            return regra
    return None


# ---------------------------------------------------------------------------
# PDFs PROTEGIDOS (config/pdf_passwords.txt, uma palavra-passe por linha)
# ---------------------------------------------------------------------------
def carregar_passwords_pdf():
    f = CONFIG_DIR / "pdf_passwords.txt"
    if not f.exists():
        return []
    return [l.strip() for l in f.read_text(encoding="utf-8").splitlines() if l.strip() and not l.startswith("#")]


def abrir_pdf_reader(caminho):
    """(reader|None, senha|None, motivo|None). Experimenta as palavras-passe guardadas se o PDF estiver protegido."""
    try:
        reader = PdfReader(caminho)
    except Exception as e:
        return None, None, f"ilegível ({type(e).__name__})"
    if not reader.is_encrypted:
        return reader, None, None
    for pw in [""] + carregar_passwords_pdf():
        try:
            if reader.decrypt(pw):
                return reader, pw, None
        except Exception as e:           # p. ex. falta instalar 'cryptography' para PDFs AES
            registar_log(f"[PDF protegido] {Path(caminho).name}: {type(e).__name__}: {e}")
            return None, None, "protegido (instala 'pip install cryptography' ou confirma a palavra-passe)"
    return None, None, "protegido por palavra-passe (acrescenta-a no menu Gerir registos, opção 8)"


def texto_pdf_completo(caminho, paginas=30):
    with SilenciadorStderr():
        reader, _, _ = abrir_pdf_reader(caminho)
        if reader is None:
            return ""
        try:
            return "\n".join((p.extract_text() or "") for p in reader.pages[:paginas])
        except Exception:
            return ""


# ---------------------------------------------------------------------------
# DATA DO DOCUMENTO e ANO
# ---------------------------------------------------------------------------
MESES_PT = {"janeiro": 1, "fevereiro": 2, "março": 3, "marco": 3, "abril": 4, "maio": 5, "junho": 6, "julho": 7,
            "agosto": 8, "setembro": 9, "outubro": 10, "novembro": 11, "dezembro": 12}
_DATA_FLEX = (r"(\d{1,2}[-/.]\d{1,2}[-/.](?:19|20)\d{2}|(?:19|20)\d{2}[-/.]\d{1,2}[-/.]\d{1,2}"
              r"|\d{1,2}\s+de\s+[a-zçã]+\s+de\s+(?:19|20)\d{2})")
RE_DATA_ROTULADA = re.compile(
    r"(?:data\s*(?:da|de|do)?\s*(?:fatura|factura|emiss[ãa]o|documento|processamento|fatura[çc][ãa]o)"
    r"|emitid[ao]\s+em|invoice\s+date|emiss[ãa]o)\W{0,12}" + _DATA_FLEX, re.IGNORECASE)
RE_DATA_TRANSACAO = re.compile(
    r"(?:data\s*/\s*hora|data(?:\s+(?:de|do|da))?\s+(?:pagamento|movimento|valor|transa[cç][ãa]o))"
    r"\W{0,16}" + _DATA_FLEX,
    re.IGNORECASE,
)


def _valida_data(y, m=None, d=None):
    if not (1990 <= y <= datetime.now().year + 1):
        return None
    if m is not None and not 1 <= m <= 12:
        return None
    if d is not None:
        try:
            datetime(y, m or 1, d)
        except ValueError:
            return None
    return (y, m, d)


def _parse_data_flex(txt):
    """'2026-03-05', '05/03/2026', '5 de março de 2026', '2026-03', '2026' -> (ano, mês|None, dia|None)."""
    txt = str(txt or "").strip().lower()
    if not txt:
        return None
    m = re.fullmatch(r"(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})(?:[ t].*)?", txt)
    if m:
        return _valida_data(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    m = re.fullmatch(r"(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})", txt)
    if m:
        return _valida_data(int(m.group(3)), int(m.group(2)), int(m.group(1)))
    m = re.fullmatch(r"(\d{1,2})\s+de\s+([a-zçã]+)\s+de\s+(\d{4})", txt)
    if m and m.group(2) in MESES_PT:
        return _valida_data(int(m.group(3)), MESES_PT[m.group(2)], int(m.group(1)))
    m = re.fullmatch(r"(\d{4})[-/.](\d{1,2})", txt)
    if m:
        return _valida_data(int(m.group(1)), int(m.group(2)))
    m = re.fullmatch(r"(\d{4})", txt)
    if m:
        return _valida_data(int(m.group(1)))
    return None


def datas_no_nome(nome):
    """Datas escritas no nome do ficheiro (AAAAMMDD, AAAA-MM-DD, DD-MM-AAAA, '18 de Julho de 2026', AAAA_MM, MM_AAAA, AAAAMM, AAAA)."""
    nome = re.sub(r"^\d{6,12}_", "", Path(str(nome)).stem)          # o prefixo de hora não é uma data
    out = []
    for m in re.finditer(r"(?<!\d)(20\d{2})[-_. ]?(0[1-9]|1[0-2])[-_. ]?(0[1-9]|[12]\d|3[01])(?!\d)", nome):
        out.append(_valida_data(int(m.group(1)), int(m.group(2)), int(m.group(3))))
    for m in re.finditer(r"(?<!\d)(0[1-9]|[12]\d|3[01])[-_.](0[1-9]|1[0-2])[-_.](20\d{2})(?!\d)", nome):
        out.append(_valida_data(int(m.group(3)), int(m.group(2)), int(m.group(1))))
    for m in re.finditer(r"(?<!\d)(\d{1,2})\s+de\s+([A-Za-zçÇãÃ]+)\s+de\s+(20\d{2})(?!\d)", nome):
        out.append(_parse_data_flex(m.group(0)))
    if not out:
        for m in re.finditer(r"(?<!\d)(20\d{2})[-_.](0[1-9]|1[0-2])(?!\d)", nome):
            out.append(_valida_data(int(m.group(1)), int(m.group(2))))
        for m in re.finditer(r"(?<!\d)(0[1-9]|1[0-2])[-_.](20\d{2})(?!\d)", nome):
            out.append(_valida_data(int(m.group(2)), int(m.group(1))))
    if not out:
        for m in re.finditer(r"(?<!\d)(20\d{2})(0[1-9]|1[0-2])(?!\d)", nome):
            out.append(_valida_data(int(m.group(1)), int(m.group(2))))
    if not out:
        for m in re.finditer(r"(?<!\d)(20\d{2})(?!\d)", nome):
            out.append(_valida_data(int(m.group(1))))
    return [d for d in out if d]


def datas_com_rotulo(texto):
    """Só datas ligadas a 'data de emissão / data da fatura...' (evita datas limite, nascimentos, etc.)."""
    return [d for d in (_parse_data_flex(m.group(1)) for m in RE_DATA_ROTULADA.finditer(texto or "")) if d]


def datas_de_transacao(texto):
    """Extrai datas explicitamente associadas a pagamento/movimento, sem usar datas soltas."""
    return [d for d in (_parse_data_flex(m.group(1)) for m in RE_DATA_TRANSACAO.finditer(texto or "")) if d]


def texto_rapido_pdf(caminho, paginas=2):
    """Texto das primeiras páginas, sem OCR (rápido). Vazio se o PDF for só imagem ou estiver protegido."""
    return texto_pdf_completo(caminho, paginas)


def formatar_data(t):
    y, m, d = t
    return f"{y:04d}-{m:02d}-{d:02d}" if d else (f"{y:04d}-{m:02d}" if m else f"{y:04d}")


def avaliar_datas(qr_data=None, texto="", nome="", data_limite=None, data_email=None, data_doc_bd=None):
    """
    Devolve {'data_doc': str|None, 'ano': int|None, 'antigo': bool}.
      data_doc: a melhor data do próprio documento (QR, data de emissão, nome do ficheiro); na falta, a data do
                email; a data limite só serve (como aproximação) quando o documento é antigo. Para documentos
                atuais, uma data limite sozinha NÃO decide a pasta.
      antigo:   True só se TODAS as datas encontradas (incluindo data limite e data do email) são anteriores a
                ANO_VERIFICACAO. Qualquer data de 2026 ou posterior => verifica-se normalmente.
    """
    qr = _parse_data_flex(qr_data)
    doc = [d for d in (qr, _parse_data_flex(data_doc_bd)) if d]
    doc += datas_com_rotulo(texto)
    doc += datas_no_nome(nome)
    email, limite = _parse_data_flex(data_email), _parse_data_flex(data_limite)
    todas = doc + [d for d in (email, limite) if d]
    if not todas:
        return {"data_doc": None, "ano": None, "antigo": False}
    chave = lambda t: (t[0], t[1] or 0, t[2] or 0)
    antigo = max(t[0] for t in todas) < ANO_VERIFICACAO
    melhor = qr or (max(doc, key=chave) if doc else None) or email or (limite if antigo else None)
    return {"data_doc": formatar_data(melhor) if melhor else None,
            "ano": melhor[0] if melhor else max(t[0] for t in todas), "antigo": antigo}


def pasta_ano_mes(data_doc):
    """'2025-03-10' -> '2025/03'; só ano -> '2025/sem mes'; sem data -> mês atual."""
    t = _parse_data_flex(data_doc)
    if not t:
        return datetime.now().strftime("%Y/%m")
    return f"{t[0]}/{t[1]:02d}" if t[1] else f"{t[0]}/sem mes"


# NIF -> (emissor, categoria), tirado da lista acima
NIFS_CONHECIDOS = {k: (nome, cat) for chaves, nome, cat in EMISSORES_CONHECIDOS for k in chaves if re.fullmatch(r"\d{9}", k)}

RE_ENTIDADE = re.compile(r"Entidade[:\s]+(\d{5})(?!\d)", re.IGNORECASE)
RE_REF_MB = re.compile(r"Refer[êe]ncia[:\s]+(\d{3}\s?\d{3}\s?\d{3})(?!\d)", re.IGNORECASE)
RE_REF_ESTADO = re.compile(
    r"(?:Refer[êe]ncia(?:\s+de\s+pagamento)?|Ref\.?)[^\d\n]{0,25}(\d{3}\s?\d{3}\s?\d{3}\s?\d{3}\s?\d{3}|\d{12,15})(?!\d)",
    re.IGNORECASE)
PALAVRAS_ESTADO = ["autoridade tributária", "autoridade tributaria", "finanças", "financas", "portal das finanças",
                   "segurança social", "seg. social", "documento único de cobrança", "duc", "iuc", "irs", "irc", "imi"]
PALAVRAS_TRANSFERENCIA = ["transferência", "transferencia", "ordenante", "beneficiário", "beneficiario", "iban", "nib",
                          "conta origem", "conta de origem", "conta destino", "conta de destino"]


def contar_contas(texto):
    """Quantas contas bancárias (IBAN, NIB ou conta mascarada) aparecem no texto."""
    n = len(re.findall(r"PT\s?\d{2}[\s\d*xX•.]{17,32}", texto, re.IGNORECASE))
    n += len(re.findall(r"\b\d{4}\s?\d{4}\s?\d{11}\s?\d{2}\b", texto))
    n += len(re.findall(r"(?:conta|iban|nib)[^\n]{0,25}?[\d*xX•]{4,}", texto, re.IGNORECASE))
    return n


def classificar_comprovativo(texto):
    """
    'Pagamento'            entidade + referência (+ montante)
    'Pagamento ao Estado'  só referência (sem entidade) e com marcas do Estado (AT, Segurança Social...)
    'Transferencia'        duas contas bancárias, sem entidade nem referência
    None                   não dá para saber
    """
    t = texto.lower()
    ent = RE_ENTIDADE.search(texto)
    ref = RE_REF_MB.search(texto)
    ref_estado = RE_REF_ESTADO.search(texto)
    if not ent and (ref or ref_estado) and any(contem(t, p) for p in PALAVRAS_ESTADO):
        return "Pagamento ao Estado"
    if ent and ref:
        return "Pagamento"
    if not ent and not ref:
        contas = contar_contas(texto)
        if contas >= 2 or (contas >= 1 and any(k in t for k in PALAVRAS_TRANSFERENCIA)) \
                or "transferência" in t or "transferencia" in t:
            return "Transferencia"
    return None


def _campos_basicos(texto):
    """(entidade, referência só com dígitos, valor) para a votação de frames e para escolher o melhor OCR."""
    ent = RE_ENTIDADE.search(texto)
    ref = RE_REF_MB.search(texto) or RE_REF_ESTADO.search(texto)
    val = RE_VALOR.search(texto)
    valor = None
    if val:
        try:
            valor = parse_valor(val.group(1))
        except ValueError:
            valor = None
    return (ent.group(1) if ent else None, re.sub(r"\D", "", ref.group(1)) if ref else None, valor)


# ---------------------------------------------------------------------------
# OCR melhorado: pré-processamento + várias tentativas
# ---------------------------------------------------------------------------
_OCR_LANG = ["por+eng"]


def preprocessar_para_ocr(img, ampliar=False):
    """Cinzento, inverte fundos escuros (modo escuro) e contraste automático. Só amplia se pedido (texto pequeno)."""
    g = ImageOps.grayscale(img)
    if ImageStat.Stat(g).mean[0] < 110:
        g = ImageOps.invert(g)
    g = ImageOps.autocontrast(g, cutoff=2)
    if ampliar and g.width < 1600:
        g = g.resize((g.width * 2, g.height * 2), Image.LANCZOS).filter(ImageFilter.SHARPEN)
    return g


def _ocr(img, psm=None):
    cfg = f"--oem 3 --psm {psm}" if psm else ""
    try:
        return pytesseract.image_to_string(img, lang=_OCR_LANG[0], config=cfg)
    except Exception:
        if _OCR_LANG[0] != "eng":      # falta o pacote de português: usa só inglês
            _OCR_LANG[0] = "eng"
            return pytesseract.image_to_string(img, lang="eng", config=cfg)
        raise


def ocr_melhor(img):
    """
    OCR com pré-processamento leve. Se o resultado for fraco, tenta a imagem ampliada (texto pequeno) e outros
    modos de leitura, e fica com o resultado que encontra mais campos (entidade, referência, montante).
    """
    pontos = lambda t: sum(c is not None for c in _campos_basicos(t))
    leve = preprocessar_para_ocr(img)
    melhor = _ocr(leve)
    score = pontos(melhor)

    def melhorou(t):
        nonlocal melhor, score
        s_t = pontos(t)
        if s_t > score or (s_t == score and len(t.strip()) > len(melhor.strip())):
            melhor, score = t, s_t

    if len(melhor.strip()) < 25:                         # quase nada lido: tenta sem tratamento e ampliada
        melhorou(_ocr(img))
        melhorou(_ocr(preprocessar_para_ocr(img, ampliar=True)))
    parece_comprovativo = any(k in melhor.lower() for k in
                              ("montante", "comprovativo", "transfer", "pagamento", "refer", "entidade", "iban"))
    if score < 3 and parece_comprovativo:
        melhorou(_ocr(preprocessar_para_ocr(img, ampliar=True)))
        for psm in (6, 4, 11):
            if score >= 3:
                break
            melhorou(_ocr(leve, psm))
    return melhor


# ---------------------------------------------------------------------------
# QR code das faturas portuguesas (AT)
# ---------------------------------------------------------------------------
TIPO_DOC_QR = {"FR": "Fatura-Recibo", "FS": "Fatura-Recibo", "VD": "Fatura-Recibo", "NC": "Nota de Crédito"}


def parse_qr_at(texto):
    """Texto do QR (campos CHAVE:valor separados por *) -> dicionário, ou None se não for de uma fatura."""
    campos = {}
    for parte in (texto or "").split("*"):
        if ":" in parte:
            k, v = parte.split(":", 1)
            campos[k.strip()] = v.strip()
    if not (re.fullmatch(r"\d{9}", campos.get("A", "")) and "F" in campos and "O" in campos):
        return None
    try:
        total = round(float(campos["O"].replace(",", ".")), 2)
    except ValueError:
        return None
    data = None
    if re.fullmatch(r"\d{8}", campos["F"]):
        try:
            data = datetime.strptime(campos["F"], "%Y%m%d").date().isoformat()
        except ValueError:
            data = None
    return {"nif_emissor": campos["A"], "nif_adquirente": campos.get("B"), "tipo": campos.get("D"),
            "data": data, "num_doc": campos.get("G"), "atcud": campos.get("H"), "total": total}


def ler_qr_imagem(img):
    """Procura um QR de fatura (começa por A:<NIF>) numa imagem PIL. Precisa do OpenCV."""
    if cv2 is None:
        return None
    try:
        bgr = np.array(img.convert("RGB"))[:, :, ::-1].copy()
        det = cv2.QRCodeDetector()
        candidatos = []
        try:
            ok, infos, _, _ = det.detectAndDecodeMulti(bgr)
            if ok:
                candidatos.extend(infos)
        except Exception:
            pass
        if not candidatos:
            texto, _, _ = det.detectAndDecode(bgr)
            candidatos.append(texto)
        for c in candidatos:
            qr = parse_qr_at(c) if c else None
            if qr:
                return qr
    except Exception as e:
        registar_log(f"[QR] erro: {e}")
    return None


def ler_qr_pdf(caminho_pdf):
    """QR da AT na primeira e na última página do PDF (onde a especificação o permite)."""
    if cv2 is None:
        return None
    try:
        _, senha, _ = abrir_pdf_reader(caminho_pdf)
        with SilenciadorStderr():
            with pdfplumber.open(caminho_pdf, password=senha) as pdf:
                if not pdf.pages:
                    return None
                paginas = [pdf.pages[0]] + ([pdf.pages[-1]] if len(pdf.pages) > 1 else [])
                for resolucao in (200, 300):
                    for pagina in paginas:
                        qr = ler_qr_imagem(pagina.to_image(resolution=resolucao).original)
                        if qr:
                            return qr
    except Exception as e:
        registar_log(f"[QR] {Path(caminho_pdf).name}: {e}")
    return None


def analisar_texto_estruturado(texto, nome_ficheiro):
    texto_lower = texto.lower()
    nome_lower = nome_ficheiro.lower()

    tipo_doc = "Diversos"
    if (any(k in texto_lower for k in ["fatura", "invoice", "fatura/recibo", "nota de débito"])
            or any(k in nome_lower for k in ["fatura", "invoice"]) or contem(nome_lower, "ft")):
        tipo_doc = "Fatura-Recibo" if "recibo" in texto_lower else "Fatura"
    elif (any(k in texto_lower for k in ["comprovativo", "talão", "transferência", "pagamento efetuado", "recibo de pagamento", "referência mb", "montante"])
          or any(k in nome_lower for k in ["comprovativo", "talão", "pagamento"]) or contem(nome_lower, "vid")):
        tipo_doc = "Comprovativo"
    elif (any(k in texto_lower for k in ["aviso de cobrança", "quotas", "trimestre", "condomínio"])
          or any(k in nome_lower for k in ["trim", "trimestre", "condominio"])):
        tipo_doc = "Aviso de Cobrança / Condomínio"
    elif "nota de crédito" in texto_lower:
        tipo_doc = "Nota de Crédito"

    emissor, categoria = "Desconhecido", "Geral"
    texto_para_analisar = f"{nome_lower} {texto_lower}"
    for chaves, nome_oficial, cat in EMISSORES_CONHECIDOS:
        if any(contem(texto_para_analisar, chave) for chave in chaves):
            emissor, categoria = nome_oficial, cat
            break

    entidade = RE_ENTIDADE.search(texto)
    referencia = RE_REF_MB.search(texto)
    ref_txt = referencia.group(1) if referencia else ""
    valor = RE_VALOR.search(texto)

    subtipo = None
    if tipo_doc == "Comprovativo":
        subtipo = classificar_comprovativo(texto)
        if subtipo == "Transferencia":
            tipo_doc = "Transferencia"
        elif subtipo == "Pagamento ao Estado":
            emissor = "Estado"
            categoria = "Estado / Segurança Social" if "segurança social" in texto_lower else "Estado / Impostos"
            m_estado = RE_REF_ESTADO.search(texto)
            ref_txt = m_estado.group(1) if m_estado else ref_txt
            entidade = None
    data_limite = re.search(
        r"(?:Limite|Vencimento|Pagar at[ée]|Data limite|Data Vencimento)[^\d]*(\d{2}[-/]\d{2}[-/]\d{4}|\d{4}[-/]\d{2}[-/]\d{2})",
        texto, re.IGNORECASE)

    val_float = 0.0
    if valor:
        try:
            val_float = parse_valor(valor.group(1))
        except ValueError:
            val_float = 0.0

    regra = regra_para_nome(nome_ficheiro)           # o nome do ficheiro manda (ex.: Detalhe_516671702 = Via Verde)
    if regra:
        if regra.get("emissor"):
            emissor, categoria = regra["emissor"], regra.get("categoria") or categoria
        elif regra.get("categoria"):
            categoria = regra["categoria"]
        if regra.get("tipo"):
            tipo_doc = regra["tipo"]

    return {
        "Categoria": categoria,
        "Emissor": emissor,
        "Tipo Doc": tipo_doc,
        "Entidade MB": entidade.group(1) if entidade else "",
        "Referência MB": ref_txt,
        "Valor": val_float,
        "Data limite": data_limite.group(1) if data_limite else "",
        "Subtipo": subtipo,
        "Regra": regra
    }


def aplicar_qr(dados, qr, divergencias):
    """O QR da AT é a fonte mais fiável: manda no emissor (NIF), no valor, na data e no tipo de documento."""
    dados["NIF emissor"], dados["ATCUD"], dados["Num doc"], dados["Data doc"] = (
        qr["nif_emissor"], qr["atcud"], qr["num_doc"], qr["data"])
    conhecido = NIFS_CONHECIDOS.get(qr["nif_emissor"])
    if conhecido:
        dados["Emissor"], dados["Categoria"] = conhecido
    elif dados["Emissor"] in ("Desconhecido", EMISSOR_COMPROVATIVO):
        dados["Emissor"], dados["Categoria"] = f"NIF {qr['nif_emissor']}", "Geral"
    divergencias[:] = [d for d in divergencias
                       if not d.startswith("Emissor (") and d != "Emissor não identificado"]
    if qr["total"] > 0:
        dados["Valor"] = qr["total"]
        divergencias[:] = [d for d in divergencias if not d.startswith("Valor (") and d != "Valor a 0.00€"]
    if dados["Tipo Doc"] in ("Diversos", "Fatura", "Fatura-Recibo", "Nota de Crédito"):
        dados["Tipo Doc"] = TIPO_DOC_QR.get(qr["tipo"], "Fatura")
    return f"QR AT (NIF {qr['nif_emissor']}, ATCUD {qr['atcud'] or '-'})"


def dupla_verificacao_pdf(caminho_pdf, data_email=None):
    texto_a = ""
    texto_b = ""
    eh_imagem_ou_video = False

    reader, senha, protegido = abrir_pdf_reader(caminho_pdf)
    with SilenciadorStderr():
        try:
            if reader is not None:
                for page in reader.pages:
                    texto_a += page.extract_text() or ""
        except Exception as e:
            registar_log(f"[Motor A] Erro: {e}")

        try:
            with pdfplumber.open(caminho_pdf, password=senha) as pdf:
                for page in pdf.pages:
                    t = page.extract_text(layout=True) or ""
                    if not t.strip():
                        eh_imagem_ou_video = True
                        img_pagina = page.to_image(resolution=200).original
                        t = ocr_melhor(img_pagina)
                    texto_b += t + "\n"
        except Exception as e:
            registar_log(f"[Motor B] Erro: {e}")

    if eh_imagem_ou_video and not texto_a.strip():
        texto_a = texto_b

    res_a = analisar_texto_estruturado(texto_a, caminho_pdf.name)
    res_b = analisar_texto_estruturado(texto_b, caminho_pdf.name)

    divergencias = []
    if res_a["Emissor"] != res_b["Emissor"]:
        divergencias.append(f"Emissor ({res_a['Emissor']} vs {res_b['Emissor']})")
    if res_a["Valor"] != res_b["Valor"]:
        divergencias.append(f"Valor ({res_a['Valor']}€ vs {res_b['Valor']}€)")
    if limpar_ref(res_a["Referência MB"]) != limpar_ref(res_b["Referência MB"]):
        divergencias.append(f"Ref MB ({res_a['Referência MB']} vs {res_b['Referência MB']})")
    if res_a["Entidade MB"] != res_b["Entidade MB"]:
        divergencias.append(f"Entidade ({res_a['Entidade MB']} vs {res_b['Entidade MB']})")
    if res_a["Subtipo"] and res_b["Subtipo"] and res_a["Subtipo"] != res_b["Subtipo"]:
        divergencias.append(f"Tipo de comprovativo ({res_a['Subtipo']} vs {res_b['Subtipo']})")

    if res_a["Emissor"] == "Desconhecido" or res_b["Emissor"] == "Desconhecido":
        divergencias.append("Emissor não identificado")
    if res_a["Tipo Doc"] != "Diversos" and (res_a["Valor"] == 0.0 or res_b["Valor"] == 0.0):
        if "Condomínio" not in res_a["Emissor"] and "Seguros" not in res_a["Categoria"]:
            divergencias.append("Valor a 0.00€")
    if protegido:
        divergencias.append(f"PDF {protegido}")

    # Regra por nome de ficheiro: o tipo de documento já se sabe, não é preciso emissor nem valor
    regra = res_a["Regra"] or res_b["Regra"]
    if regra:
        divergencias[:] = [d for d in divergencias if d != "Emissor não identificado"
                           and not (regra.get("emissor") and d.startswith("Emissor ("))]
        if regra.get("estado") == "Informativo" or regra.get("tipo") in ("Extrato", "Diversos"):
            divergencias[:] = [d for d in divergencias if d != "Valor a 0.00€"]
        if regra.get("emissor") and protegido:
            divergencias[:] = [d for d in divergencias if not d.startswith("PDF protegido")]

    dados = {
        "Categoria": res_b["Categoria"] if res_b["Categoria"] != "Geral" else res_a["Categoria"],
        "Emissor": res_b["Emissor"] if res_b["Emissor"] != "Desconhecido" else res_a["Emissor"],
        "Tipo Doc": res_a["Tipo Doc"],
        "Entidade MB": res_b["Entidade MB"] or res_a["Entidade MB"],
        "Referência MB": res_b["Referência MB"] or res_a["Referência MB"],
        "Valor": res_b["Valor"] if res_b["Valor"] > 0 else res_a["Valor"],
        "Data limite": res_b["Data limite"] or res_a["Data limite"],
        "Estado": "Por pagar",
        "Origem": "Gmail",
        "Ficheiro": caminho_pdf.name,
        "Notas": ""
    }
    subtipo = res_b["Subtipo"] or res_a["Subtipo"]

    # Num comprovativo (ex.: print do MB WAY) o emissor costuma ser o banco, não quem cobra.
    # Se tem referência ou valor, aceita-se e o cruzamento com as faturas trata de o associar.
    if (res_a["Tipo Doc"] in ("Comprovativo", "Transferencia") and "Emissor não identificado" in divergencias
            and res_a["Emissor"] == res_b["Emissor"] and (dados["Referência MB"] or dados["Valor"] > 0)):
        divergencias.remove("Emissor não identificado")
        dados["Emissor"] = EMISSOR_COMPROVATIVO
        dados["Categoria"] = "Transferências" if dados["Tipo Doc"] == "Transferencia" else "Pagamentos"

    # Regras dos comprovativos: um pagamento tem sempre entidade + referência + montante;
    # só o pagamento ao Estado vem apenas com referência; a transferência não tem entidade nem referência.
    if dados["Tipo Doc"] == "Comprovativo":
        tem_ent, tem_ref = bool(dados["Entidade MB"]), bool(dados["Referência MB"])
        if subtipo is None and tem_ent != tem_ref:
            divergencias.append("Comprovativo incompleto (só tem " + ("entidade" if tem_ent else "referência") + ")")
        elif subtipo == "Pagamento" and not (tem_ent and tem_ref):
            divergencias.append("Pagamento sem entidade e referência completas")
        elif subtipo == "Pagamento ao Estado" and not tem_ref:
            divergencias.append("Pagamento ao Estado sem referência")

    notas_extra = []
    if subtipo in ("Transferencia", "Pagamento ao Estado"):
        notas_extra.append(subtipo)

    # QR da AT (só vale a pena em faturas ou quando o emissor não foi identificado)
    qr = None
    if dados["Tipo Doc"] in ("Fatura", "Fatura-Recibo", "Aviso de Cobrança / Condomínio", "Nota de Crédito", "Diversos") \
            or "Emissor não identificado" in divergencias:
        qr = ler_qr_pdf(caminho_pdf)
        if qr:
            notas_extra.append(aplicar_qr(dados, qr, divergencias))

    if dados["Tipo Doc"] == "Diversos":
        dados["Estado"] = "Informativo"
    elif dados["Tipo Doc"] in ("Comprovativo", "Transferencia", "Fatura-Recibo"):
        dados["Estado"] = "Pago"

    if regra and regra.get("estado"):
        dados["Estado"] = regra["estado"]
    if regra and protegido and regra.get("emissor"):
        notas_extra.append("PDF protegido: conteúdo não lido (acrescenta a palavra-passe e reprocessa)")

    # Data do documento e ano: anos anteriores a ANO_VERIFICACAO ficam "Ok" e não são verificados
    info = avaliar_datas(qr_data=(qr or {}).get("data"), texto=f"{texto_a}\n{texto_b}", nome=caminho_pdf.name,
                         data_limite=dados["Data limite"], data_email=data_email)
    if dados["Tipo Doc"] in ("Comprovativo", "Transferencia"):
        datas_pagamento = datas_de_transacao(f"{texto_a}\n{texto_b}")
        if datas_pagamento:
            data_pagamento = datas_pagamento[0]
            info["data_doc"] = formatar_data(data_pagamento)
            info["ano"] = data_pagamento[0]
            info["antigo"] = data_pagamento[0] < ANO_VERIFICACAO
    dados["Data doc"] = info["data_doc"]
    dados["_antigo"] = info["antigo"]
    if info["antigo"]:
        if dados["Estado"] in ESTADOS_PARA_OK:
            dados["Estado"] = "Ok"
        dados["Notas"] = f"Anterior a {ANO_VERIFICACAO}: sem verificação" + \
                         (f" (ignorado: {' | '.join(divergencias)})" if divergencias else "")
        return dados, True, ""

    if divergencias:
        motivo = " | ".join(divergencias)
        dados["Estado"] = "Por Identificar"
        dados["Notas"] = f"DIVERGÊNCIA DUPLA VERIFICAÇÃO: {motivo}"
        return dados, False, motivo

    metodo = "OCR + Motor B" if eh_imagem_ou_video else "Motor A+B OK"
    dados["Notas"] = " | ".join([f"Validado ({metodo}) em {datetime.now().strftime('%Y-%m-%d %H:%M')}"] + notas_extra)
    return dados, True, ""


def classe_doc(tipo):
    t = (tipo or "").strip().lower()
    if t in ("fatura", "aviso de cobrança / condomínio", "nota de crédito"):
        return "A_PAGAR"
    if t in ("comprovativo", "transferencia", "transferência", "pagamento ordenado fr"):
        return "PAGAMENTO"
    if t == "fatura-recibo":
        return "FATURA_RECIBO"
    return "OUTRO"


def _data_para_comparar(data_doc, data_limite):
    return str(data_doc or data_limite or "").strip()


def e_duplicado(dados):
    """
    Devolve (nivel, id, via):
      'certo'    mesmo QR (NIF + ATCUD); ou mesma classe + referência + entidade + valor com a MESMA data;
                 ou mesmo emissor + valor + data limite (faturas)
      'possivel' mesma classe + referência + entidade + valor mas sem datas para confirmar (pode ser um
                 pagamento recorrente com referência fixa: não se descarta, fica apenas assinalado)
      (None, None, None) caso contrário
    """
    conn = ligar_bd()
    try:
        if dados.get("NIF emissor") and dados.get("ATCUD"):
            r = conn.execute("SELECT id FROM faturas WHERE nif_emissor = ? AND atcud = ? LIMIT 1",
                             (dados["NIF emissor"], dados["ATCUD"])).fetchone()
            if r:
                return "certo", r[0], "qr"

        ref = limpar_ref(dados["Referência MB"])
        valor = dados["Valor"]
        nova_data = _data_para_comparar(dados.get("Data doc"), dados.get("Data limite"))
        classe = classe_doc(dados["Tipo Doc"])
        possivel = None
        if len(ref) >= 6 and valor > 0 and classe != "OUTRO":
            rows = conn.execute(
                "SELECT id, emissor, entidade_mb, referencia_mb, tipo_doc, data_limite, data_doc FROM faturas "
                "WHERE ABS(valor - ?) < 0.005", (valor,)).fetchall()
            for rid, emissor, ent, ref_ext, tipo_ext, dl, dd in rows:
                if limpar_ref(ref_ext) != ref or classe_doc(tipo_ext) != classe:
                    continue
                if ent and dados["Entidade MB"] and str(ent) != str(dados["Entidade MB"]):
                    continue
                if (emissor not in EMISSORES_GENERICOS and dados["Emissor"] not in EMISSORES_GENERICOS
                        and emissor != dados["Emissor"]):
                    continue
                data_ext = _data_para_comparar(dd, dl)
                if nova_data and data_ext:
                    if nova_data == data_ext:
                        return "certo", rid, "ref"
                    continue                       # datas diferentes = períodos diferentes
                possivel = possivel or rid
        if dados["Data limite"] and classe in ("A_PAGAR", "FATURA_RECIBO") and valor > 0:
            r = conn.execute(
                "SELECT id FROM faturas WHERE emissor = ? AND ABS(valor - ?) < 0.005 AND data_limite = ? "
                "AND tipo_doc NOT IN ('Comprovativo', 'Transferencia') LIMIT 1",
                (dados["Emissor"], valor, dados["Data limite"])).fetchone()
            if r:
                return "certo", r[0], "data"
        if possivel:
            return "possivel", possivel, "ref"
    finally:
        conn.close()
    return None, None, None


def mover_sem_sobrescrever(origem, pasta_destino, nome=None):
    pasta_destino.mkdir(parents=True, exist_ok=True)
    nome = nome or origem.name
    destino = pasta_destino / nome
    if destino.exists():
        destino = pasta_destino / f"{datetime.now().strftime('%H%M%S%f')}_{nome}"
    shutil.move(str(origem), str(destino))
    return destino


def processar_pdf(pdf):
    """Devolve 'ok', 'duplicado' ou 'quarentena'."""
    h = sha256_ficheiro(pdf)
    if hash_ja_registado(h):                      # o mesmo ficheiro já foi processado (mesmo que em quarentena)
        registar_log(f"Duplicado (mesmo conteúdo): {pdf.name}")
        mover_sem_sobrescrever(pdf, ARQUIVO_DIR / "duplicados")
        return "duplicado"

    dados, validado, _motivo = dupla_verificacao_pdf(pdf, data_email=data_email_do_pdf(pdf.name))
    antigo = dados.pop("_antigo", False)
    dados["Hash"] = h
    if WHATSAPP_DIR in pdf.parents or pdf.parent == WHATSAPP_DIR:
        dados["Origem"] = "WhatsApp"

    if dados["Tipo Doc"] != "Diversos":
        nivel, dup_id, via = e_duplicado(dados)
        if nivel == "certo" and (via == "qr" or (validado and not antigo)):
            registar_log(f"Duplicado descartado: {pdf.name} ({dados['Emissor']} - {dados['Valor']}€) = #{dup_id} [{via}]")
            mover_sem_sobrescrever(pdf, ARQUIVO_DIR / "duplicados")
            return "duplicado"
        if nivel == "possivel" and validado and not antigo:
            dados["Notas"] = f"{dados['Notas']} | Possível repetido de #{dup_id} (sem data para confirmar)"

    if not validado:
        novo = mover_sem_sobrescrever(pdf, POR_IDENTIFICAR_DIR)
        dados["Ficheiro"] = f"por identificar/{novo.name}"
        registar_bd(dados)
        return "quarentena"

    ano_mes = pasta_ano_mes(dados.get("Data doc"))
    fornecedor = re.sub(r'[\\/*?:"<>|]', "", dados["Emissor"])
    if dados["Tipo Doc"] == "Diversos":
        destino_dir = ARQUIVO_DIR / "diversos" / ano_mes
        novo = mover_sem_sobrescrever(pdf, destino_dir)
        dados["Ficheiro"] = f"diversos/{ano_mes}/{novo.name}"
    else:
        destino_dir = ARQUIVO_DIR / ano_mes / fornecedor
        novo = mover_sem_sobrescrever(pdf, destino_dir)
        dados["Ficheiro"] = f"{ano_mes}/{fornecedor}/{novo.name}"
    registar_log(f"[{'OK-ANTIGO' if antigo else 'VALIDADO'}] {pdf.name} | {dados['Emissor']} | {dados['Valor']}€")
    registar_bd(dados)
    return "ok"


def tarefa_auditoria_background(pdfs, estado_progresso):
    duplicados = quarentena = 0
    try:
        for i, pdf in enumerate(pdfs, 1):
            nome_original = pdf.name
            try:
                resultado = processar_pdf(pdf)
                if resultado == "duplicado":
                    duplicados += 1
                elif resultado == "quarentena":
                    quarentena += 1
                registar_resultado_pdf(nome_original, resultado)
            except Exception as e:
                registar_log(f"Erro a processar {nome_original}: {e}")
                registar_resultado_pdf(nome_original, "erro")
            estado_progresso["atual"] = i
            progresso("auditoria", i / max(1, len(pdfs)))
    finally:
        estado_progresso["duplicados"] = duplicados
        estado_progresso["quarentena"] = quarentena
        estado_progresso["concluido"] = True  # garante que a animação termina mesmo com erro


# ---------------------------------------------------------------------------
# EXCEL (vista derivada da base de dados, uma folha por mês, mês mais recente primeiro)
# ---------------------------------------------------------------------------
def caminho_do_registo(rel):
    if rel.startswith("manual/"):
        return None
    if rel.startswith("por identificar/"):
        return BASE_DIR / rel
    return ARQUIVO_DIR / rel


def sincronizar_bd_para_excel():
    inicializar_base_dados()
    conn = ligar_bd()
    try:
        registos = conn.execute(
            "SELECT categoria, emissor, tipo_doc, entidade_mb, referencia_mb, valor, data_limite, estado, origem, ficheiro, notas, "
            "data_doc, substr(data_registo, 1, 10) FROM faturas ORDER BY COALESCE(NULLIF(data_doc, ''), data_registo) DESC, id DESC").fetchall()
    finally:
        conn.close()

    def chave_folha(data_doc, estado, registo):
        t = _parse_data_flex(data_doc) or _parse_data_flex(registo) or (datetime.now().year, datetime.now().month, None)
        if t[0] < ANO_VERIFICACAO and estado not in ESTADOS_PENDENTES:
            return f"{t[0]}"                                   # anos anteriores: uma folha por ano
        return f"{t[0]}-{t[1]:02d}" if t[1] else f"{t[0]}-sem mes"

    por_mes = {}
    for reg in registos:
        por_mes.setdefault(chave_folha(reg[11], reg[7], reg[12]), []).append(reg)
    if not por_mes:
        por_mes[datetime.now().strftime("%Y-%m")] = []

    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    header_fill = PatternFill(start_color="E0E0E0", end_color="E0E0E0", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="000000")

    for mes in sorted(por_mes.keys(), reverse=True):
        ws = wb.create_sheet(title=mes)
        ws.append(COLUNAS)
        for col in range(1, len(COLUNAS) + 1):
            c = ws.cell(row=1, column=col)
            c.fill, c.font = header_fill, header_font
            c.alignment = Alignment(horizontal="center", vertical="center")

        for reg in por_mes[mes]:
            ws.append(list(reg[:12]))
            linha = ws.max_row
            estado, tipo = reg[7], reg[2]

            fill = font = None
            if estado == "Por Identificar":
                fill, font = FILL_VERMELHO, FONT_VERMELHO
            elif tipo == "Diversos":
                fill, font = FILL_AZUL, FONT_AZUL
            elif estado == "Por pagar":
                fill, font = FILL_AMARELO, FONT_AMARELO
            elif estado == "Pago (a confirmar)":
                fill, font = FILL_LARANJA, FONT_LARANJA
            elif estado and estado.lower() in ("pago", "fatura-recibo", "ordenado"):
                fill, font = FILL_VERDE, FONT_VERDE
            elif estado == "Ok":
                fill, font = FILL_CINZA, FONT_CINZA
            if fill:
                for col in range(1, len(COLUNAS) + 1):
                    ws.cell(row=linha, column=col).fill = fill
                    ws.cell(row=linha, column=col).font = font

            destino = caminho_do_registo(reg[9] or "")
            if destino is not None:
                cel = ws.cell(row=linha, column=10)
                cel.hyperlink = destino.resolve().as_uri()
                cel.font = Font(name="Calibri", size=11, color="0563C1", underline="single")

            vcel = ws.cell(row=linha, column=6)
            vcel.number_format = '#,##0.00 "€"'
            vcel.alignment = Alignment(horizontal="right", vertical="center")

        for col in range(1, len(COLUNAS) + 1):
            letra = openpyxl.utils.get_column_letter(col)
            maior = max((len(str(ws.cell(row=r, column=col).value or "")) for r in range(1, ws.max_row + 1)), default=0)
            ws.column_dimensions[letra].width = min(max(maior + 4, 14), 70)

    # Cópia de segurança do Excel anterior (o Excel é reconstruído a partir da BD)
    if EXCEL_FILE.exists():
        bk_dir = LOGS_DIR / "backup_excel"
        bk_dir.mkdir(parents=True, exist_ok=True)
        shutil.copy2(EXCEL_FILE, bk_dir / f"{EXCEL_FILE.stem}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
        for antigo in sorted(bk_dir.glob("*.xlsx"))[:-10]:
            antigo.unlink(missing_ok=True)

    try:
        wb.save(EXCEL_FILE)
    except PermissionError:
        alt = EXCEL_FILE.with_name(f"{EXCEL_FILE.stem}_{datetime.now().strftime('%H%M%S')}.xlsx")
        wb.save(alt)
        aviso(f"{Y}      -> O Excel estava aberto. Guardado como {alt.name}.{RESET}")


# ---------------------------------------------------------------------------
# GESTÃO DE REGISTOS  (adicionar = livre; alterar/apagar = 2FA)
# ---------------------------------------------------------------------------
CAMPOS_EDITAVEIS = {
    "1": ("Estado", "estado"),
    "2": ("Categoria", "categoria"),
    "3": ("Emissor", "emissor"),
    "4": ("Valor", "valor"),
    "5": ("Data limite", "data_limite"),
    "6": ("Notas", "notas"),
}


def listar_registos(filtro=None, limite=30):
    conn = ligar_bd()
    try:
        if filtro:
            like = f"%{filtro}%"
            rows = conn.execute(
                "SELECT id, substr(data_registo,1,10), emissor, tipo_doc, valor, estado, ficheiro FROM faturas "
                "WHERE emissor LIKE ? OR ficheiro LIKE ? OR categoria LIKE ? ORDER BY id DESC LIMIT ?",
                (like, like, like, limite)).fetchall()
        else:
            rows = conn.execute(
                "SELECT id, substr(data_registo,1,10), emissor, tipo_doc, valor, estado, ficheiro FROM faturas "
                "ORDER BY id DESC LIMIT ?", (limite,)).fetchall()
    finally:
        conn.close()
    if not rows:
        print(f"{Y}Sem registos.{RESET}")
        return
    print(f"\n{BOLD}{'ID':>4}  {'Data':<10}  {'Emissor':<28} {'Tipo':<14} {'Valor':>10}  {'Estado':<15} Ficheiro{RESET}")
    for r in rows:
        print(f"{r[0]:>4}  {r[1] or '':<10}  {str(r[2])[:28]:<28} {str(r[3])[:14]:<14} {(r[4] or 0):>9.2f}€  {str(r[5]):<15} {r[6]}")


def _pedir_id():
    try:
        return int(input("ID do registo: ").strip())
    except ValueError:
        print(f"{R}ID inválido.{RESET}")
        return None


def _obter_registo(rid):
    conn = ligar_bd()
    try:
        return conn.execute("SELECT id, emissor, tipo_doc, valor, estado, ficheiro FROM faturas WHERE id = ?", (rid,)).fetchone()
    finally:
        conn.close()


def adicionar_registo():
    print(f"\n{C}Novo registo manual (sem 2FA){RESET}")
    try:
        valor = float(input("Valor (€): ").strip().replace(",", ".") or 0)
    except ValueError:
        print(f"{R}Valor inválido.{RESET}")
        return
    estado = input(f"Estado {ESTADOS_VALIDOS} [Por pagar]: ").strip() or "Por pagar"
    if estado not in ESTADOS_VALIDOS:
        print(f"{R}Estado inválido.{RESET}")
        return
    dados = {
        "Categoria": input("Categoria: ").strip() or "Geral",
        "Emissor": input("Emissor: ").strip() or "Desconhecido",
        "Tipo Doc": input("Tipo de documento [Fatura]: ").strip() or "Fatura",
        "Entidade MB": input("Entidade MB (opcional): ").strip(),
        "Referência MB": input("Referência MB (opcional): ").strip(),
        "Valor": round(valor, 2),
        "Data limite": input("Data limite (opcional): ").strip(),
        "Estado": estado,
        "Origem": "Manual",
        "Ficheiro": f"manual/{datetime.now().strftime('%Y%m%d_%H%M%S')}",
        "Notas": input("Notas (opcional): ").strip(),
    }
    registar_bd(dados)
    registar_log(f"Registo manual adicionado: {dados['Emissor']} {dados['Valor']}€")
    sincronizar_bd_para_excel()
    print(f"{G}✓ Registo adicionado.{RESET}")


def alterar_registo():
    rid = _pedir_id()
    if rid is None:
        return
    reg = _obter_registo(rid)
    if not reg:
        print(f"{R}Registo não encontrado.{RESET}")
        return
    print(f"Registo {reg[0]}: {reg[1]} | {reg[2]} | {reg[3]}€ | {reg[4]} | {reg[5]}")
    for k, (nome, _col) in CAMPOS_EDITAVEIS.items():
        print(f"  {k}) {nome}")
    escolha = input("Campo a alterar: ").strip()
    if escolha not in CAMPOS_EDITAVEIS:
        print(f"{R}Opção inválida.{RESET}")
        return
    nome_campo, coluna = CAMPOS_EDITAVEIS[escolha]
    novo = input(f"Novo valor para {nome_campo}: ").strip()
    if coluna == "valor":
        try:
            novo = round(float(novo.replace(",", ".")), 2)
        except ValueError:
            print(f"{R}Valor inválido.{RESET}")
            return
    elif coluna == "estado" and novo not in ESTADOS_VALIDOS:
        print(f"{R}Estado inválido. Usa um de {ESTADOS_VALIDOS}.{RESET}")
        return

    if not verificar_dupla_autenticacao(f"Alterar {nome_campo} do registo {rid} para '{novo}'"):
        return
    conn = ligar_bd()
    try:
        conn.execute(f"UPDATE faturas SET {coluna} = ? WHERE id = ?", (novo, rid))  # coluna vem da lista fixa acima
        conn.commit()
    finally:
        conn.close()
    registar_log(f"Registo {rid}: {nome_campo} alterado para '{novo}'")
    sincronizar_bd_para_excel()
    print(f"{G}✓ Registo alterado.{RESET}")


def apagar_registo():
    rid = _pedir_id()
    if rid is None:
        return
    reg = _obter_registo(rid)
    if not reg:
        print(f"{R}Registo não encontrado.{RESET}")
        return
    print(f"Registo {reg[0]}: {reg[1]} | {reg[2]} | {reg[3]}€ | {reg[4]} | {reg[5]}")
    if input("Confirmas que queres apagar? (escreve APAGAR): ").strip() != "APAGAR":
        print("Cancelado.")
        return
    if not verificar_dupla_autenticacao(f"Apagar o registo {rid} ({reg[1]}, {reg[3]}€)"):
        return

    origem = caminho_do_registo(reg[5] or "")
    if origem is not None and origem.exists():
        mover_sem_sobrescrever(origem, ARQUIVO_DIR / "removidos")  # o ficheiro vai para 'removidos', não é destruído
    conn = ligar_bd()
    try:
        ligacao = conn.execute("SELECT tipo_doc, ligado_a FROM faturas WHERE id = ?", (rid,)).fetchone()
        if ligacao and ligacao[0] == "Comprovativo" and ligacao[1]:
            # apagar um comprovativo volta a pôr a fatura associada como 'Por pagar'
            conn.execute("UPDATE faturas SET estado = 'Por pagar', ligado_a = NULL WHERE id = ?", (ligacao[1],))
        conn.execute("UPDATE faturas SET ligado_a = NULL WHERE ligado_a = ?", (rid,))
        conn.execute("DELETE FROM faturas WHERE id = ?", (rid,))
        conn.commit()
    finally:
        conn.close()
    registar_log(f"Registo {rid} apagado ({reg[1]}, {reg[3]}€, {reg[5]})")
    sincronizar_bd_para_excel()
    print(f"{G}✓ Registo apagado (o ficheiro, se existia, foi para arquivo/removidos).{RESET}")


ESTADOS_NORMALIZADOS = {
    "pago": "Pago", "paga": "Pago", "pago (a confirmar)": "Pago (a confirmar)", "por pagar": "Por pagar",
    "informativo": "Informativo", "diversos": "Informativo", "por identificar": "Por Identificar",
    "ordenado": "Ordenado", "transferencia": "Pago", "transferência": "Pago", "ok": "Ok",
}


def _texto_ou_none(v):
    if v is None:
        return None
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    if hasattr(v, "isoformat"):
        v = v.isoformat()[:10]
    v = str(v).strip()
    return v or None


def importar_correcoes_excel():
    """Aplica à base de dados as correções feitas à mão no Excel (por ficheiro). Pede 2FA."""
    inicializar_base_dados()
    entrada = input(f"Caminho do Excel corrigido [{EXCEL_FILE}]: ").strip().strip('"') or str(EXCEL_FILE)
    caminho = Path(entrada)
    if not caminho.exists():
        print(f"{R}Ficheiro não encontrado: {caminho}{RESET}")
        return
    cab_para_col = {"Categoria": "categoria", "Emissor": "emissor", "Tipo Doc": "tipo_doc", "Entidade MB": "entidade_mb",
                    "Referência MB": "referencia_mb", "Valor": "valor", "Data limite": "data_limite", "Estado": "estado",
                    "Notas": "notas", "Data doc": "data_doc"}
    novos = {}
    wb = openpyxl.load_workbook(caminho, read_only=True, data_only=True)
    for ws in wb.worksheets:
        linhas = ws.iter_rows(values_only=True)
        try:
            cab = [str(c).strip() if c is not None else "" for c in next(linhas)]
        except StopIteration:
            continue
        if "Ficheiro" not in cab:
            continue
        idx = {col: cab.index(h) for h, col in cab_para_col.items() if h in cab}
        i_f = cab.index("Ficheiro")
        for lin in linhas:
            rel = lin[i_f] if i_f < len(lin) else None
            if rel:
                novos[str(rel).replace("\\", "/").strip()] = {c: (lin[i] if i < len(lin) else None) for c, i in idx.items()}
    wb.close()

    conn = ligar_bd()
    try:
        atuais = {r[1]: r for r in conn.execute(
            "SELECT id, ficheiro, categoria, emissor, tipo_doc, entidade_mb, referencia_mb, valor, data_limite, estado, notas, data_doc FROM faturas")}
    finally:
        conn.close()
    colunas = ["categoria", "emissor", "tipo_doc", "entidade_mb", "referencia_mb", "valor", "data_limite", "estado", "notas", "data_doc"]

    alteracoes, ignorados, estados_invalidos = [], 0, []
    por_campo = {}
    for rel, novo in novos.items():
        atual = atuais.get(rel)
        if atual is None:
            ignorados += 1
            continue
        antigo = dict(zip(["id", "ficheiro"] + colunas, atual))
        mudar = {}
        for col in colunas:
            if col not in novo:
                continue
            v = novo[col]
            if col == "valor":
                try:
                    v = round(float(str(v).replace(",", ".")), 2) if v not in (None, "") else None
                except ValueError:
                    continue
                if v is None or abs(v - (antigo[col] or 0)) < 0.005:
                    continue
            elif col in ("entidade_mb", "referencia_mb"):
                v = _texto_ou_none(v)
                if limpar_ref(v) == limpar_ref(antigo[col]):
                    continue
                v = re.sub(r"\D", "", v) if v else v
            elif col == "estado":
                v = ESTADOS_NORMALIZADOS.get((_texto_ou_none(v) or "").lower())
                if v is None:
                    estados_invalidos.append((rel, novo[col]))
                    continue
                if v == antigo[col]:
                    continue
            else:
                v = _texto_ou_none(v)
                if v == _texto_ou_none(antigo[col]):
                    continue
            mudar[col] = v
            por_campo[col] = por_campo.get(col, 0) + 1
        if mudar:
            alteracoes.append((antigo["id"], rel, mudar))

    print(f"\n{C}{BOLD}=== CORREÇÕES ENCONTRADAS NO EXCEL ==={RESET}")
    print(f"  Registos com alterações: {W}{len(alteracoes)}{RESET} (linhas do Excel sem registo na base de dados: {ignorados})")
    if por_campo:
        print("  Por campo: " + " · ".join(f"{k}: {n}" for k, n in sorted(por_campo.items(), key=lambda x: -x[1])))
    for rid, rel, mudar in alteracoes[:8]:
        print(f"   #{rid} {rel[-45:]}: " + ", ".join(f"{k}={v!r}" for k, v in mudar.items()))
    if estados_invalidos:
        print(f"  {Y}Estados não reconhecidos (ignorados): {len(estados_invalidos)}, p. ex. {estados_invalidos[0][1]!r}{RESET}")
    if not alteracoes:
        return
    if input("Escreve IMPORTAR para aplicar à base de dados: ").strip() != "IMPORTAR":
        print("Cancelado.")
        return
    if not verificar_dupla_autenticacao(f"Importar {len(alteracoes)} correção(ões) do Excel para a base de dados"):
        return

    (LOGS_DIR / "backup_db").mkdir(parents=True, exist_ok=True)
    bk = LOGS_DIR / "backup_db" / f"Faturas_Shadow_77_antes_importar_{datetime.now().strftime('%Y%m%d_%H%M%S')}.db"
    origem = sqlite3.connect(DB_FILE)
    destino = sqlite3.connect(bk)
    origem.backup(destino)
    destino.close()
    origem.close()

    conn = ligar_bd()
    try:
        for rid, rel, mudar in alteracoes:
            sets = ", ".join(f"{c} = ?" for c in mudar)      # nomes de colunas vêm da lista fixa acima
            conn.execute(f"UPDATE faturas SET {sets} WHERE id = ?", list(mudar.values()) + [rid])
        conn.commit()
    finally:
        conn.close()
    registar_log(f"Importadas {len(alteracoes)} correções do Excel ({caminho.name}); cópia da BD em {bk.name}")
    sincronizar_bd_para_excel()
    print(f"{G}✓ {len(alteracoes)} registo(s) atualizado(s). Cópia da base de dados: {bk}{RESET}")


def adicionar_password_pdf():
    pw = getpass.getpass("Palavra-passe do PDF (não aparece no ecrã): ").strip()
    if not pw:
        return
    if pw in carregar_passwords_pdf():
        print(f"{Y}Essa palavra-passe já está guardada.{RESET}")
        return
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    with open(CONFIG_DIR / "pdf_passwords.txt", "a", encoding="utf-8") as f:
        f.write(pw + "\n")
    registar_log("Palavra-passe de PDF acrescentada a config/pdf_passwords.txt")
    print(f"{G}✓ Guardada em {CONFIG_DIR / 'pdf_passwords.txt'} (ficheiro de texto: não o partilhes).{RESET}")


def aplicar_regras_e_reprocessar():
    """Aplica as regras por nome aos registos existentes e relê (com palavras-passe) os que estão em quarentena."""
    inicializar_base_dados()
    carregar_regras_nome(forcar=True)
    cols = ["id", "ficheiro", "emissor", "categoria", "tipo_doc", "estado", "data_doc", "data_limite", "notas"]
    conn = ligar_bd()
    try:
        rows = [dict(zip(cols, r)) for r in conn.execute(f"SELECT {', '.join(cols)} FROM faturas")]
    finally:
        conn.close()

    alteracoes = []                                     # (id, ficheiro, {coluna: valor}, origem)
    for d in rows:
        nome = Path(d["ficheiro"] or "").name
        regra = regra_para_nome(nome)
        if not regra:
            continue
        mudar = {}
        if regra.get("emissor") and d["emissor"] != regra["emissor"]:
            mudar["emissor"] = regra["emissor"]
        if (regra.get("categoria") and d["categoria"] != regra["categoria"]):
            mudar["categoria"] = regra["categoria"]
        if regra.get("tipo") and d["tipo_doc"] != regra["tipo"]:
            mudar["tipo_doc"] = regra["tipo"]
        info = avaliar_datas(nome=nome, data_limite=d["data_limite"], data_doc_bd=d["data_doc"])
        alvo = regra.get("estado")
        if not alvo and regra.get("tipo") == "Fatura" and d["estado"] in ("Informativo", "Por Identificar"):
            alvo = "Por pagar"
        if alvo and d["estado"] in ("Informativo", "Por Identificar", alvo):
            if info["antigo"] and alvo in ESTADOS_PARA_OK:
                alvo = "Ok"
            if alvo != d["estado"]:
                mudar["estado"] = alvo
        if info["data_doc"] and info["data_doc"] != (d["data_doc"] or ""):
            mudar["data_doc"] = info["data_doc"]
        if mudar:
            alteracoes.append((d["id"], d["ficheiro"], mudar, "regra"))

    # quarentena sem regra: reler (OCR e palavras-passe)
    quarentena = []
    for d in rows:
        if d["estado"] != "Por Identificar" or regra_para_nome(Path(d["ficheiro"] or "").name):
            continue
        p = caminho_do_registo(d["ficheiro"] or "")
        if p is not None and p.exists():
            quarentena.append((d, p))
    reproc = 0
    if quarentena:
        resp = input(f"{len(quarentena)} documento(s) em quarentena. Reler agora (OCR e palavras-passe; pode demorar)? (s/N): ")
        if resp.strip().lower() in ("s", "sim"):
            for i, (d, p) in enumerate(quarentena, 1):
                progresso_txt = f"\r  {GRAY}A reler {i}/{len(quarentena)}{RESET}"
                sys.stdout.write(progresso_txt)
                sys.stdout.flush()
                try:
                    dados, validado, motivo = dupla_verificacao_pdf(p)
                except Exception as e:
                    registar_log(f"Reprocessar {p.name}: {e}")
                    continue
                dados.pop("_antigo", None)
                if validado:
                    mudar = {"categoria": dados["Categoria"], "emissor": dados["Emissor"], "tipo_doc": dados["Tipo Doc"],
                             "entidade_mb": dados["Entidade MB"], "referencia_mb": dados["Referência MB"],
                             "valor": dados["Valor"], "data_limite": dados["Data limite"], "estado": dados["Estado"],
                             "notas": dados["Notas"]}
                    if dados.get("Data doc"):
                        mudar["data_doc"] = dados["Data doc"]
                    alteracoes.append((d["id"], d["ficheiro"], mudar, "relido"))
                    reproc += 1
                else:
                    nova = f"DIVERGÊNCIA DUPLA VERIFICAÇÃO: {motivo}"
                    if nova != (d["notas"] or ""):
                        alteracoes.append((d["id"], d["ficheiro"], {"notas": nova}, "relido"))
            sys.stdout.write("\r\033[2K")

    n_regra = sum(1 for a in alteracoes if a[3] == "regra")
    print(f"\n{C}{BOLD}=== REGRAS E REPROCESSAMENTO ==={RESET}")
    print(f"  Registos alterados por regra de nome: {W}{n_regra}{RESET} · relidos e identificados: {W}{reproc}{RESET}")
    for rid, rel, mudar, origem in alteracoes[:10]:
        print(f"   #{rid} {Path(rel).name[:45]}: " + ", ".join(f"{k}={v!r}" for k, v in list(mudar.items())[:4]))
    if not alteracoes:
        print("  Nada a alterar.")
        return
    if input("Escreve APLICAR para gravar na base de dados: ").strip() != "APLICAR":
        print("Cancelado.")
        return
    if not verificar_dupla_autenticacao(f"Aplicar regras e reprocessar: {len(alteracoes)} registo(s)"):
        return
    (LOGS_DIR / "backup_db").mkdir(parents=True, exist_ok=True)
    bk = LOGS_DIR / "backup_db" / f"Faturas_Shadow_77_antes_regras_{datetime.now().strftime('%Y%m%d_%H%M%S')}.db"
    origem_db, destino_db = sqlite3.connect(DB_FILE), sqlite3.connect(bk)
    origem_db.backup(destino_db)
    destino_db.close()
    origem_db.close()
    conn = ligar_bd()
    try:
        for rid, rel, mudar, origem in alteracoes:
            sets = ", ".join(f"{c} = ?" for c in mudar)        # nomes de colunas fixos, definidos acima
            conn.execute(f"UPDATE faturas SET {sets} WHERE id = ?", list(mudar.values()) + [rid])
        conn.commit()
    finally:
        conn.close()
    registar_log(f"Regras/reprocessamento: {len(alteracoes)} registo(s); cópia da BD: {bk.name}")
    sincronizar_bd_para_excel()
    print(f"{G}✓ {len(alteracoes)} registo(s) atualizado(s). Cópia da base de dados: {bk}{RESET}")
    print(f"  Para mover os ficheiros para as pastas certas (ano/mês/emissor): python comparador_faturas_shadow_77.py --aplicar")


def reconfigurar_seguranca():
    if _carregar_seguranca() is None:
        configurar_seguranca()
    elif verificar_dupla_autenticacao("Alterar PIN / app autenticadora"):
        configurar_seguranca()


def menu_gestao():
    inicializar_base_dados()
    while True:
        print(f"\n{C}{BOLD}=== GERIR REGISTOS ==={RESET}")
        print("  1) Listar últimos registos")
        print("  2) Pesquisar")
        print("  3) Adicionar registo manual")
        print("  4) Alterar registo   (2FA)")
        print("  5) Apagar registo    (2FA)")
        print("  6) Configurar PIN / app autenticadora")
        print("  7) Importar correções do Excel (2FA)")
        print("  8) Adicionar palavra-passe de PDF")
        print("  9) Aplicar regras de nome e reprocessar a quarentena (2FA)")
        print("  0) Voltar")
        op = input("Opção: ").strip()
        if op == "1":
            listar_registos()
        elif op == "2":
            listar_registos(input("Pesquisar por emissor/ficheiro/categoria: ").strip())
        elif op == "3":
            adicionar_registo()
        elif op == "4":
            alterar_registo()
        elif op == "5":
            apagar_registo()
        elif op == "6":
            reconfigurar_seguranca()
        elif op == "7":
            importar_correcoes_excel()
        elif op == "8":
            adicionar_password_pdf()
        elif op == "9":
            aplicar_regras_e_reprocessar()
        elif op == "0":
            return


# ---------------------------------------------------------------------------
# WHATSAPP: imagens e vídeos -> PDF
# ---------------------------------------------------------------------------
def _melhor_frame_video(caminho, n_frames=10):
    """
    Lê vários frames do vídeo, faz OCR a cada um e VOTA nos valores de entidade, referência e montante
    (o que mais frames dizem). Devolve o frame que melhor concorda com o resultado da votação.
    """
    if cv2 is None:
        raise RuntimeError("OpenCV não instalado (pip install opencv-python)")
    cap = cv2.VideoCapture(str(caminho))
    frames = []                      # (imagem, texto)
    try:
        total = int(cap.get(cv2.CAP_PROP_FRAME_COUNT) or 0)
        if total > n_frames:
            posicoes = sorted({int(total * (0.05 + 0.9 * i / (n_frames - 1))) for i in range(n_frames)})
        else:
            posicoes = list(range(max(total, 1)))
        for pos in posicoes:
            cap.set(cv2.CAP_PROP_POS_FRAMES, pos)
            ok, frame = cap.read()
            if not ok:
                continue
            img = Image.fromarray(cv2.cvtColor(frame, cv2.COLOR_BGR2RGB))
            try:
                texto = _ocr(preprocessar_para_ocr(img, ampliar=True))
            except Exception:
                texto = ""
            frames.append((img, texto))
    finally:
        cap.release()
    if not frames:
        return None

    votos = [{}, {}, {}]             # entidade, referência, valor
    for _, texto in frames:
        for i, campo in enumerate(_campos_basicos(texto)):
            if campo is not None:
                votos[i][campo] = votos[i].get(campo, 0) + 1
    consenso = [max(v, key=v.get) if v else None for v in votos]

    def pontuacao(texto):
        so_digitos_txt = re.sub(r"\D", "", texto)
        pts = 0
        if consenso[0] and consenso[0] in so_digitos_txt:
            pts += 1
        if consenso[1] and consenso[1] in so_digitos_txt:
            pts += 1
        if consenso[2] is not None and re.search(rf"(?<!\d){consenso[2]:.2f}".replace(".", "[.,]") + r"(?!\d)", texto):
            pts += 1
        return pts * 100000 + len(texto.strip())

    return max(frames, key=lambda f: pontuacao(f[1]))[0]


def converter_media_whatsapp_para_pdf():
    """Converte imagens/vídeos da pasta whatsapp (e subpastas) em PDF. Os originais vão para arquivo/media_originais."""
    WHATSAPP_DIR.mkdir(parents=True, exist_ok=True)
    ficheiros = [f for f in WHATSAPP_DIR.rglob("*")
                 if f.is_file() and f.suffix.lower() in (EXT_IMAGENS | EXT_VIDEOS)]
    if not ficheiros:
        info(f"      {W}-> Sem imagens/vídeos novos na pasta whatsapp.{RESET}")
        return

    info(f"      {W}-> A converter {len(ficheiros)} ficheiro(s) multimédia para PDF...{RESET}")
    convertidos = ignorados = falhas = 0
    for i, f in enumerate(ficheiros, 1):
        try:
            if f.suffix.lower() in EXT_IMAGENS:
                with Image.open(f) as im:
                    img = ImageOps.exif_transpose(im).convert("RGB")
                    img.load()
                if max(img.size) < MIN_LADO_IMAGEM:
                    ignorados += 1   # sticker / ícone / miniatura
                    continue
            else:
                img = _melhor_frame_video(f)
                if img is None:
                    raise ValueError("não foi possível ler o vídeo")

            destino = f.with_suffix(".pdf")
            if destino.exists():
                destino = f.with_name(f"{f.stem}_{datetime.now().strftime('%H%M%S%f')}.pdf")
            img.save(destino, "PDF", resolution=200.0)
            mover_sem_sobrescrever(f, ARQUIVO_DIR / "media_originais")
            convertidos += 1
        except Exception as e:
            falhas += 1
            registar_log(f"Erro ao converter {f.name}: {e}")
        finally:
            progresso("media", i / len(ficheiros))

    stat("media_convertidos", convertidos)
    info(f"      {G}• Convertidos: {convertidos}{RESET}")
    if ignorados:
        info(f"      {GRAY}• Ignorados (imagens pequenas, ex.: stickers): {ignorados}{RESET}")
    if falhas:
        aviso(f"{R}      • {falhas} ficheiro(s) do WhatsApp não puderam ser convertidos (ver log; vídeos precisam de opencv-python).{RESET}")


# ---------------------------------------------------------------------------
# CRUZAMENTO FATURAS <-> COMPROVATIVOS
# ---------------------------------------------------------------------------
def _juntar_nota(existente, nova):
    return f"{existente} | {nova}" if existente else nova


def _data_como_datetime(valor):
    partes = _parse_data_flex(valor)
    if not partes:
        return None
    return datetime(partes[0], partes[1] or 1, partes[2] or 1)


def _pagamento_data_compativel(comprovativo, fatura):
    pagamento = _data_como_datetime(comprovativo[7])
    emissao = _data_como_datetime(fatura[6])
    return pagamento is not None and (emissao is None or pagamento >= emissao)


def _escolher_fatura_por_data(comprovativo, candidatos):
    if len(candidatos) == 1:
        return candidatos[0]
    pagamento = _data_como_datetime(comprovativo[7])
    if pagamento is None:
        return None

    dentro_do_prazo = []
    mesmo_mes_emissao = []
    for fatura in candidatos:
        emissao = _data_como_datetime(fatura[6])
        limite = _data_como_datetime(fatura[7])
        if emissao and limite and emissao <= pagamento <= limite:
            dentro_do_prazo.append(fatura)
        if emissao and (emissao.year, emissao.month) == (pagamento.year, pagamento.month):
            mesmo_mes_emissao.append(fatura)
    if len(dentro_do_prazo) == 1:
        return dentro_do_prazo[0]
    if len(mesmo_mes_emissao) == 1:
        return mesmo_mes_emissao[0]
    return None


def conciliar_pagamentos():
    """
    Associa cada comprovativo a uma fatura por pagar (um comprovativo paga no máximo uma fatura):
      1) mesma referência MB (e entidade, se ambas existirem)  -> 'Pago' (ou 'Pago (a confirmar)' se o valor diferir)
      2) mesmo valor (+ mesmo emissor, ou único candidato)      -> 'Pago (a confirmar)'
    É idempotente: as ligações ficam guardadas na base de dados.
    """
    conn = ligar_bd()
    confirmados = a_confirmar = 0
    try:
        comps = conn.execute(
            "SELECT id, emissor, entidade_mb, referencia_mb, valor, ficheiro, notas, data_doc FROM faturas "
            "WHERE tipo_doc IN ('Comprovativo', 'Transferencia') AND estado != 'Ok' AND (ligado_a IS NULL OR "
            "NOT EXISTS (SELECT 1 FROM faturas x WHERE x.id = faturas.ligado_a)) ORDER BY id").fetchall()
        pend = conn.execute(
            "SELECT id, emissor, entidade_mb, referencia_mb, valor, notas, data_doc, data_limite FROM faturas "
            "WHERE estado = 'Por pagar' AND tipo_doc IN (?, ?) AND ligado_a IS NULL ORDER BY id",
            TIPOS_A_PAGAR).fetchall()

        usados, ligados = set(), set()

        def ligar(c, f, estado, motivo):
            nonlocal confirmados, a_confirmar
            conn.execute("UPDATE faturas SET estado = ?, ligado_a = ?, notas = ? WHERE id = ?",
                         (estado, c[0], _juntar_nota(f[5], f"Comprovativo: {c[5]} ({motivo})"), f[0]))
            conn.execute("UPDATE faturas SET ligado_a = ?, notas = ? WHERE id = ?",
                         (f[0], _juntar_nota(c[6], f"Paga fatura de {f[1]} {f[4]:.2f}€ ({motivo})"), c[0]))
            usados.add(f[0])
            ligados.add(c[0])
            if estado == "Pago":
                confirmados += 1
            else:
                a_confirmar += 1
            registar_log(f"Cruzamento: comprovativo {c[0]} -> fatura {f[0]} [{estado}] ({motivo})")

        # A referência e a data devem apontar para uma única fatura.
        for c in comps:
            ref = limpar_ref(c[3])
            if not ref or not c[7]:
                continue
            candidatos = []
            for f in pend:
                if f[0] in usados or limpar_ref(f[3]) != ref:
                    continue
                if c[2] and f[2] and limpar_ref(c[2]) != limpar_ref(f[2]):
                    continue
                if not _pagamento_data_compativel(c, f):
                    continue
                candidatos.append(f)
            f = _escolher_fatura_por_data(c, candidatos)
            if f is None:
                continue
            valor_igual = c[4] is not None and f[4] is not None and abs(c[4] - f[4]) <= 0.009
            entidade_igual = bool(c[2] and f[2] and limpar_ref(c[2]) == limpar_ref(f[2]))
            if valor_igual and entidade_igual:
                ligar(c, f, "Pago", "entidade, referência, montante e data compatíveis")
            else:
                ligar(c, f, "Pago (a confirmar)", "referência/data compatíveis; confirmar entidade ou montante")

        # Sem referência, só associa um candidato único por emissor, valor e data.
        for c in comps:
            if c[0] in ligados or not c[4]:
                continue
            if not c[7]:
                continue
            candidatos = []
            for f in pend:
                if f[0] in usados or not f[4] or abs(f[4] - c[4]) > 0.009:
                    continue
                if limpar_ref(c[3]) and limpar_ref(f[3]) and limpar_ref(c[3]) != limpar_ref(f[3]):
                    continue   # referências diferentes = pagamentos diferentes
                if c[1] not in EMISSORES_GENERICOS and c[1] != f[1]:
                    continue
                if not _pagamento_data_compativel(c, f):
                    continue
                candidatos.append(f)
            f = _escolher_fatura_por_data(c, candidatos)
            if f is not None:
                motivo = "emissor, montante e data compatíveis (confirmar referência)"
                ligar(c, f, "Pago (a confirmar)", motivo)

        conn.commit()
    finally:
        conn.close()

    stat("pagos", confirmados)
    stat("a_confirmar", a_confirmar)
    info(f"      {G}• Faturas marcadas como pagas: {confirmados}{RESET}")
    info(f"      {Y}• Pagas a confirmar: {a_confirmar}{RESET}")


EMISSORES_EM_EXTRATO = ("Via Verde",)      # emissores cobrados por débito/cartão: confirmam-se no extrato do banco


def confirmar_em_extratos():
    """
    Faturas por pagar de EMISSORES_EM_EXTRATO (Via Verde): procura nos extratos do Millennium BCP (ficheiros EXT/EXTRATO/
    CARTAO) uma linha com o nome do emissor e o mesmo valor. Se encontrar, a fatura passa a 'Pago'.
    Só funciona com extratos de texto legível (PDF sem imagem e sem palavra-passe por descobrir).
    """
    conn = ligar_bd()
    try:
        marcas = ",".join("?" * len(EMISSORES_EM_EXTRATO))
        faturas = conn.execute(
            f"SELECT id, emissor, valor FROM faturas WHERE estado = 'Por pagar' AND tipo_doc = 'Fatura' "
            f"AND emissor IN ({marcas}) AND valor > 0", EMISSORES_EM_EXTRATO).fetchall()
        if not faturas:
            return
        extratos = conn.execute(
            "SELECT id, ficheiro FROM faturas WHERE emissor = 'Millennium BCP' AND "
            "(UPPER(ficheiro) LIKE '%EXT%' OR UPPER(ficheiro) LIKE '%CARTAO%')").fetchall()
    finally:
        conn.close()
    textos = {}
    for eid, rel in extratos:
        p = caminho_do_registo(rel or "")
        if p is not None and p.exists():
            t = texto_pdf_completo(p)
            if t.strip():
                textos[eid] = (Path(rel).name, t.splitlines())
    if not textos:
        info(f"      {GRAY}• Sem extratos legíveis do Millennium BCP para confirmar {len(faturas)} fatura(s).{RESET}")
        return

    confirmadas = 0
    conn = ligar_bd()
    try:
        for fid, emissor, valor in faturas:
            formatos = {f"{valor:.2f}", f"{valor:.2f}".replace(".", ",")}
            rx_valor = re.compile(r"(?<![\d.,])(?:" + "|".join(re.escape(f) for f in formatos) + r")(?!\d)")
            rx_nome = re.compile(re.sub(r"\\ ", r"\\s*", re.escape(emissor)), re.IGNORECASE)
            for eid, (nome, linhas) in textos.items():
                achou = False
                for i, linha in enumerate(linhas):
                    janela = linha + " " + (linhas[i + 1] if i + 1 < len(linhas) else "")
                    if rx_nome.search(linha) and rx_valor.search(janela):
                        achou = True
                        break
                if achou:
                    conn.execute("UPDATE faturas SET estado = 'Pago', ligado_a = ?, "
                                 "notas = COALESCE(notas || ' | ', '') || ? WHERE id = ?",
                                 (eid, f"Confirmado no extrato {nome}", fid))
                    confirmadas += 1
                    registar_log(f"Fatura {fid} ({emissor} {valor:.2f}€) confirmada no extrato {nome}")
                    break
        conn.commit()
    finally:
        conn.close()
    stat("confirmados_extrato", confirmadas)
    info(f"      {G}• Confirmadas nos extratos do banco: {confirmadas}{RESET}")


def cruzar_pagamentos():
    conciliar_pagamentos()
    confirmar_em_extratos()


def relatorio_pagamentos():
    conn = ligar_bd()
    try:
        por_pagar = conn.execute(
            "SELECT id, emissor, valor, data_limite, ficheiro FROM faturas WHERE estado = 'Por pagar' "
            "AND tipo_doc IN (?, ?) ORDER BY id", TIPOS_A_PAGAR).fetchall()
        confirmar = conn.execute(
            "SELECT id, emissor, valor, notas FROM faturas WHERE estado = 'Pago (a confirmar)' ORDER BY id").fetchall()
        sem_fatura = conn.execute(
            "SELECT id, valor, referencia_mb, ficheiro FROM faturas WHERE tipo_doc IN ('Comprovativo', 'Transferencia') AND estado != 'Ok' AND "
            "(ligado_a IS NULL OR NOT EXISTS (SELECT 1 FROM faturas x WHERE x.id = faturas.ligado_a)) ORDER BY id").fetchall()
    finally:
        conn.close()

    print(f"\n{C}{BOLD}=== ESTADO DOS PAGAMENTOS ==={RESET}")
    if not por_pagar and not confirmar:
        print(f"{G}{BOLD}✓ Está tudo pago.{RESET}")
    if por_pagar:
        total = sum((r[2] or 0) for r in por_pagar)
        print(f"{R}{BOLD}Por pagar: {len(por_pagar)} fatura(s), total {total:,.2f}€{RESET}")
        for r in por_pagar:
            print(f"   #{r[0]:<4} {str(r[1])[:30]:<30} {(r[2] or 0):>9.2f}€   limite: {r[3] or '-'}   {r[4]}")
    if confirmar:
        print(f"{Y}{BOLD}Pagas a confirmar: {len(confirmar)}{RESET}")
        for r in confirmar:
            print(f"   #{r[0]:<4} {str(r[1])[:30]:<30} {(r[2] or 0):>9.2f}€   {str(r[3] or '')[-90:]}")
    if sem_fatura:
        print(f"{GRAY}Comprovativos sem fatura associada: {len(sem_fatura)}{RESET}")
        for r in sem_fatura:
            print(f"   #{r[0]:<4} {(r[1] or 0):>9.2f}€   ref: {r[2] or '-'}   {r[3]}")


# ---------------------------------------------------------------------------
# FLUXO PRINCIPAL
# ---------------------------------------------------------------------------
def auditar_documentos():
    for pasta in (ENTRADA_DIR, WHATSAPP_DIR, POR_IDENTIFICAR_DIR):
        pasta.mkdir(parents=True, exist_ok=True)
    pdfs = list(ENTRADA_DIR.glob("*.[pP][dD][fF]")) + [p for p in WHATSAPP_DIR.rglob("*") if p.suffix.lower() == ".pdf"]
    total = len(pdfs)
    if total == 0:
        info(f"      {W}-> Nenhum PDF pendente para auditar.{RESET}")
        return
    estado = {"atual": 0, "total": total, "concluido": False, "duplicados": 0, "quarentena": 0}
    preencher_hashes_em_falta()      # registos de versões antigas passam a ser comparáveis por conteúdo
    tarefa_auditoria_background(pdfs, estado)
    validados = total - estado["duplicados"] - estado["quarentena"]
    stat("validados", validados)
    stat("duplicados", estado["duplicados"])
    stat("quarentena", estado["quarentena"])
    info(f"      {G}• Validados com sucesso: {validados}{RESET}")
    info(f"      {Y}• Duplicados descartados: {estado['duplicados']}{RESET}")
    if estado["quarentena"]:
        info(f"      {R}• Quarentena (divergência): {estado['quarentena']}{RESET}")


def _fase(nome, rotulo, funcao, *args):
    """Corre uma fase; um erro numa fase não impede as seguintes."""
    info(f"{B}{rotulo}{RESET}")
    try:
        funcao(*args)
    except Exception as e:
        registar_log(f"Erro na fase '{nome}': {e}")
        aviso(f"{R}      -> Erro na fase '{nome}': {e}{RESET}")
    finally:
        progresso(nome, 1.0)


def sincronizar_lembretes_calendario(creds):
    """Cria ou atualiza lembretes das faturas por pagar no calendário ligado."""
    conta = ler_integracao()
    if not conta:
        return
    conn = ligar_bd()
    try:
        pendentes = conn.execute(
            "SELECT id, emissor, valor, data_limite FROM faturas "
            "WHERE estado = 'Por pagar' AND data_limite IS NOT NULL AND data_limite != ''"
        ).fetchall()
        guardados = {
            row[0]: (row[1], row[2])
            for row in conn.execute("SELECT invoice_id, provider, event_id FROM calendar_events")
        }
    finally:
        conn.close()

    eventos = {}
    if conta["provider"] == "google":
        if not creds:
            aviso(f"{Y}      -> Google Calendar ignorado: autenticação Google indisponível.{RESET}")
            return
        service = build("calendar", "v3", credentials=creds)
        for invoice_id, emissor, valor, data_limite in pendentes:
            parsed = _parse_data_flex(data_limite)
            if not parsed or not all(parsed):
                registar_log(f"Lembrete ignorado: data limite inválida no registo {invoice_id}: {data_limite!r}")
                continue
            due = date(*parsed)
            body = {
                "summary": f"Fatura por pagar — {emissor}",
                "description": f"Valor: {valor if valor is not None else 'não indicado'} €",
                "start": {"date": due.isoformat()},
                "end": {"date": (due + timedelta(days=1)).isoformat()},
                "reminders": {"useDefault": False, "overrides": [{"method": "popup", "minutes": 1440}]},
                "extendedProperties": {"private": {"faturas_shadow_77_id": str(invoice_id)}},
            }
            provider, event_id = guardados.get(invoice_id, (None, None))
            if provider != "google":
                event_id = None
            try:
                if event_id:
                    try:
                        service.events().update(
                            calendarId="primary", eventId=event_id, body=body
                        ).execute()
                    except HttpError as e:
                        if getattr(e, "status_code", None) != 404 and getattr(getattr(e, "resp", None), "status", None) != 404:
                            raise
                        event_id = None
                if not event_id:
                    created = service.events().insert(calendarId="primary", body=body).execute()
                    event_id = created["id"]
                eventos[invoice_id] = ("google", event_id)
            except HttpError as e:
                registar_log(f"Erro ao sincronizar lembrete Google da fatura {invoice_id}: {e}")
    else:
        import keyring
        from caldav import DAVClient
        from caldav.lib.error import NotFoundError
        from icalendar import Alarm, Calendar, Event
        password = keyring.get_password(KEYRING_SERVICE, conta["email"])
        if not password:
            aviso(f"{Y}      -> iCloud Calendar ignorado: credencial segura indisponível.{RESET}")
            return
        client = DAVClient(
            url="https://caldav.icloud.com",
            username=conta["email"],
            password=password,
            timeout=30,
        )
        calendars = client.principal().calendars()
        if not calendars:
            aviso(f"{Y}      -> A conta iCloud não disponibilizou um calendário.{RESET}")
            return
        calendar = calendars[0]
        for invoice_id, emissor, valor, data_limite in pendentes:
            parsed = _parse_data_flex(data_limite)
            if not parsed or not all(parsed):
                registar_log(f"Lembrete ignorado: data limite inválida no registo {invoice_id}: {data_limite!r}")
                continue
            due = date(*parsed)
            uid = f"faturas-shadow-77-{invoice_id}@local"
            ical = Calendar()
            ical.add("prodid", "-//Faturas Shadow_77//Invoice reminders//EN")
            ical.add("version", "2.0")
            event = Event()
            event.add("uid", uid)
            event.add("dtstamp", datetime.now().astimezone())
            event.add("summary", f"Fatura por pagar — {emissor}")
            event.add("description", f"Valor: {valor if valor is not None else 'não indicado'} €")
            event.add("dtstart", due)
            event.add("dtend", due + timedelta(days=1))
            alarm = Alarm()
            alarm.add("action", "DISPLAY")
            alarm.add("description", "Lembrete de pagamento de fatura")
            alarm.add("trigger", timedelta(days=-1))
            event.add_component(alarm)
            ical.add_component(event)
            try:
                try:
                    existing = calendar.event_by_uid(uid)
                    existing.data = ical.to_ical().decode("utf-8")
                    existing.save()
                except NotFoundError:
                    calendar.add_event(ical=ical.to_ical().decode("utf-8"))
                eventos[invoice_id] = ("icloud", uid)
            except Exception as e:
                registar_log(f"Erro ao sincronizar lembrete iCloud da fatura {invoice_id}: {e}")

    if conta["provider"] == "google":
        active_ids = set(eventos)
        for invoice_id, (provider, event_id) in guardados.items():
            if invoice_id not in active_ids and provider == "google":
                try:
                    service.events().delete(calendarId="primary", eventId=event_id).execute()
                except HttpError as e:
                    if getattr(e, "status_code", None) != 404 and getattr(getattr(e, "resp", None), "status", None) != 404:
                        registar_log(f"Erro ao remover lembrete Google da fatura {invoice_id}: {e}")
    else:
        for invoice_id, (provider, event_id) in guardados.items():
            if invoice_id not in eventos and provider == "icloud":
                try:
                    calendar.event_by_uid(event_id).delete()
                except NotFoundError:
                    pass
                except Exception as e:
                    registar_log(f"Erro ao remover lembrete iCloud da fatura {invoice_id}: {e}")

    conn = ligar_bd()
    try:
        for invoice_id, (provider, event_id) in eventos.items():
            conn.execute(
                "INSERT INTO calendar_events(invoice_id, provider, event_id, updated_at) VALUES (?, ?, ?, CURRENT_TIMESTAMP) "
                "ON CONFLICT(invoice_id) DO UPDATE SET provider=excluded.provider, event_id=excluded.event_id, updated_at=CURRENT_TIMESTAMP",
                (invoice_id, provider, event_id),
            )
        for invoice_id in set(guardados) - set(eventos):
            conn.execute("DELETE FROM calendar_events WHERE invoice_id = ?", (invoice_id,))
        conn.commit()
    finally:
        conn.close()
    info(f"      {G}• Lembretes de calendário sincronizados: {len(eventos)}{RESET}")


def _pipeline(creds):
    try:
        _fase("gmail", "[1/6] Email", descarregar_e_gerir_emails_gmail, creds)
        _fase("media", "[2/6] WhatsApp: imagens/vídeos -> PDF", converter_media_whatsapp_para_pdf)
        _fase("auditoria", "[3/6] Auditoria dos documentos", auditar_documentos)
        _fase("cruzamento", "[4/6] Cruzamento faturas <-> comprovativos e extratos", cruzar_pagamentos)
        _fase("calendario", "[5/6] Calendário: lembretes de vencimento", sincronizar_lembretes_calendario, creds)
        _fase("excel", "[6/6] Excel", sincronizar_bd_para_excel)
    finally:
        PROG["concluido"] = True


def mostrar_resumo(segundos):
    g = STATS.get
    print(f"{G}{BOLD}✔ Concluído em {segundos:.0f}s{RESET}")
    print(f"  {W}Email:{RESET}      {g('pdfs_gmail', 0)} PDF(s) descarregado(s) · mensagens mantidas")
    print(f"  {W}WhatsApp:{RESET}   {g('media_convertidos', 0)} imagem(ns)/vídeo(s) convertido(s) para PDF")
    cor_q = R if g('quarentena', 0) else RESET
    print(f"  {W}Auditoria:{RESET}  {g('validados', 0)} validado(s) · {g('duplicados', 0)} duplicado(s) · {cor_q}{g('quarentena', 0)} em quarentena{RESET}")
    print(f"  {W}Pagamentos:{RESET} {g('pagos', 0)} fatura(s) paga(s) · {g('a_confirmar', 0)} a confirmar · "
          f"{g('confirmados_extrato', 0)} confirmada(s) em extratos do banco")
    for a in AVISOS:
        print(f"  {Y}⚠ {a.strip()}{RESET}")


def processar_faturas(auto=False):
    """
    No terminal mostra só o cabeçalho, o dino e a barra de progresso; os passos correm em segundo plano
    e o resumo aparece no fim. Fora de um terminal (tarefa agendada) escreve os passos em texto simples.
    auto=True: nunca abre a autenticação Google nem faz perguntas.
    """
    global SILENCIOSO
    resetar_estado()
    tty = sys.stdout.isatty()
    animacao_abertura(limpar=tty)
    inicializar_base_dados()
    inicio = time.time()

    SILENCIOSO = tty
    try:
        verificar_tesseract()
        conta = ler_integracao()
        if not conta and not auto:
            configurar_integracao()
            conta = ler_integracao()
        creds = (
            obter_credenciais_google(interativo=not auto)
            if conta and conta["provider"] == "google"
            else None
        )
        trabalho = threading.Thread(target=_pipeline, args=(creds,), daemon=True)
        trabalho.start()
        if tty:
            animacao_dino(PROG)
        trabalho.join()
    finally:
        SILENCIOSO = False

    if tty:
        mostrar_resumo(time.time() - inicio)
    relatorio_pagamentos()
    print()
    registar_log(f"Processamento {'AUTO ' if auto else ''}concluído: {dict(STATS)}")


def menu_principal():
    conta = ler_integracao()
    if not conta:
        conta = configurar_integracao()
        conta = ler_integracao() if conta else None
    if conta and conta["provider"] == "google":
        obter_credenciais_google(interativo=True)
    while True:
        print(f"\n{C}{BOLD}=== FATURAS SHADOW_77 v9.7 ==={RESET}")
        print(f"  1) {t('processar')}")
        print(f"  2) {t('automatico')}")
        print(f"  3) {t('gestao')}")
        print(f"  4) {t('pagamentos')}")
        print(f"  5) {t('email_limpeza')}")
        print(f"  0) {t('sair')}")
        op = input(f"{t('opcao')}: ").strip().lower()
        if op == "1":
            processar_faturas()
        elif op in ("2", "auto", "a"):
            processar_faturas(auto=True)
        elif op == "3":
            menu_gestao()
        elif op == "4":
            inicializar_base_dados()
            relatorio_pagamentos()
        elif op == "5":
            confirmacao = "BACKUP E LIMPAR"
            if input(f"{t('confirmar_email')} {t('escreve')} {confirmacao}: ").strip() != confirmacao:
                print(f"{Y}{t('cancelado')}{RESET}")
                continue
            creds = obter_credenciais_google(interativo=True)
            if creds:
                inicializar_base_dados()
                arquivar_e_apagar_emails_sem_estrela(creds)
        elif op == "0":
            return


def modo_auto(cada_minutos=0):
    """Executa o processo completo. Com --cada N repete de N em N minutos (Ctrl+C para parar)."""
    while True:
        try:
            processar_faturas(auto=True)
        except Exception as e:
            print(f"{R}[ERRO NO MODO AUTO]: {e}{RESET}")
            registar_log(f"Erro no modo AUTO: {e}")
        if not cada_minutos:
            return
        print(f"{GRAY}Próxima execução em {cada_minutos} min (Ctrl+C para parar)...{RESET}")
        time.sleep(cada_minutos * 60)


if __name__ == "__main__":
    try:
        if "--auto" in sys.argv:
            cada = 0
            if "--cada" in sys.argv:
                try:
                    cada = max(1, int(sys.argv[sys.argv.index("--cada") + 1]))
                except (IndexError, ValueError):
                    print(f"{R}Uso: python faturas_shadow_77.py --auto [--cada MINUTOS]{RESET}")
                    sys.exit(1)
            modo_auto(cada)
        else:
            menu_principal()
    except KeyboardInterrupt:
        sys.stdout.write("\033[?25h")
        print(f"\n{Y}Interrompido.{RESET}")
    except Exception as e:
        sys.stdout.write("\033[?25h")
        print(f"\n{R}[ERRO CRÍTICO NO SCRIPT]: {e}{RESET}")
        registar_log(f"Erro crítico: {e}")
