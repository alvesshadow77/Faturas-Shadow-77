import hashlib
import json
import os
import shutil
import sqlite3
import sys
from datetime import datetime
from pathlib import Path

import pdfplumber


_DADOS_POR_UTILIZADOR = Path(
    os.environ.get(
        "LOCALAPPDATA",
        str(Path.home() / "AppData" / "Local"),
    )
) / "Faturas Shadow_77"
BASE_DIR = Path(os.environ.get("FATURAS_SHADOW_77_DIR", str(_DADOS_POR_UTILIZADOR)))
DB_FILE = BASE_DIR / "Faturas Shadow_77.db"
EXCEL_FILE = BASE_DIR / "Faturas_2026.xlsx"
ARQUIVO_DIR = BASE_DIR / "arquivo"
REMOVIDOS_DIR = ARQUIVO_DIR / "removidos"
BACKUP_DIR = BASE_DIR / "logs" / "backup_pre_limpeza"

TIPOS_FISCAL = {
    "Fatura",
    "Fatura-Recibo",
    "Nota de Credito",
    "Nota de Crédito",
    "Aviso de Cobranca / Condominio",
    "Aviso de Cobrança / Condomínio",
    "Comprovativo",
    "Transferencia",
    "Transferência",
}


def caminho_real(rel):
    rel = str(rel or "").replace("\\", "/").strip()
    if not rel or rel.startswith("manual/"):
        return None
    if rel.lower().startswith("por identificar/"):
        return BASE_DIR / rel
    return ARQUIVO_DIR / rel


def sha256(path):
    h = hashlib.sha256()
    with path.open("rb") as f:
        for block in iter(lambda: f.read(1024 * 1024), b""):
            h.update(block)
    return h.hexdigest()


def texto_pdf(path):
    if not path or not path.exists():
        return ""
    try:
        with pdfplumber.open(path) as pdf:
            return "\n".join((page.extract_text() or "") for page in pdf.pages[:8]).lower()
    except Exception:
        return ""


def nome_claro(row, motivo):
    emissor = str(row["emissor"] or "Documento").strip()
    tipo = str(row["tipo_doc"] or "Documento").strip()
    valor = row["valor"]
    valor_nome = f"_{valor:.2f}EUR".replace(".", "-") if isinstance(valor, (int, float)) else ""
    origem = Path(str(row["ficheiro"] or "documento")).stem
    prefixo = {
        "Publicidade": "PUBLICIDADE",
        "Duplicado": "DUPLICADO",
        "Nao fiscal": "NAO_FISCAL",
        "Operador incorreto": "OPERADOR_ERRADO",
    }.get(motivo, "REMOVIDO")
    base = f"{prefixo}_{emissor}_{tipo}_{origem}{valor_nome}".replace("/", "-")
    return "".join(ch if ch.isalnum() or ch in " ._-()" else "_" for ch in base)[:180]


def destino_livre(path):
    if not path.exists():
        return path
    i = 2
    while True:
        candidate = path.with_name(f"{path.stem}_{i}{path.suffix}")
        if not candidate.exists():
            return candidate


def analisar():
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    rows = conn.execute("SELECT * FROM faturas").fetchall()

    # Identifica duplicados pelo hash do ficheiro.
    hash_to_keep = {}
    hash_candidates = {}
    for row in rows:
        path = caminho_real(row["ficheiro"])
        if path and path.exists():
            try:
                digest = sha256(path)
                hash_candidates.setdefault(digest, []).append(row)
            except OSError:
                pass

    for candidates in hash_candidates.values():
        if len(candidates) > 1:
            keep = sorted(
                candidates,
                key=lambda r: (
                    r["estado"] not in ("Por Identificar", "Informativo"),
                    r["emissor"] not in ("", "Desconhecido"),
                    r["valor"] or 0,
                ),
                reverse=True,
            )[0]
            hash_to_keep[sha256(caminho_real(keep["ficheiro"]))] = keep["id"]

    selected = []
    for row in rows:
        path = caminho_real(row["ficheiro"])
        rel = str(row["ficheiro"] or "").replace("\\", "/")
        lower_emissor = str(row["emissor"] or "").lower()
        tipo = str(row["tipo_doc"] or "")
        lower_name = rel.lower()
        text = ""

        motivo = None
        detalhe = ""

        if path and path.exists():
            try:
                digest = sha256(path)
                if digest in hash_to_keep and hash_to_keep[digest] != row["id"]:
                    motivo = "Duplicado"
                    detalhe = f"Duplicado pelo SHA-256 do registo #{hash_to_keep[digest]}"
            except OSError:
                pass

        if motivo is None and "arquivo/duplicados/" in lower_name:
            motivo = "Duplicado"
            detalhe = "Ficheiro já está na pasta arquivo/duplicados"

        if motivo is None and (tipo not in TIPOS_FISCAL or tipo == "Diversos"):
            motivo = "Nao fiscal"
            detalhe = f"Tipo documental nao fiscal: {tipo or 'vazio'}"

        if motivo is None and ("nos" in lower_emissor or "digi" in lower_emissor or "/nos/" in lower_name or "/digi/" in lower_name):
            motivo = "Operador incorreto"
            detalhe = "Emissor/pasta NOS/DIGI confirmado como incorreto ou nao pertence ao utilizador"

        publicidade = [
            "comissoes para clientes particulares",
            "termos e condicoes",
            "termos e condições",
            "newsletter",
            "publicidade",
        ]
        if motivo is None and (tipo not in TIPOS_FISCAL or tipo == "Diversos"):
            text = texto_pdf(path).lower() if path else ""
        if motivo is None and any(term in text for term in publicidade):
            motivo = "Publicidade"
            detalhe = "Conteudo identificado como publicidade, condicoes ou newsletter"

        if motivo:
            selected.append((row, motivo, detalhe))

    conn.close()
    return selected


def aplicar():
    run_id = datetime.now().strftime("%Y%m%d_%H%M%S")
    run_dir = REMOVIDOS_DIR / run_id
    run_dir.mkdir(parents=True, exist_ok=True)

    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    db_backup = BACKUP_DIR / f"Faturas_Shadow_77_{run_id}.db"
    excel_backup = BACKUP_DIR / f"Faturas_2026_{run_id}.xlsx"
    shutil.copy2(DB_FILE, db_backup)
    if EXCEL_FILE.exists():
        shutil.copy2(EXCEL_FILE, excel_backup)

    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    conn.execute("""
        CREATE TABLE IF NOT EXISTS documentos_removidos (
            run_id TEXT,
            orig_id INTEGER,
            categoria TEXT,
            emissor TEXT,
            tipo_doc TEXT,
            entidade_mb TEXT,
            referencia_mb TEXT,
            valor REAL,
            data_limite TEXT,
            estado TEXT,
            origem TEXT,
            ficheiro TEXT,
            notas TEXT,
            data_doc TEXT,
            motivo TEXT,
            detalhe TEXT,
            destino TEXT,
            removido_em TEXT
        )
    """)

    selected = analisar()
    for row, motivo, detalhe in selected:
        source = caminho_real(row["ficheiro"])
        destination = None
        if source and source.exists():
            suffix = source.suffix.lower()
            descriptive = nome_claro(row, motivo)
            destination = destino_livre(run_dir / f"{descriptive}{suffix}")
            shutil.move(str(source), str(destination))

        conn.execute("""
            INSERT INTO documentos_removidos (
                run_id, orig_id, categoria, emissor, tipo_doc, entidade_mb,
                referencia_mb, valor, data_limite, estado, origem, ficheiro,
                notas, data_doc, motivo, detalhe, destino, removido_em
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            run_id, row["id"], row["categoria"], row["emissor"], row["tipo_doc"],
            row["entidade_mb"], row["referencia_mb"], row["valor"], row["data_limite"],
            row["estado"], row["origem"], row["ficheiro"], row["notas"], row["data_doc"],
            motivo, detalhe,
            destination.relative_to(BASE_DIR).as_posix() if destination else None,
            datetime.now().isoformat(timespec="seconds"),
        ))
        conn.execute("DELETE FROM faturas WHERE id = ?", (row["id"],))

    conn.commit()
    conn.close()

    manifest = {
        "run_id": run_id,
        "data": datetime.now().isoformat(timespec="seconds"),
        "total": len(selected),
        "itens": [
            {
                "orig_id": row["id"],
                "motivo": motivo,
                "ficheiro": row["ficheiro"],
                "detalhe": detalhe,
            }
            for row, motivo, detalhe in selected
        ],
    }
    (run_dir / "manifesto.json").write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    sys.path.insert(0, str(BASE_DIR / "scripts"))
    import faturas_shadow_77 as main
    main.DB_FILE = DB_FILE
    main.EXCEL_FILE = EXCEL_FILE
    main.inicializar_base_dados()
    main.sincronizar_bd_para_excel()

    print(f"Removidos: {len(selected)}")
    print(f"Pasta: {run_dir}")
    print(f"Backups: {db_backup}")
    if EXCEL_FILE.exists():
        print(f"Excel atualizado: {EXCEL_FILE}")


if __name__ == "__main__":
    if len(sys.argv) == 2 and sys.argv[1] == "--aplicar":
        aplicar()
    else:
        selected = analisar()
        counts = {}
        for _, motivo, _ in selected:
            counts[motivo] = counts.get(motivo, 0) + 1
        print("Sem aplicar alterações.")
        print(dict(sorted(counts.items())))
        print("Total:", len(selected))
