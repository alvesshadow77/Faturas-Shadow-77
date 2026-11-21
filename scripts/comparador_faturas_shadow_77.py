"""
FATURAS SHADOW_77 - COMPARADOR

Analisa TODOS os PDFs da pasta de faturas (D:\\Faturas), deteta os repetidos e põe-nos de lado
(nunca apaga nada) antes de se atualizar o Faturas Shadow_77.

Como decide que dois ficheiros são repetidos
  1) IDÊNTICOS: mesmo conteúdo (SHA-256). Confiança alta.
  2) MESMA FATURA / MESMO PAGAMENTO: mesma classe de documento + mesma REFERÊNCIA MB, com ENTIDADE,
     VALOR e EMISSOR compatíveis (e datas compatíveis, quando existem).
        - confiança ALTA  : a data confirma (mesma data limite, ou fotos/vídeos com datas no nome a <= 1 dia)
        - confiança MÉDIA : não há data para confirmar (só se movem se o confirmares)
     Uma fatura e o seu comprovativo NÃO são repetidos: são um par (aparecem no separador "Pares").
     Mesma referência com valores/entidades/emissores diferentes, ou datas diferentes (pagamentos
     recorrentes), vão para "Conflitos" e não são tocados.

Os dados de cada ficheiro (emissor, entidade, referência, valor...) vêm, por esta ordem, de:
  Excel corrigido (Faturas_2026.xlsx) > base de dados (Faturas Shadow_77.db) > sem dados (só se compara o conteúdo).

SEPARAR POR ANO (precisa de faturas_shadow_77.py na mesma pasta, porque usa as mesmas regras de datas):
  cada documento recebe a sua data (QR, "data de emissão", data no nome do ficheiro, data limite) e passa
  para  arquivo/<ano>/<mês>/<emissor>/.  Tudo o que for anterior a 2026 fica "Ok" (se estava por pagar,
  por identificar ou a confirmar) e deixa de ser verificado: só se verifica 2026 e os anos seguintes.
  Se QUALQUER data do documento for de 2026 ou posterior, ele continua a ser verificado. Documentos sem
  nenhuma data ficam onde estão (separador "Ano desconhecido").

Registos sem ficheiro no sítio certo (a base diz uma coisa, o disco outra):
  o comparador RE-LIGA cada registo ao ficheiro com o mesmo nome de base (ignora o prefixo de hora, ex.
  "192942_" ou "213029360460_") quando o conteúdo é o mesmo, fica com os dados (incl. as tuas correções do
  Excel) e, ao aplicar, REPÕE o ficheiro no caminho registado do arquivo. Linhas repetidas da base de dados
  (mesmo documento registado 2x) também passam para faturas_repetidas.

Uso:
  python comparador_faturas_shadow_77.py                  menu
  python comparador_faturas_shadow_77.py --analisar       só relatório (não mexe em nada)
  python comparador_faturas_shadow_77.py --aplicar        põe de lado os repetidos (pede PIN + app autenticadora)
  python comparador_faturas_shadow_77.py --desfazer       repõe a última aplicação
  opções: --excel CAMINHO | --sem-excel | --incluir-medios

Nada é apagado: os repetidos vão para D:\\Faturas\\repetidos\\<data-hora>\\ (mesma estrutura de pastas),
as linhas da base de dados passam para a tabela faturas_repetidas, e há cópia da base de dados em
logs\\backup_db. Tudo se desfaz com --desfazer.
"""
import argparse
import hashlib
import json
import os
import re
import shutil
import sqlite3
import sys
import time
from collections import defaultdict
from datetime import datetime, timedelta
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass
os.system('')  # cores ANSI no Windows

R, G, Y, B, C, W = "\033[91m", "\033[92m", "\033[93m", "\033[94m", "\033[96m", "\033[97m"
GRAY, BOLD, RESET = "\033[90m", "\033[1m", "\033[0m"

_DADOS_POR_UTILIZADOR = Path(
    os.environ.get(
        "LOCALAPPDATA",
        str(Path.home() / "AppData" / "Local"),
    )
) / "Faturas Shadow_77"
BASE_DIR = Path(os.environ.get("FATURAS_SHADOW_77_DIR", str(_DADOS_POR_UTILIZADOR)))
ARQUIVO_DIR = BASE_DIR / "arquivo"
DB_FILE = BASE_DIR / "Faturas Shadow_77.db"
LOGS_DIR = BASE_DIR / "logs"
REPETIDOS_DIR = BASE_DIR / "repetidos"
RELATORIOS_DIR = BASE_DIR / "relatorios"
EXCEL_PADRAO = BASE_DIR / "Faturas_2026.xlsx"

PASTAS_IGNORADAS = {"repetidos", "logs", "config", "relatorios", "media_originais"}
SUBPASTAS_JA_DE_LADO = {("arquivo", "duplicados"), ("arquivo", "removidos")}
EMISSORES_GENERICOS = {"", "desconhecido", "comprovativo (por associar)", "mb way / sibs", "cgd",
                       "millennium bcp", "novo banco"}

ESTADOS_PENDENTES = ("Por pagar", "Por Identificar", "Pago (a confirmar)")
ESTADOS_PARA_OK = ESTADOS_PENDENTES + ("Informativo",)
EMISSORES_SEM_PASTA = {"", "desconhecido", "comprovativo (por associar)"}
_PRINCIPAL = []


def _principal():
    """O módulo principal (mesmas regras de datas e de 2FA)."""
    if not _PRINCIPAL:
        try:
            sys.path.insert(0, str(Path(__file__).resolve().parent))
            import faturas_shadow_77 as principal
            _PRINCIPAL.append(principal)
        except Exception as e:
            _PRINCIPAL.append(None)
            print(f"{Y}⚠ Não consegui carregar o módulo principal ({e}); não vou separar por ano.{RESET}")
    return _PRINCIPAL[0]


EXCEL_PARA_BD = {
    "Categoria": "categoria", "Emissor": "emissor", "Tipo Doc": "tipo_doc", "Entidade MB": "entidade_mb",
    "Referência MB": "referencia_mb", "Valor": "valor", "Data limite": "data_limite", "Estado": "estado",
    "Origem": "origem", "Ficheiro": "ficheiro", "Notas": "notas", "Data doc": "data_doc",
}
COLUNAS_BD = ["id", "categoria", "emissor", "tipo_doc", "entidade_mb", "referencia_mb", "valor", "data_limite",
              "estado", "origem", "ficheiro", "notas", "data_registo", "ligado_a", "data_doc"]


# ---------------------------------------------------------------------------
# utilitários
# ---------------------------------------------------------------------------
def cabecalho():
    largura = os.get_terminal_size().columns if sys.stdout.isatty() else 80
    linhas = [("=" * 40, C), ("FATURAS SHADOW_77 v9.7", C + BOLD), ("COMPARADOR", W), ("=" * 40, C)]
    print()
    for texto, cor in linhas:
        print(" " * max(0, (largura - len(texto)) // 2) + f"{cor}{texto}{RESET}")
    print()


def progresso_linha(texto, i, n):
    if sys.stdout.isatty():
        sys.stdout.write(f"\r\033[2K  {GRAY}{texto} {i}/{n}{RESET}")
        sys.stdout.flush()
        if i >= n:
            sys.stdout.write("\r\033[2K")


def so_digitos(valor):
    return re.sub(r"\D", "", str(valor)) if valor not in (None, "") else ""


def cents(valor):
    try:
        return int(round(float(valor) * 100))
    except (TypeError, ValueError):
        return 0


def classe(tipo):
    t = (tipo or "").strip().lower()
    if t in ("fatura", "aviso de cobrança / condomínio", "nota de crédito"):
        return "A_PAGAR"
    if t in ("comprovativo", "transferencia", "transferência", "pagamento ordenado fr"):
        return "PAGAMENTO"
    if t == "fatura-recibo":
        return "FATURA_RECIBO"
    return "OUTRO"


def base_nome(nome):
    """Nome sem o prefixo de hora que o script junta quando o ficheiro já existia."""
    return re.sub(r"^(?:\d{12}_|\d{6}_)", "", Path(str(nome)).name).lower()


def prioridade_zona(rel):
    r = rel.lower()
    if r.startswith("arquivo/"):
        return 0
    if r.startswith("por identificar/"):
        return 1
    if r.startswith("whatsapp/"):
        return 2
    if r.startswith("entrada/"):
        return 3
    return 4


def zona(rel):
    p = rel.split("/")
    if p[0] == "arquivo" and len(p) > 1:
        return "arquivo/" + ("diversos" if p[1] == "diversos" else p[1])
    if p[0] == "entrada" and len(p) > 2:
        return "entrada/" + p[1]
    return p[0] if len(p) == 1 else p[0] + "/" + p[1]


def qualidade_linha(row):
    q = 0
    if (row.get("estado") or "") not in ("", "Por Identificar"):
        q += 30
    if (row.get("emissor") or "").strip().lower() not in EMISSORES_GENERICOS:
        q += 10
    if (row.get("valor") or 0) > 0:
        q += 10
    if len(so_digitos(row.get("referencia_mb"))) >= 6:
        q += 5
    return q


def chave_caminho(p):
    return os.path.normcase(str(Path(p)))


def caminho_do_registo(rel):
    """Mesma regra do módulo principal: 'por identificar/' vive em BASE_DIR."""
    rel = (rel or "").replace("\\", "/").strip()
    if not rel or rel.startswith("manual/"):
        return None
    if rel.lower().startswith("por identificar/"):
        return BASE_DIR / rel
    return ARQUIVO_DIR / rel


def _data_de_texto(txt):
    txt = str(txt or "")
    m = re.search(r"(20\d{2})[-/_.]?(\d{2})[-/_.]?(\d{2})", txt)
    if m:
        try:
            return datetime(int(m.group(1)), int(m.group(2)), int(m.group(3))).date()
        except ValueError:
            pass
    m = re.search(r"(\d{2})[-/.](\d{2})[-/.](20\d{2})", txt)
    if m:
        try:
            return datetime(int(m.group(3)), int(m.group(2)), int(m.group(1))).date()
        except ValueError:
            pass
    return None


def data_do_documento(f):
    """('limite', data) se houver data limite; senão ('ficheiro', data do nome); senão None."""
    row = f["row"] or {}
    d = _data_de_texto(row.get("data_limite"))
    if d:
        return ("limite", d)
    d = _data_de_texto(f["path"].name)
    if d:
        return ("ficheiro", d)
    return None


def comparar_datas(a, b):
    da, db = data_do_documento(a), data_do_documento(b)
    if not da or not db or da[0] != db[0]:
        return "desconhecido"
    if da[0] == "limite":
        return "igual" if da[1] == db[1] else "diferente"
    return "igual" if abs((da[1] - db[1]).days) <= 1 else "diferente"


# ---------------------------------------------------------------------------
# recolha de dados
# ---------------------------------------------------------------------------
def listar_pdfs():
    ficheiros, ja_de_lado = [], 0
    for dirpath, dirnames, filenames in os.walk(BASE_DIR):
        rel_dir = Path(dirpath).relative_to(BASE_DIR).parts
        dirnames[:] = [d for d in dirnames
                       if not (len(rel_dir) == 0 and d.lower() in PASTAS_IGNORADAS)]
        if len(rel_dir) >= 2 and (rel_dir[0].lower(), rel_dir[1].lower()) in SUBPASTAS_JA_DE_LADO:
            ja_de_lado += sum(1 for f in filenames if f.lower().endswith(".pdf"))
            dirnames[:] = []
            continue
        for nome in filenames:
            if nome.lower().endswith(".pdf"):
                p = Path(dirpath) / nome
                try:
                    ficheiros.append({"path": p, "rel": p.relative_to(BASE_DIR).as_posix(),
                                      "size": p.stat().st_size, "hash": None, "row": None, "fonte": "-"})
                except OSError:
                    pass
    return ficheiros, ja_de_lado


def carregar_bd():
    if not DB_FILE.exists():
        return {}, []
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    try:
        existentes = [r[1] for r in conn.execute("PRAGMA table_info(faturas)")]
        cols = [c for c in COLUNAS_BD if c in existentes]
        rows = [dict(r) for r in conn.execute(f"SELECT {', '.join(cols)} FROM faturas")]
    except sqlite3.Error:
        rows = []
    finally:
        conn.close()
    por_caminho = {}
    for r in rows:
        p = caminho_do_registo(r.get("ficheiro"))
        if p is not None:
            por_caminho[chave_caminho(p)] = r
    return por_caminho, rows


def carregar_excel(caminho):
    """Linhas do Excel corrigido, indexadas pelo caminho do ficheiro."""
    out = {}
    wb = openpyxl.load_workbook(caminho, read_only=True, data_only=True)
    for ws in wb.worksheets:
        linhas = ws.iter_rows(values_only=True)
        try:
            cab = [str(c).strip() if c is not None else "" for c in next(linhas)]
        except StopIteration:
            continue
        idx = {EXCEL_PARA_BD[c]: i for i, c in enumerate(cab) if c in EXCEL_PARA_BD}
        if "ficheiro" not in idx:
            continue
        for lin in linhas:
            rel = lin[idx["ficheiro"]] if idx["ficheiro"] < len(lin) else None
            if not rel:
                continue
            row = {col: (lin[i] if i < len(lin) else None) for col, i in idx.items()}
            p = caminho_do_registo(str(rel))
            if p is not None:
                out[chave_caminho(p)] = row
    wb.close()
    return out


def sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for bloco in iter(lambda: f.read(1024 * 1024), b""):
            h.update(bloco)
    return h.hexdigest()


# ---------------------------------------------------------------------------
# análise
# ---------------------------------------------------------------------------
def pontuacao(f):
    """Quanto maior, mais merece ser o ficheiro que fica."""
    r, p = f["row"], 0
    if r is not None:
        p += 40
    if not f["rel"].lower().startswith("por identificar"):
        p += 50
    if r:
        if (r.get("estado") or "") not in ("", "Por Identificar"):
            p += 30
        if (r.get("emissor") or "").strip().lower() not in EMISSORES_GENERICOS:
            p += 10
        if (r.get("valor") or 0) > 0:
            p += 10
        if len(so_digitos(r.get("referencia_mb"))) >= 6:
            p += 5
    if not re.match(r"^\d{12}_", f["path"].name):
        p += 5
    return p


def ordem_de_preferencia(f):
    r = f["row"] or {}
    return (-pontuacao(f), r.get("id") if r.get("id") is not None else 10 ** 9, len(f["rel"]), f["rel"].lower())


def _antigo(f):
    return bool(f.get("ano_info") and f["ano_info"]["antigo"])


def _avaliar_ano(principal, f):
    row = f["row"]
    nomes = f"{f['path'].name} {Path(str((f.get('db_row') or {}).get('ficheiro') or '')).name}"
    kw = dict(nome=nomes, data_limite=row.get("data_limite"), data_doc_bd=row.get("data_doc"))
    info = principal.avaliar_datas(**kw)
    if info["ano"] is None or info["antigo"] or len(info["data_doc"] or "") < 10:   # sem data completa: confirma no texto do PDF (sem OCR)
        texto = principal.texto_rapido_pdf(f["path"])
        if texto.strip():
            info = principal.avaliar_datas(texto=texto, **kw)
    return info


def _pasta_emissor(emissor):
    e = (emissor or "").strip()
    return "Sem emissor" if e.lower() in EMISSORES_SEM_PASTA else re.sub(r'[\\/*?:"<>|]', "", e)


def _rel_registo(path):
    """Caminho absoluto -> valor guardado na base de dados (relativo a arquivo/, ou 'por identificar/...')."""
    path = Path(path)
    try:
        return path.relative_to(ARQUIVO_DIR).as_posix()
    except ValueError:
        return path.relative_to(BASE_DIR).as_posix()


def _planear_movimento(principal, f):
    """O que fazer a um ficheiro que fica: mudar de pasta (ano/mês/emissor ou repor no arquivo) e/ou atualizar estado e data."""
    row, dbr, info = f["row"], f.get("db_row"), f.get("ano_info")
    if dbr is None:
        return None
    atual_rel = dbr.get("ficheiro") or ""
    nome = Path(atual_rel).name if f.get("relink") else f["path"].name
    em_quarentena = atual_rel.lower().startswith("por identificar/") and (row.get("estado") or "") == "Por Identificar"
    destino, motivo = None, None
    if info and info["data_doc"] and (not em_quarentena or info["antigo"]):
        pasta = principal.pasta_ano_mes(info["data_doc"])
        if (row.get("tipo_doc") or "") == "Diversos":
            destino = ARQUIVO_DIR / "diversos" / pasta / nome
        else:
            destino = ARQUIVO_DIR / pasta / _pasta_emissor(row.get("emissor")) / nome
        motivo = "Por ano"
    elif f.get("relink") and f.get("destino_registado") is not None:
        destino, motivo = f["destino_registado"], "Repor no arquivo"
    estado_atual = row.get("estado") or ""
    estado_novo = "Ok" if (info and info["antigo"] and estado_atual in ESTADOS_PARA_OK) else None
    data_nova = info["data_doc"] if (info and info["data_doc"] and info["data_doc"] != (dbr.get("data_doc") or "")) else None
    mover = destino is not None and chave_caminho(destino) != chave_caminho(f["path"])
    if not mover and estado_novo is None and data_nova is None:
        return None
    if not mover:
        destino, motivo = None, "Atualizar estado/data"
    return {"file": f, "destino": destino if mover else None, "estado_novo": estado_novo, "data_nova": data_nova,
            "motivo": motivo, "relink": bool(f.get("relink")), "estado_antes": estado_atual}


def _row_resumo(f):
    r = f["row"] or {}
    return {"emissor": r.get("emissor") or "", "tipo": r.get("tipo_doc") or "", "entidade": so_digitos(r.get("entidade_mb")),
            "ref": so_digitos(r.get("referencia_mb")), "valor": r.get("valor") if r.get("valor") is not None else "",
            "estado": r.get("estado") or ""}


def _fundir(dbrow, xrow):
    row = dict(dbrow or {})
    if xrow is not None:
        row.update(xrow)                      # o Excel corrigido manda
        if dbrow is not None:
            row["ficheiro"] = dbrow.get("ficheiro", xrow.get("ficheiro"))
    return row


def analisar(excel_caminho=None):
    ficheiros, ja_de_lado = listar_pdfs()
    bd_por_caminho, bd_rows = carregar_bd()
    excel = carregar_excel(excel_caminho) if excel_caminho else {}

    for f in ficheiros:
        k = chave_caminho(f["path"])
        dbrow, xrow = bd_por_caminho.get(k), excel.get(k)
        f["relink"], f["destino_registado"] = False, None
        if dbrow is not None or xrow is not None:
            f["row"] = _fundir(dbrow, xrow)
            f["fonte"] = "Excel" if xrow is not None else "BD"
            f["db_id"] = dbrow.get("id") if dbrow else None
            f["db_row"] = dbrow
        else:
            f["db_id"], f["db_row"] = None, None

    # conteúdo: só se calcula o hash de ficheiros com o mesmo tamanho
    por_tamanho = defaultdict(list)
    for f in ficheiros:
        por_tamanho[f["size"]].append(f)
    a_hashear = [f for lista in por_tamanho.values() if len(lista) > 1 for f in lista]
    for i, f in enumerate(a_hashear, 1):
        try:
            f["hash"] = sha256(f["path"])
        except OSError:
            f["hash"] = None
        progresso_linha("A comparar o conteúdo dos ficheiros", i, len(a_hashear))

    # --- re-ligar registos cujo ficheiro não está no caminho registado (pelo nome de base + mesmo conteúdo)
    existentes = {chave_caminho(f["path"]) for f in ficheiros}
    sem_row = defaultdict(list)
    for f in ficheiros:
        if f["row"] is None:
            sem_row[base_nome(f["path"].name)].append(f)
    orfaos_brutos = []
    for r in bd_rows:
        p = caminho_do_registo(r.get("ficheiro"))
        if p is not None and chave_caminho(p) not in existentes:
            orfaos_brutos.append((r, p))
    orfaos_brutos.sort(key=lambda t: -qualidade_linha(_fundir(t[0], excel.get(chave_caminho(t[1])))))  # melhor linha primeiro
    orfaos, ambiguos, linhas_repetidas, religados = [], [], [], 0
    for r, p in orfaos_brutos:
        cands = [f for f in sem_row.get(base_nome(r.get("ficheiro")), []) if f["row"] is None]
        hashes = {f["hash"] for f in cands}
        if cands and (len(cands) == 1 or (None not in hashes and len(hashes) == 1)):
            f = sorted(cands, key=lambda x: (prioridade_zona(x["rel"]), len(x["rel"]), x["rel"].lower()))[0]
            f["row"] = _fundir(r, excel.get(chave_caminho(p)))
            f["fonte"] = "Excel (re-ligado)" if chave_caminho(p) in excel else "BD (re-ligado)"
            f["db_id"], f["db_row"], f["relink"], f["destino_registado"] = r.get("id"), r, True, p
            religados += 1
        elif cands:
            ambiguos.append(r)
        else:
            orfaos.append((r, p))
    # linhas sem ficheiro cujo documento já tem outra linha (mesmo nome de base) = linhas repetidas
    por_base_com_row = defaultdict(list)
    for f in ficheiros:
        if f["row"] is not None:
            por_base_com_row[base_nome(f["path"].name)].append(f)
    orfaos_reais = []
    for r, p in orfaos:
        dono = por_base_com_row.get(base_nome(r.get("ficheiro")))
        if dono:
            linhas_repetidas.append({"row": _fundir(r, excel.get(chave_caminho(p))), "db_row": r, "mantido": dono[0]})
        else:
            orfaos_reais.append(r)

    principal = _principal()
    for f in ficheiros:
        f["ano_info"] = _avaliar_ano(principal, f) if (principal and f["row"] is not None) else None

    repetidos, grupos, conflitos = [], [], []
    marcados = set()

    # 1) idênticos (hash)
    por_hash = defaultdict(list)
    for f in a_hashear:
        if f["hash"]:
            por_hash[f["hash"]].append(f)
    n_grupo = 0
    for membros in por_hash.values():
        if len(membros) < 2:
            continue
        membros.sort(key=ordem_de_preferencia)
        n_grupo += 1
        manter, resto = membros[0], membros[1:]
        grupos.append({"n": n_grupo, "motivo": "Idêntico (mesmo conteúdo)", "confianca": "Alta",
                       "mantido": manter, "repetidos": resto})
        for d in resto:
            repetidos.append({"file": d, "motivo": "Idêntico (mesmo conteúdo)", "confianca": "Alta", "mantido": manter})
            marcados.add(chave_caminho(d["path"]))

    # 2) mesma classe + mesma referência (+ entidade/valor/emissor/data compatíveis)
    restantes = [f for f in ficheiros if chave_caminho(f["path"]) not in marcados]
    cand = defaultdict(list)
    for f in restantes:
        if not f["row"]:
            continue
        cl = classe(f["row"].get("tipo_doc"))
        ref = so_digitos(f["row"].get("referencia_mb"))
        if cl != "OUTRO" and len(ref) >= 6:
            cand[(cl, ref)].append(f)
    for (cl, ref), membros in cand.items():
        if len(membros) < 2:
            continue
        valores = {cents(m["row"].get("valor")) for m in membros if cents(m["row"].get("valor")) > 0}
        entidades = {so_digitos(m["row"].get("entidade_mb")) for m in membros if so_digitos(m["row"].get("entidade_mb"))}
        emissores = {(m["row"].get("emissor") or "").strip().lower() for m in membros} - EMISSORES_GENERICOS
        problema = None
        if len(valores) > 1:
            problema = "Mesma referência com valores diferentes: " + ", ".join(f"{v / 100:.2f}" for v in sorted(valores))
        elif len(entidades) > 1:
            problema = "Mesma referência com entidades diferentes: " + ", ".join(sorted(entidades))
        elif len(emissores) > 1:
            problema = "Mesma referência com emissores diferentes: " + ", ".join(sorted(emissores))
        if problema:
            if not all(_antigo(m) for m in membros):      # anos antigos não se verificam
                conflitos.append({"motivo": problema, "ref": ref, "membros": membros})
            continue
        membros.sort(key=ordem_de_preferencia)
        manter = membros[0]
        n_grupo += 1
        resto_ok = []
        for m in membros[1:]:
            c = comparar_datas(manter, m)
            if c == "diferente":
                if not (_antigo(manter) and _antigo(m)):
                    conflitos.append({"motivo": "Mesma referência e valor mas datas diferentes (pagamento recorrente?)",
                                      "ref": ref, "membros": [manter, m]})
                continue
            resto_ok.append((m, "Alta" if c == "igual" else "Média"))
        if resto_ok:
            grupos.append({"n": n_grupo, "motivo": "Mesma referência/entidade/valor", "confianca": None,
                           "mantido": manter, "repetidos": [m for m, _ in resto_ok]})
            for m, conf in resto_ok:
                repetidos.append({"file": m, "motivo": "Mesma referência/entidade/valor", "confianca": conf, "mantido": manter})
                marcados.add(chave_caminho(m["path"]))

    mantidos = [f for f in ficheiros if chave_caminho(f["path"]) not in marcados]

    # 3) ficheiros que ficam: repor no arquivo, separar por ano/mês/emissor, marcar "Ok" o que é anterior a ANO
    movimentos = []
    if principal:
        for f in mantidos:
            x = _planear_movimento(principal, f)
            if x:
                movimentos.append(x)
    elif True:
        for f in mantidos:                                # sem módulo principal: só repõe os re-ligados
            d = f.get("destino_registado")
            if f.get("relink") and d is not None and chave_caminho(d) != chave_caminho(f["path"]):
                movimentos.append({"file": f, "destino": d, "estado_novo": None, "data_nova": None,
                                   "motivo": "Repor no arquivo", "relink": True, "estado_antes": (f["row"] or {}).get("estado")})
    reposicoes = [x for x in movimentos if x["motivo"] == "Repor no arquivo" and x["destino"] is not None]

    # 4) pares fatura <-> pagamento (entre os que ficam)
    pares, por_ref, sem_ref = [], defaultdict(lambda: {"A_PAGAR": [], "PAG": []}), []
    for f in mantidos:
        if not f["row"] or _antigo(f):                    # só se verifica ANO_VERIFICACAO e seguintes
            continue
        cl = classe(f["row"].get("tipo_doc"))
        ref = so_digitos(f["row"].get("referencia_mb"))
        if cl == "OUTRO":
            continue
        if len(ref) < 6:
            sem_ref.append(f)
            continue
        por_ref[ref]["A_PAGAR" if cl == "A_PAGAR" else "PAG"].append(f)
    for ref, d in por_ref.items():
        for fat in d["A_PAGAR"]:
            vf = cents(fat["row"].get("valor"))
            pagos = [p for p in d["PAG"] if cents(p["row"].get("valor")) in (0, vf) or vf == 0]
            difere = [p for p in d["PAG"] if p not in pagos]
            if pagos:
                pares.append({"ref": ref, "fatura": fat, "pagamento": pagos[0], "estado": "Paga"})
            elif difere:
                pares.append({"ref": ref, "fatura": fat, "pagamento": difere[0], "estado": "Pagamento com valor diferente"})
            else:
                pares.append({"ref": ref, "fatura": fat, "pagamento": None, "estado": "Sem pagamento"})
        if not d["A_PAGAR"]:
            for p in d["PAG"]:
                tipo = "Fatura-recibo (já inclui o pagamento)" if classe(p["row"].get("tipo_doc")) == "FATURA_RECIBO" \
                    else "Pagamento sem fatura"
                pares.append({"ref": ref, "fatura": None, "pagamento": p, "estado": tipo})

    # 5) possíveis repetidos sem referência (para rever à mão)
    por_sem = defaultdict(list)
    for f in sem_ref:
        r = f["row"]
        por_sem[(classe(r.get("tipo_doc")), (r.get("emissor") or "").strip().lower(), cents(r.get("valor")))].append(f)
    rever = [m for (cl, em, v), m in por_sem.items() if len(m) > 1 and v > 0 and em not in EMISSORES_GENERICOS]

    sem_registo = [f for f in ficheiros if f["row"] is None]

    return {"ficheiros": ficheiros, "ja_de_lado": ja_de_lado, "repetidos": repetidos, "grupos": grupos,
            "conflitos": conflitos, "pares": pares, "rever": rever, "sem_registo": sem_registo, "orfaos": orfaos_reais,
            "ambiguos": ambiguos, "religados": religados, "linhas_repetidas": linhas_repetidas, "reposicoes": reposicoes,
            "movimentos": movimentos, "mantidos": mantidos, "excel": excel_caminho if excel else None, "n_bd": len(bd_rows)}


# ---------------------------------------------------------------------------
# relatório
# ---------------------------------------------------------------------------
def _folha(wb, titulo, cab, linhas):
    ws = wb.create_sheet(titulo)
    ws.append(cab)
    for c in range(1, len(cab) + 1):
        cel = ws.cell(row=1, column=c)
        cel.font = Font(bold=True)
        cel.fill = PatternFill(start_color="E0E0E0", end_color="E0E0E0", fill_type="solid")
        cel.alignment = Alignment(horizontal="center")
    for l in linhas:
        ws.append(l)
    for c in range(1, len(cab) + 1):
        maior = max((len(str(ws.cell(row=r, column=c).value or "")) for r in range(1, ws.max_row + 1)), default=10)
        ws.column_dimensions[get_column_letter(c)].width = min(max(maior + 3, 12), 80)
    ws.freeze_panes = "A2"
    return ws


def gerar_relatorio(res, run_id):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    rep = res["repetidos"]
    mb = sum(d["file"]["size"] for d in rep) / 1048576
    resumo = [
        ["PDFs analisados", len(res["ficheiros"])],
        ["Já de lado anteriormente (arquivo/duplicados e removidos)", res["ja_de_lado"]],
        ["Registos na base de dados", res["n_bd"]],
        ["Fonte dos dados corrigidos", str(res["excel"]) if res["excel"] else "(só base de dados)"],
        ["Repetidos a pôr de lado", len(rep)],
        ["   - idênticos (mesmo conteúdo)", sum(1 for d in rep if d["motivo"].startswith("Idêntico"))],
        ["   - mesma referência, confiança alta", sum(1 for d in rep if not d["motivo"].startswith("Idêntico") and d["confianca"] == "Alta")],
        ["   - mesma referência, confiança média", sum(1 for d in rep if d["confianca"] == "Média")],
        ["Espaço a libertar (MB)", round(mb, 1)],
        ["Registos re-ligados ao ficheiro pelo nome", res["religados"]],
        ["Linhas repetidas da base de dados (a pôr de lado)", len(res["linhas_repetidas"])],
        ["Ficheiros a repor no caminho registado do arquivo", len(res["reposicoes"])],
        ["Ficheiros a mover para arquivo/<ano>/<mês>/<emissor>", sum(1 for x in res["movimentos"] if x["motivo"] == "Por ano" and x["destino"])],
        ["Estados a passar a Ok (anteriores ao ano de verificação)", sum(1 for x in res["movimentos"] if x["estado_novo"])],
        ["Documentos sem data (ano desconhecido, ficam como estão)",
         sum(1 for f in res["mantidos"] if f["row"] is not None and f.get("ano_info") and not f["ano_info"]["ano"])],
        ["Documentos que ficam", len(res["mantidos"])],
        ["Conflitos (não tocados)", len(res["conflitos"])],
        ["Faturas pagas (com comprovativo)", sum(1 for p in res["pares"] if p["estado"] == "Paga"),],
        ["Faturas sem pagamento", sum(1 for p in res["pares"] if p["estado"] == "Sem pagamento")],
        ["Pagamentos sem fatura", sum(1 for p in res["pares"] if p["estado"] == "Pagamento sem fatura")],
        ["Ficheiros sem registo na base de dados/Excel", len(res["sem_registo"])],
        ["Registos sem ficheiro (sem candidato no disco)", len(res["orfaos"])],
        ["Registos com ficheiro ambíguo (conteúdos diferentes)", len(res["ambiguos"])],
        ["Execução", run_id],
    ]
    _folha(wb, "Resumo", ["Indicador", "Valor"], resumo)

    cab = ["Motivo", "Confiança", "Ficheiro a pôr de lado", "KB", "Emissor", "Tipo", "Entidade", "Referência", "Valor", "Estado",
           "Fica (ficheiro mantido)", "Fonte dos dados"]
    linhas = []
    for d in sorted(rep, key=lambda x: (x["confianca"] != "Alta", x["mantido"]["rel"], x["file"]["rel"])):
        s = _row_resumo(d["file"])
        linhas.append([d["motivo"], d["confianca"], d["file"]["rel"], round(d["file"]["size"] / 1024, 1), s["emissor"], s["tipo"],
                       s["entidade"], s["ref"], s["valor"], s["estado"], d["mantido"]["rel"], d["file"]["fonte"]])
    _folha(wb, "Repetidos", cab, linhas)

    linhas = []
    for g in res["grupos"]:
        for papel, lista in (("FICA", [g["mantido"]]), ("de lado", g["repetidos"])):
            for f in lista:
                s = _row_resumo(f)
                linhas.append([g["n"], papel, g["motivo"], f["rel"], s["emissor"], s["tipo"], s["entidade"], s["ref"], s["valor"], s["estado"]])
    _folha(wb, "Grupos", ["Grupo", "Papel", "Motivo", "Ficheiro", "Emissor", "Tipo", "Entidade", "Referência", "Valor", "Estado"], linhas)

    linhas = []
    for i, c in enumerate(res["conflitos"], 1):
        for f in c["membros"]:
            s = _row_resumo(f)
            linhas.append([i, c["motivo"], f["rel"], s["emissor"], s["tipo"], s["entidade"], s["ref"], s["valor"], s["estado"]])
    _folha(wb, "Conflitos", ["N.º", "Motivo", "Ficheiro", "Emissor", "Tipo", "Entidade", "Referência", "Valor", "Estado"], linhas)

    linhas = []
    for p in res["pares"]:
        fa, pa = p["fatura"], p["pagamento"]
        base = fa or pa
        s = _row_resumo(base)
        linhas.append([p["estado"], p["ref"], s["entidade"], s["emissor"], s["valor"], fa["rel"] if fa else "", pa["rel"] if pa else ""])
    _folha(wb, "Pares fatura-pagamento", ["Estado", "Referência", "Entidade", "Emissor", "Valor", "Fatura", "Pagamento"], linhas)

    linhas = []
    for i, grupo in enumerate(res["rever"], 1):
        for f in grupo:
            s = _row_resumo(f)
            linhas.append([i, f["rel"], s["emissor"], s["tipo"], s["valor"], s["estado"]])
    _folha(wb, "Rever (sem referencia)", ["Grupo", "Ficheiro", "Emissor", "Tipo", "Valor", "Estado"], linhas)

    linhas = []
    for x in res["movimentos"]:
        s = _row_resumo(x["file"])
        linhas.append([x["motivo"], x["file"]["rel"], x["destino"].relative_to(BASE_DIR).as_posix() if x["destino"] else "(fica)",
                       s["emissor"], s["tipo"], s["valor"], x["estado_antes"], x["estado_novo"] or "", x["data_nova"] or ""])
    _folha(wb, "Mover e atualizar", ["O quê", "Ficheiro atual", "Vai para", "Emissor", "Tipo", "Valor", "Estado antes",
                                     "Estado depois", "Data doc nova"], linhas)

    por_ano = defaultdict(lambda: defaultdict(int))
    desconhecidos = []
    for f in res["mantidos"]:
        if f["row"] is None or not f.get("ano_info"):
            continue
        ano = f["ano_info"]["ano"]
        x = next((m for m in res["movimentos"] if m["file"] is f and m["estado_novo"]), None)
        estado = x["estado_novo"] if x else (f["row"].get("estado") or "-")
        por_ano[ano if ano else "sem data"][estado] += 1
        if not ano:
            desconhecidos.append(f)
    estados_cols = ["Ok", "Pago", "Pago (a confirmar)", "Por pagar", "Por Identificar", "Informativo", "Ordenado"]
    def chave_ano(a):
        return (-1, 0) if a == "sem data" else (0, -a)
    _folha(wb, "Por ano", ["Ano", "Documentos"] + estados_cols + ["Outros"],
           [[a, sum(por_ano[a].values())] + [por_ano[a].get(e, 0) for e in estados_cols] +
            [sum(v for k, v in por_ano[a].items() if k not in estados_cols)] for a in sorted(por_ano, key=chave_ano, reverse=True)])
    _folha(wb, "Ano desconhecido", ["Ficheiro", "Emissor", "Tipo", "Valor", "Estado"],
           [[f["rel"], _row_resumo(f)["emissor"], _row_resumo(f)["tipo"], _row_resumo(f)["valor"], _row_resumo(f)["estado"]]
            for f in desconhecidos])

    _folha(wb, "Linhas repetidas (BD)", ["ID", "Ficheiro registado", "Estado", "Emissor", "Valor", "Fica (ficheiro com a linha boa)"],
           [[x["db_row"].get("id"), x["db_row"].get("ficheiro"), x["row"].get("estado"), x["row"].get("emissor"), x["row"].get("valor"),
             x["mantido"]["rel"]] for x in res["linhas_repetidas"]])
    _folha(wb, "Sem registo", ["Ficheiro", "KB"], [[f["rel"], round(f["size"] / 1024, 1)] for f in res["sem_registo"]])
    _folha(wb, "Registos sem ficheiro", ["ID", "Ficheiro (BD)", "Emissor", "Tipo", "Valor", "Estado", "Situação"],
           [[r.get("id"), r.get("ficheiro"), r.get("emissor"), r.get("tipo_doc"), r.get("valor"), r.get("estado"), "sem candidato no disco"]
            for r in res["orfaos"]] +
           [[r.get("id"), r.get("ficheiro"), r.get("emissor"), r.get("tipo_doc"), r.get("valor"), r.get("estado"),
             "vários ficheiros com o mesmo nome mas conteúdos diferentes"] for r in res["ambiguos"]])

    RELATORIOS_DIR.mkdir(parents=True, exist_ok=True)
    destino = RELATORIOS_DIR / f"comparador_{run_id}.xlsx"
    wb.save(destino)
    return destino


def mostrar_resumo(res):
    rep = res["repetidos"]
    ident = sum(1 for d in rep if d["motivo"].startswith("Idêntico"))
    alta = sum(1 for d in rep if not d["motivo"].startswith("Idêntico") and d["confianca"] == "Alta")
    media = sum(1 for d in rep if d["confianca"] == "Média")
    mb = sum(d["file"]["size"] for d in rep) / 1048576
    zonas = defaultdict(int)
    for d in rep:
        zonas[zona(d["file"]["rel"])] += 1
    print(f"{C}{BOLD}=== RESULTADO DA COMPARAÇÃO ==={RESET}")
    print(f"  PDFs analisados:        {W}{len(res['ficheiros'])}{RESET}   (já de lado antes: {res['ja_de_lado']})")
    print(f"  Dados usados:           {'Excel corrigido + base de dados' if res['excel'] else 'base de dados'}")
    print(f"  {G}Repetidos a pôr de lado: {len(rep)}{RESET}  ({mb:.1f} MB)")
    print(f"      idênticos (conteúdo):            {ident}")
    print(f"      mesma referência, confiança alta:  {alta}")
    print(f"      mesma referência, confiança média: {media}")
    print(f"      por pasta: " + " · ".join(f"{z}: {n}" for z, n in sorted(zonas.items(), key=lambda x: -x[1])))
    print(f"  {G}Registos re-ligados ao ficheiro:   {res['religados']}{RESET}   → a repor no arquivo: {len(res['reposicoes'])}")
    print(f"  Linhas repetidas da base de dados: {len(res['linhas_repetidas'])}")
    n_ano = sum(1 for x in res["movimentos"] if x["motivo"] == "Por ano" and x["destino"])
    n_ok = sum(1 for x in res["movimentos"] if x["estado_novo"])
    print(f"  {G}Separar por ano:{RESET} {n_ano} ficheiro(s) a mover para arquivo/<ano>/<mês>/<emissor> · "
          f"{n_ok} estado(s) a passar a Ok (anteriores a 2026, sem verificação)")
    anos = defaultdict(int)
    for f in res["mantidos"]:
        if f["row"] is not None and f.get("ano_info"):
            anos[f["ano_info"]["ano"] or "sem data"] += 1
    if anos:
        print("  Documentos por ano:     " + " · ".join(f"{a}: {n}" for a, n in sorted(anos.items(), key=lambda t: (t[0] == "sem data", str(t[0])), reverse=False)))
    print(f"  Documentos que ficam:   {len(res['mantidos'])}")
    print(f"  {Y}Conflitos (não tocados): {len(res['conflitos'])}{RESET}")
    pagas = sum(1 for p in res["pares"] if p["estado"] == "Paga")
    sem = sum(1 for p in res["pares"] if p["estado"] == "Sem pagamento")
    sf = sum(1 for p in res["pares"] if p["estado"] == "Pagamento sem fatura")
    print(f"  Pares:                  {pagas} fatura(s) paga(s) · {sem} sem pagamento · {sf} pagamento(s) sem fatura")
    print(f"  Sem registo:            {len(res['sem_registo'])} ficheiro(s) · registos sem ficheiro: {len(res['orfaos'])}"
          f" (+{len(res['ambiguos'])} ambíguos)")


# ---------------------------------------------------------------------------
# aplicar / desfazer
# ---------------------------------------------------------------------------
SQL_REPETIDAS = '''CREATE TABLE IF NOT EXISTS faturas_repetidas (
    orig_id INTEGER, categoria TEXT, emissor TEXT, tipo_doc TEXT, entidade_mb TEXT, referencia_mb TEXT,
    valor REAL, data_limite TEXT, estado TEXT, origem TEXT, ficheiro TEXT, notas TEXT, data_registo TEXT,
    ligado_a INTEGER, run_id TEXT, motivo TEXT, mantido_ficheiro TEXT, movido_para TEXT, data_acao TEXT)'''


def autenticar(descricao):
    """Usa a mesma 2FA do módulo principal e falha de forma segura."""
    try:
        sys.path.insert(0, str(Path(__file__).resolve().parent))
        import faturas_shadow_77 as principal
        return principal.verificar_dupla_autenticacao(descricao)
    except ImportError as e:
        print(f"{R}✗ Não consegui carregar a dupla autenticação ({e}).{RESET}")
        print("  Coloca o comparador na pasta scripts e instala as dependências do projeto.")
        return False


def _caminho_livre(destino):
    if not destino.exists():
        return destino
    return destino.with_name(f"{destino.stem}_{datetime.now().strftime('%H%M%S%f')}{destino.suffix}")


def _limpar_pastas_vazias(pasta):
    """Remove pastas vazias, mas nunca as de 1.º nível (arquivo, entrada, scripts...)."""
    pasta = Path(pasta)
    while pasta.parent != BASE_DIR and BASE_DIR in pasta.parents:
        try:
            pasta.rmdir()
        except OSError:
            break
        pasta = pasta.parent


def _destino_livre(destino, conn, rid):
    """Caminho que não existe no disco nem está registado noutra linha da base de dados."""
    destino = Path(destino)
    candidato, n = destino, 0
    while True:
        ocupado = candidato.exists() and chave_caminho(candidato) != chave_caminho(destino) or \
                  (candidato.exists() and n == 0)
        if not ocupado and conn is not None:
            r = conn.execute("SELECT 1 FROM faturas WHERE ficheiro = ? AND id != ? LIMIT 1",
                             (_rel_registo(candidato), rid)).fetchone()
            ocupado = r is not None
        if not ocupado:
            return candidato
        n += 1
        candidato = destino.with_name(f"{destino.stem}_{n}{destino.suffix}")


def _guardar_linha(conn, row, dbr, run_id, motivo, mantido_rel, movido_para, data, tem_ligado, mant_id):
    """Passa uma linha de faturas para faturas_repetidas (e mantém a ligação fatura<->comprovativo)."""
    conn.execute(
        "INSERT INTO faturas_repetidas (orig_id, categoria, emissor, tipo_doc, entidade_mb, referencia_mb, valor, "
        "data_limite, estado, origem, ficheiro, notas, data_registo, ligado_a, run_id, motivo, mantido_ficheiro, "
        "movido_para, data_acao) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
        (dbr.get("id"), row.get("categoria"), row.get("emissor"), row.get("tipo_doc"), row.get("entidade_mb"),
         row.get("referencia_mb"), row.get("valor"), row.get("data_limite"), row.get("estado"), row.get("origem"),
         dbr.get("ficheiro"), row.get("notas"), dbr.get("data_registo"), dbr.get("ligado_a"), run_id,
         motivo, mantido_rel, movido_para, data))
    if tem_ligado:
        lig = dbr.get("ligado_a")
        if lig and mant_id:   # passa a ligação fatura<->comprovativo para o documento que fica
            conn.execute("UPDATE faturas SET ligado_a = ? WHERE id = ? AND ligado_a IS NULL", (lig, mant_id))
            conn.execute("UPDATE faturas SET ligado_a = ? WHERE id = ?", (mant_id, lig))
        conn.execute("UPDATE faturas SET ligado_a = NULL WHERE ligado_a = ? AND id != ?", (dbr.get("id"), mant_id or -1))
    conn.execute("DELETE FROM faturas WHERE id = ?", (dbr.get("id"),))


def aplicar(res, incluir_medios):
    escolhidos = [d for d in res["repetidos"] if d["confianca"] == "Alta" or incluir_medios]
    linhas = res["linhas_repetidas"]
    movimentos = res["movimentos"]
    if not escolhidos and not linhas and not movimentos:
        print(f"{Y}Nada para fazer.{RESET}")
        return None
    run_id = datetime.now().strftime("%Y%m%d_%H%M%S")

    conn, tem_ligado = None, False
    if DB_FILE.exists():
        (LOGS_DIR / "backup_db").mkdir(parents=True, exist_ok=True)
        origem = sqlite3.connect(DB_FILE)
        destino_bk = sqlite3.connect(LOGS_DIR / "backup_db" / f"Faturas_Shadow_77_{run_id}.db")
        origem.backup(destino_bk)
        destino_bk.close()
        origem.close()
        conn = sqlite3.connect(DB_FILE, timeout=30)
        conn.execute(SQL_REPETIDAS)
        tem_ligado = "ligado_a" in [r[1] for r in conn.execute("PRAGMA table_info(faturas)")]

    pasta_run = REPETIDOS_DIR / run_id
    manifesto = {"run_id": run_id, "data": datetime.now().isoformat(timespec="seconds"), "itens": [], "desfeito": False}
    n_files = n_linhas = n_repor = 0
    total = len(escolhidos) + len(linhas) + len(movimentos)
    feito = 0
    try:
        # 1) ficheiros repetidos -> repetidos/<run>/
        for d in escolhidos:
            f = d["file"]
            dest = _caminho_livre(pasta_run / f["rel"])
            dest.parent.mkdir(parents=True, exist_ok=True)
            pai = f["path"].parent
            shutil.move(str(f["path"]), str(dest))
            _limpar_pastas_vazias(pai)
            item = {"tipo": "ficheiro", "origem": f["rel"], "destino": dest.relative_to(BASE_DIR).as_posix(),
                    "motivo": d["motivo"], "mantido": d["mantido"]["rel"], "db_id": f.get("db_id")}
            manifesto["itens"].append(item)
            n_files += 1
            if conn is not None and f.get("db_row") is not None:
                _guardar_linha(conn, f["row"], f["db_row"], run_id, d["motivo"], d["mantido"]["rel"], item["destino"],
                               manifesto["data"], tem_ligado, d["mantido"].get("db_id"))
            feito += 1
            progresso_linha("A aplicar", feito, total)

        # 2) linhas repetidas da base de dados (sem ficheiro)
        for x in linhas:
            dbr = x["db_row"]
            item = {"tipo": "linha", "origem": dbr.get("ficheiro"), "destino": f"(linha {dbr.get('id')})",
                    "motivo": "Linha repetida (mesmo documento registado 2x)", "mantido": x["mantido"]["rel"], "db_id": dbr.get("id")}
            manifesto["itens"].append(item)
            if conn is not None:
                _guardar_linha(conn, x["row"], dbr, run_id, item["motivo"], x["mantido"]["rel"], item["destino"],
                               manifesto["data"], tem_ligado, x["mantido"].get("db_id"))
            n_linhas += 1
            feito += 1
            progresso_linha("A aplicar", feito, total)

        # 3) mover (repor no arquivo / separar por ano) e atualizar estado e data na base de dados
        for x in movimentos:
            f, destino = x["file"], x["destino"]
            dbr = f["db_row"]
            feito += 1
            progresso_linha("A aplicar", feito, total)
            final = None
            if destino is not None:
                if not f["path"].exists():
                    print(f"{Y}  ⚠ ficheiro em falta: {f['rel']}{RESET}")
                    continue
                final = _destino_livre(destino, conn, dbr.get("id"))
                final.parent.mkdir(parents=True, exist_ok=True)
                pai = f["path"].parent
                shutil.move(str(f["path"]), str(final))
                _limpar_pastas_vazias(pai)
            ficheiro_novo = _rel_registo(final) if final else dbr.get("ficheiro")
            sets, vals = [], []
            if ficheiro_novo != dbr.get("ficheiro"):
                sets.append("ficheiro = ?"); vals.append(ficheiro_novo)
            if x["estado_novo"]:
                sets.append("estado = ?"); vals.append(x["estado_novo"])
            if x["data_nova"]:
                sets.append("data_doc = ?"); vals.append(x["data_nova"])
            if conn is not None and sets:
                conn.execute(f"UPDATE faturas SET {', '.join(sets)} WHERE id = ?", vals + [dbr.get("id")])
            manifesto["itens"].append({
                "tipo": "mover", "origem": f["rel"], "destino": final.relative_to(BASE_DIR).as_posix() if final else None,
                "db_id": dbr.get("id"), "ficheiro_antigo": dbr.get("ficheiro"), "ficheiro_novo": ficheiro_novo,
                "estado_antigo": dbr.get("estado"), "estado_novo": x["estado_novo"],
                "data_doc_antigo": dbr.get("data_doc"), "data_doc_novo": x["data_nova"]})
            if final:
                n_repor += 1
        if conn is not None:
            conn.commit()
    finally:
        if conn is not None:
            conn.close()
        pasta_run.mkdir(parents=True, exist_ok=True)
        (pasta_run / "manifesto.json").write_text(json.dumps(manifesto, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"{G}✓ {n_files} ficheiro(s) posto(s) de lado em {pasta_run}{RESET}")
    n_estados = sum(1 for x in movimentos if x["estado_novo"])
    print(f"{G}✓ {n_linhas} linha(s) repetida(s) da base de dados · {n_repor} ficheiro(s) movido(s) (arquivo/ano/mês/emissor) · "
          f"{n_estados} estado(s) passado(s) a Ok{RESET}")
    print(f"  Cópia da base de dados: {LOGS_DIR / 'backup_db' / ('Faturas_Shadow_77_' + run_id + '.db')}")
    print(f"  Para repor tudo: python comparador_faturas_shadow_77.py --desfazer")
    return run_id


def _repor_linha(conn, run_id, movido_para):
    row = conn.execute("SELECT * FROM faturas_repetidas WHERE run_id = ? AND movido_para = ?", (run_id, movido_para)).fetchone()
    if row is None:
        return True
    cols = [c[0] for c in conn.execute("SELECT * FROM faturas_repetidas LIMIT 0").description]
    r = dict(zip(cols, row))
    existentes = [c[1] for c in conn.execute("PRAGMA table_info(faturas)")]
    campos = [c for c in ["categoria", "emissor", "tipo_doc", "entidade_mb", "referencia_mb", "valor", "data_limite",
                          "estado", "origem", "ficheiro", "notas", "data_registo", "ligado_a"] if c in existentes]
    livre = conn.execute("SELECT 1 FROM faturas WHERE id = ?", (r["orig_id"],)).fetchone() is None
    colunas = (["id"] if livre else []) + campos
    valores = ([r["orig_id"]] if livre else []) + [r[c] for c in campos]
    try:
        conn.execute(f"INSERT INTO faturas ({', '.join(colunas)}) VALUES ({', '.join('?' * len(colunas))})", valores)
    except sqlite3.IntegrityError:
        return False
    conn.execute("DELETE FROM faturas_repetidas WHERE run_id = ? AND movido_para = ?", (run_id, movido_para))
    return True


def desfazer():
    manifestos = sorted(REPETIDOS_DIR.glob("*/manifesto.json")) if REPETIDOS_DIR.exists() else []
    manifestos = [m for m in manifestos if not json.loads(m.read_text(encoding="utf-8")).get("desfeito")]
    if not manifestos:
        print(f"{Y}Não há aplicações para desfazer.{RESET}")
        return
    caminho = manifestos[-1]
    man = json.loads(caminho.read_text(encoding="utf-8"))
    print(f"A repor a aplicação {man['run_id']} ({len(man['itens'])} item(ns))...")
    repostos = falhas = 0
    conn = sqlite3.connect(DB_FILE, timeout=30) if DB_FILE.exists() else None
    try:
        for it in reversed(man["itens"]):          # ordem inversa: primeiro as reposições
            tipo = it.get("tipo", "ficheiro")
            if tipo == "linha":
                if conn is not None and _repor_linha(conn, man["run_id"], it["destino"]):
                    repostos += 1
                else:
                    falhas += 1
                continue
            if tipo == "mover":
                if it.get("destino"):
                    origem, dest = BASE_DIR / it["origem"], BASE_DIR / it["destino"]
                    if not dest.exists() or origem.exists():
                        print(f"{Y}  ⚠ não foi possível repor {it['origem']}{RESET}")
                        falhas += 1
                        continue
                    origem.parent.mkdir(parents=True, exist_ok=True)
                    pai = dest.parent
                    shutil.move(str(dest), str(origem))
                    _limpar_pastas_vazias(pai)
                if conn is not None and it.get("db_id"):
                    sets, vals = ["ficheiro = ?"], [it["ficheiro_antigo"]]
                    if it.get("estado_novo"):
                        sets.append("estado = ?"); vals.append(it["estado_antigo"])
                    if it.get("data_doc_novo"):
                        sets.append("data_doc = ?"); vals.append(it["data_doc_antigo"])
                    conn.execute(f"UPDATE faturas SET {', '.join(sets)} WHERE id = ?", vals + [it["db_id"]])
                repostos += 1
                continue
            origem, dest = BASE_DIR / it["origem"], BASE_DIR / it["destino"]
            if not dest.exists():
                falhas += 1
                continue
            if origem.exists():
                print(f"{Y}  ⚠ já existe um ficheiro em {it['origem']}; mantido onde está.{RESET}")
                falhas += 1
                continue
            origem.parent.mkdir(parents=True, exist_ok=True)
            pai = dest.parent
            shutil.move(str(dest), str(origem))
            _limpar_pastas_vazias(pai)
            repostos += 1
            if tipo == "ficheiro" and conn is not None and not _repor_linha(conn, man["run_id"], it["destino"]):
                print(f"{Y}  ⚠ o registo de {it['origem']} já existe na base de dados.{RESET}")
        if conn is not None:
            conn.commit()
    finally:
        if conn is not None:
            conn.close()
    if falhas == 0:
        man["desfeito"] = True
        caminho.write_text(json.dumps(man, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"{G}✓ {repostos} item(ns) reposto(s).{RESET}" + (f" {Y}{falhas} com aviso.{RESET}" if falhas else ""))


# ---------------------------------------------------------------------------
# menu / CLI
# ---------------------------------------------------------------------------
def escolher_excel(args):
    if args.sem_excel:
        return None
    if args.excel:
        p = Path(args.excel)
        if not p.exists():
            print(f"{R}Excel não encontrado: {p}{RESET}")
            sys.exit(1)
        return p
    return EXCEL_PADRAO if EXCEL_PADRAO.exists() else None


def correr(modo, args):
    if not BASE_DIR.exists():
        print(f"{R}A pasta {BASE_DIR} não existe (define FATURAS_SHADOW_77_DIR se for outra).{RESET}")
        return
    excel = escolher_excel(args)
    if excel:
        print(f"{GRAY}Dados corrigidos: {excel}{RESET}")
    t0 = time.time()
    res = analisar(excel)
    run_id = datetime.now().strftime("%Y%m%d_%H%M%S")
    relatorio = gerar_relatorio(res, run_id)
    mostrar_resumo(res)
    print(f"\n  Relatório completo: {W}{relatorio}{RESET}  ({time.time() - t0:.0f}s)\n")
    if modo != "aplicar":
        return

    media = sum(1 for d in res["repetidos"] if d["confianca"] == "Média")
    incluir = args.incluir_medios
    if media and not incluir and sys.stdin.isatty():
        resp = input(f"Há {media} possível(is) repetido(s) de confiança MÉDIA (sem data para confirmar). Incluir? (s/N): ")
        incluir = resp.strip().lower() in ("s", "sim")
    n = sum(1 for d in res["repetidos"] if d["confianca"] == "Alta" or incluir)
    nl, nr = len(res["linhas_repetidas"]), len(res["movimentos"])
    if n == 0 and nl == 0 and nr == 0:
        print(f"{Y}Nada para fazer.{RESET}")
        return
    print(f"\nVai ser feito: {W}{n}{RESET} ficheiro(s) posto(s) de lado · {W}{nl}{RESET} linha(s) repetida(s) da base de dados · "
          f"{W}{nr}{RESET} ficheiro(s)/registo(s) a mover ou atualizar (por ano, repor no arquivo, estado Ok).")
    print("(nada é apagado; dá para desfazer)")
    if input("Escreve APLICAR para continuar: ").strip() != "APLICAR":
        print("Cancelado.")
        return
    if not autenticar(f"Comparador: {n} repetido(s), {nl} linha(s), {nr} movimento(s)/atualização(ões)"):
        return
    aplicar(res, incluir)


def menu(args):
    while True:
        print(f"\n{C}{BOLD}=== COMPARADOR ==={RESET}")
        print("  1) Analisar (só relatório, não mexe em nada)")
        print("  2) Analisar e pôr de lado os repetidos")
        print("  3) Desfazer a última aplicação")
        print("  0) Sair")
        op = input("Opção: ").strip()
        if op == "1":
            correr("analisar", args)
        elif op == "2":
            correr("aplicar", args)
        elif op == "3":
            desfazer()
        elif op == "0":
            return


def main():
    ap = argparse.ArgumentParser(description="Comparador de faturas repetidas")
    ap.add_argument("--analisar", action="store_true")
    ap.add_argument("--aplicar", action="store_true")
    ap.add_argument("--desfazer", action="store_true")
    ap.add_argument("--excel")
    ap.add_argument("--sem-excel", action="store_true")
    ap.add_argument("--incluir-medios", action="store_true")
    args = ap.parse_args()
    cabecalho()
    try:
        if args.desfazer:
            desfazer()
        elif args.aplicar:
            correr("aplicar", args)
        elif args.analisar:
            correr("analisar", args)
        else:
            menu(args)
    except KeyboardInterrupt:
        print(f"\n{Y}Interrompido.{RESET}")


if __name__ == "__main__":
    main()
