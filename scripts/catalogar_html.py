import csv
import html
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
OUTPUT_DIR = BASE_DIR / "catalogo"
OUTPUT_HTML = OUTPUT_DIR / "index.html"
OUTPUT_CSV = OUTPUT_DIR / "catalogo.csv"
PAYS_DIR = BASE_DIR / "arquivo" / "removidos" / "paye_frances"

IGNORAR = {"logs", "config", "scripts", "dist", "Instalador_Output", "relatorios", "catalogo"}
MESES = {"01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12"}

PAYE_TERMS = [
    "recibo vencimentos",
    "recibo de vencimento",
    "bulletin de paie",
    "bulletin de salaire",
    "fiche de paie",
    "fiche de paye",
    "paye",
    "salaire",
    "net a payer",
    "net à payer",
]

PT_TERMS = [
    "portugal", "fatura", "recibo", "entidade", "referencia", "contribuinte",
    "n.º", "nº", "atcud", "iva", "millennium", "vodafone", "nos",
]

FR_TERMS = [
    "bulletin de paie", "bulletin de salaire", "salaire", "revenus",
    "cotisations", "cotisation", "paye", "france", "siret", "urssaf",
]


def pdf_text(path):
    # O catálogo usa nomes e caminhos para evitar varrer milhares de PDFs durante a criação.
    # Payslips claros podem ser confirmados separadamente pelo manifesto.
    return ""


def classificar(path):
    name = path.name.lower()
    text = pdf_text(path)
    haystack = f"{name}\n{text}"

    is_paye = any(term in haystack for term in PAYE_TERMS)
    fr = any(term in haystack for term in FR_TERMS)
    pt = any(term in haystack for term in PT_TERMS)

    if is_paye:
        country = "Paye / France"
    elif fr and not pt:
        country = "FR"
    elif pt and not fr:
        country = "PT"
    elif fr:
        country = "FR"
    elif pt:
        country = "PT"
    else:
        country = "Indefinido"

    return country, is_paye


def periodo(path):
    rel = path.relative_to(BASE_DIR).as_posix()
    parts = rel.split("/")
    year = month = ""
    for i, part in enumerate(parts):
        if len(part) == 4 and part.isdigit() and 1990 <= int(part) <= 2100:
            year = part
            if i + 1 < len(parts) and parts[i + 1] in MESES:
                month = parts[i + 1]
            break
    return year, month


def destino_livre(path):
    if not path.exists():
        return path
    i = 2
    while True:
        candidate = path.with_name(f"{path.stem}_{i}{path.suffix}")
        if not candidate.exists():
            return candidate


def recolher():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    rows = []
    paye_files = []
    for dirpath, dirnames, filenames in os.walk(BASE_DIR):
        root = Path(dirpath)
        rel_root = root.relative_to(BASE_DIR)
        if len(rel_root.parts) == 1 and rel_root.parts[0] in IGNORAR:
            dirnames[:] = []
            continue
        dirnames[:] = [d for d in dirnames if d not in IGNORAR]
        for filename in filenames:
            path = root / filename
            if filename.startswith("~$"):
                continue
            rel = path.relative_to(BASE_DIR).as_posix()
            country, is_paye = classificar(path)
            year, month = periodo(path)
            historical = bool(year and (int(year) < 2026 or (int(year) == 2026 and int(month or "13") <= 8)))
            category = "Paye / vencimento" if is_paye else country
            rows.append({
                "ficheiro": filename,
                "relativo": rel,
                "pais": country,
                "categoria": category,
                "ano": year,
                "mes": month,
                "historico_ate_agosto_2026": "Sim" if historical else "Nao",
                "tipo": path.suffix.lower().replace(".", "") or "ficheiro",
                "tamanho": path.stat().st_size if path.exists() else 0,
                "modificado": datetime.fromtimestamp(path.stat().st_mtime).isoformat(timespec="seconds") if path.exists() else "",
            })
            if is_paye:
                paye_files.append(path)
    rows.sort(key=lambda r: (r["ano"], r["mes"], r["pais"], r["relativo"]))
    return rows, paye_files


def escrever_csv(rows):
    with OUTPUT_CSV.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=list(rows[0].keys()), delimiter=";")
        writer.writeheader()
        writer.writerows(rows)


def mover_paye(paye_files):
    PAYS_DIR.mkdir(parents=True, exist_ok=True)
    manifest = {
        "data": datetime.now().isoformat(timespec="seconds"),
        "itens": [],
    }
    db_backup = BASE_DIR / "logs" / "backup_pre_paye"
    db_backup.mkdir(parents=True, exist_ok=True)
    shutil.copy2(DB_FILE, db_backup / f"Faturas_Shadow_77_{datetime.now():%Y%m%d_%H%M%S}.db")

    conn = sqlite3.connect(DB_FILE)
    for source in paye_files:
        try:
            rel = source.relative_to(BASE_DIR).as_posix()
        except ValueError:
            continue
        if not source.exists():
            continue
        destination = destino_livre(PAYS_DIR / source.name)
        shutil.move(str(source), str(destination))
        conn.execute(
            "DELETE FROM faturas WHERE replace(ficheiro, '\\\\', '/') = ? "
            "OR replace(ficheiro, '\\\\', '/') like ?",
            (rel, f"%/{source.name}"),
        )
        manifest["itens"].append({
            "origem": rel,
            "destino": destination.relative_to(BASE_DIR).as_posix(),
            "acao": "movido para quarentena paye",
        })
    conn.commit()
    conn.close()

    manifest_file = PAYS_DIR / "manifesto.json"
    manifest_file.write_text(
        __import__("json").dumps(manifest, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return manifest


def escrever_html(rows):
    pt_count = sum(r["pais"] == "PT" for r in rows)
    fr_count = sum(r["pais"] == "FR" for r in rows)
    paye_count = sum(r["pais"] == "Paye / France" for r in rows)
    hist_count = sum(r["historico_ate_agosto_2026"] == "Sim" for r in rows)

    data_json = __import__("json").dumps(rows, ensure_ascii=False)
    head = f"""<!doctype html>
<html lang="pt"><head><meta charset="utf-8">
<title>Catálogo Faturas Shadow_77</title>
<style>
:body {{ font-family: Segoe UI, Arial, sans-serif; margin: 24px; background:#f5f7fa; color:#1f2937 }}
.cards {{ display:flex; gap:12px; flex-wrap:wrap; margin-bottom:16px }}
.card {{ background:white; border:1px solid #dbe1e8; border-radius:10px; padding:12px 16px; min-width:150px }}
input,select {{ padding:8px; margin:4px 6px 12px 0; border:1px solid #bcc6d1; border-radius:6px }}
table {{ width:100%; border-collapse:collapse; background:white; font-size:13px }}
th {{ position:sticky; top:0; background:#24384a; color:white; text-align:left; padding:8px }}
td {{ border-bottom:1px solid #e6ebf0; padding:7px; vertical-align:top }}
tr:hover {{ background:#eef6ff }}
a {{ color:#075ea8 }}
.PT {{ color:#116329; font-weight:bold }} .FR {{ color:#6b21a8; font-weight:bold }}
.Paye {{ color:#a70000; font-weight:bold }}
</style></head><body>
<h1>Catálogo de documentos — Faturas Shadow_77</h1>
<div class="cards">
<div class="card"><b>{len(rows)}</b><br>ficheiros</div>
<div class="card"><b>{pt_count}</b><br>Portugal</div>
<div class="card"><b>{fr_count}</b><br>França</div>
<div class="card"><b>{paye_count}</b><br>Paye / vencimentos</div>
<div class="card"><b>{hist_count}</b><br>até agosto de 2026</div>
</div>
<input id="search" size="45" placeholder="Pesquisar nome, pasta, país...">
<select id="country"><option value="">Todos os países</option><option>PT</option><option>FR</option><option>Paye / France</option><option>Indefinido</option></select>
<select id="period"><option value="">Todos os períodos</option><option value="Sim">Até agosto de 2026</option><option value="Nao">Depois de agosto de 2026</option></select>
<div id="count"></div>
<table><thead><tr><th>País</th><th>Ano/mês</th><th>Categoria</th><th>Ficheiro</th><th>Tamanho</th><th>Caminho</th><th>Modificado</th></tr></thead>
<tbody id="body"></tbody></table>
<script>
const rows={data_json};
function fmt(bytes) {{
  if (!bytes) return "0";
  const units=["B","KB","MB","GB"]; let i=0, n=bytes;
  while(n>=1024&&i<units.length-1){{n/=1024;i++}}
  return n.toFixed(i ? 1 : 0)+" "+units[i];
}}
function render() {{
  const q=document.getElementById("search").value.toLowerCase();
  const country=document.getElementById("country").value;
  const period=document.getElementById("period").value;
  const filtered=rows.filter(r =>
    (!q || (r.ficheiro+" "+r.relativo+" "+r.pais+" "+r.categoria).toLowerCase().includes(q)) &&
    (!country || r.pais===country) &&
    (!period || r.historico_ate_agosto_2026===period)
  );
  document.getElementById("count").textContent=filtered.length+" resultado(s)";
  document.getElementById("body").innerHTML=filtered.map(r =>
    `<tr><td class="${{r.pais.startsWith("Paye")?"Paye":r.pais}}">${{r.pais}}</td><td>${{r.ano||"?"}}/${{r.mes||"?"}}</td><td>${{r.categoria}}</td><td>${{r.ficheiro.replaceAll("&","&amp;").replaceAll("<","&lt;")}}</td><td>${{fmt(r.tamanho)}}</td><td><a href="${{encodeURI(r.relativo)}}">${{r.relativo}}</a></td><td>${{r.modificado}}</td></tr>`
  ).join("");
}}
document.getElementById("search").addEventListener("input",render);
document.getElementById("country").addEventListener("change",render);
document.getElementById("period").addEventListener("change",render);
render();
</script></body></html>
""".replace("{data_json}", data_json).replace(":body", "body")
    OUTPUT_HTML.write_text(head, encoding="utf-8")


def reconstruir_excel():
    sys.path.insert(0, str(BASE_DIR / "scripts"))
    import faturas_shadow_77 as main
    main.DB_FILE = DB_FILE
    main.EXCEL_FILE = EXCEL_FILE
    main.inicializar_base_dados()
    main.sincronizar_bd_para_excel()


if __name__ == "__main__":
    rows, paye_files = recolher()
    escrever_csv(rows)
    escrever_html(rows)
    print("Catalogo:", OUTPUT_HTML)
    print("CSV:", OUTPUT_CSV)
    print("Ficheiros:", len(rows))
    print("Paye encontrados:", len(paye_files))
    if "--mover-paye" in sys.argv:
        manifest = mover_paye(paye_files)
        reconstruir_excel()
        print("Paye movidos:", len(manifest["itens"]))
