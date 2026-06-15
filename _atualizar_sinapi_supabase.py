from __future__ import annotations

import argparse
import os
import tempfile
import zipfile
from datetime import date
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen

from _build_sinapi_import import main as build_sinapi_csvs
from _import_sinapi_supabase import main as import_sinapi_supabase


BASE_URL = "https://www.caixa.gov.br/Downloads/sinapi-relatorios-mensais"
FONTES_DIR = Path("Gestão de obras") / "SINAPI_fontes"
IMPORT_ROOT = Path("Gestão de obras") / "SINAPI_importacao_previa"


def previous_months(start: date, limit: int = 12):
    year = start.year
    month = start.month
    for _ in range(limit):
        month -= 1
        if month == 0:
            month = 12
            year -= 1
        yield year, month


def candidate_urls(today: date):
    suffixes = ("_Retificacao03", "_Retificacao02", "_Retificacao01", "")
    for year, month in previous_months(today):
        for suffix in suffixes:
            name = f"SINAPI-{year}-{month:02d}-formato-xlsx{suffix}.zip"
            yield year, month, name, f"{BASE_URL}/{name}"


def download_zip(url: str, destination: Path) -> bool:
    req = Request(url, headers={"User-Agent": "Mozilla/5.0"})
    try:
        with urlopen(req, timeout=120) as resp:
            data = resp.read()
    except (HTTPError, URLError, TimeoutError):
        return False
    if not data.startswith(b"PK"):
        return False
    destination.parent.mkdir(parents=True, exist_ok=True)
    destination.write_bytes(data)
    return True


def latest_sinapi_zip(today: date) -> tuple[str, Path]:
    for year, month, name, url in candidate_urls(today):
        destination = FONTES_DIR / name
        if destination.exists() and destination.stat().st_size > 0:
            return f"{year}-{month:02d}", destination
        print(f"Testando {url}")
        if download_zip(url, destination):
            return f"{year}-{month:02d}", destination
    raise RuntimeError("Nao foi encontrado pacote SINAPI XLSX nos ultimos 12 meses.")


def extract_reference_xlsx(zip_path: Path, competencia: str) -> Path:
    extract_dir = FONTES_DIR / zip_path.stem
    extract_dir.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(zip_path) as zf:
        zf.extractall(extract_dir)
    matches = sorted(extract_dir.glob("SINAPI_Refer*.xlsx"))
    if not matches:
        matches = sorted(extract_dir.glob("*Refer*.xlsx"))
    if not matches:
        raise RuntimeError(f"Nao encontrei planilha SINAPI_Referencia no pacote {zip_path}")
    return matches[0]


def main() -> int:
    parser = argparse.ArgumentParser(description="Baixa, processa e importa o SINAPI mais recente no Supabase.")
    parser.add_argument("--no-import", action="store_true", help="Gera CSVs e auditoria sem enviar ao Supabase.")
    parser.add_argument("--today", default=None, help="Data de referencia YYYY-MM-DD, usada apenas para testes.")
    args = parser.parse_args()

    today = date.fromisoformat(args.today) if args.today else date.today()
    competencia, zip_path = latest_sinapi_zip(today)
    print(f"Pacote selecionado: {zip_path.name} ({competencia})")

    xlsx_path = extract_reference_xlsx(zip_path, competencia)
    out_dir = IMPORT_ROOT / competencia
    build_sinapi_csvs(xlsx_path=xlsx_path, out_dir=out_dir, ufs=("MG", "RJ", "SP"))

    if args.no_import:
        return 0

    os.environ["SINAPI_IMPORT_DIR"] = str(out_dir)
    return import_sinapi_supabase()


if __name__ == "__main__":
    raise SystemExit(main())
