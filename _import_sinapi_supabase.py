from __future__ import annotations

import csv
import json
import os
import sys
from pathlib import Path
from urllib.error import HTTPError
from urllib.parse import urlencode
from urllib.request import Request, urlopen


DEFAULT_SUPABASE_URL = "https://bkeecvmfammwukdvvcjz.supabase.co"
DEFAULT_IMPORT_DIR = Path("Gestão de obras") / "SINAPI_importacao_previa" / "2026-05"
DEFAULT_BATCH_SIZE = 500


def env_required(name: str) -> str:
    value = os.environ.get(name, "").strip()
    if not value:
        raise RuntimeError(f"Variavel de ambiente obrigatoria ausente: {name}")
    return value


def read_csv(path: Path) -> list[dict]:
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f))


def clean_decimal(value: str):
    value = (value or "").strip()
    return None if value == "" else value


def clean_int(value: str):
    value = (value or "").strip()
    return None if value == "" else int(float(value))


class SupabaseRest:
    def __init__(self, url: str, key: str):
        self.url = url.rstrip("/")
        self.key = key

    def request(self, method: str, path: str, payload=None, params=None, prefer=None):
        query = f"?{urlencode(params or {}, doseq=True)}" if params else ""
        data = None
        headers = {
            "apikey": self.key,
            "Authorization": f"Bearer {self.key}",
            "Content-Type": "application/json",
        }
        if prefer:
            headers["Prefer"] = prefer
        if payload is not None:
            data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        req = Request(f"{self.url}/rest/v1/{path}{query}", data=data, headers=headers, method=method)
        try:
            with urlopen(req, timeout=120) as resp:
                body = resp.read().decode("utf-8")
                return json.loads(body) if body else None
        except HTTPError as exc:
            body = exc.read().decode("utf-8", errors="replace")
            raise RuntimeError(f"Erro Supabase {method} {path}: HTTP {exc.code} {body}") from exc

    def upsert(self, table: str, rows: list[dict], on_conflict: str = "id", batch_size: int = DEFAULT_BATCH_SIZE):
        total = len(rows)
        for start in range(0, total, batch_size):
            batch = rows[start : start + batch_size]
            self.request(
                "POST",
                table,
                batch,
                params={"on_conflict": on_conflict},
                prefer="resolution=merge-duplicates,return=minimal",
            )
            print(f"{table}: {min(start + batch_size, total)}/{total}")

    def select_one(self, table: str, params: dict):
        data = self.request("GET", table, params=params)
        if not data:
            return None
        return data[0]


def ensure_fonte(client: SupabaseRest) -> str:
    client.upsert(
        "orc_fontes",
        [{"nome": "SINAPI", "tipo": "SINAPI", "ativo": True}],
        on_conflict="nome",
        batch_size=1,
    )
    fonte = client.select_one("orc_fontes", {"select": "id", "nome": "eq.SINAPI"})
    if not fonte:
        raise RuntimeError("Nao foi possivel obter a fonte SINAPI no Supabase.")
    return fonte["id"]


def load_payloads(import_dir: Path, fonte_id: str):
    services = []
    for row in read_csv(import_dir / "sinapi_servicos.csv"):
        services.append(
            {
                "id": row["id"],
                "codigo_interno": row["codigo_interno"],
                "descricao": row["descricao"],
                "etapa": row["etapa"] or None,
                "disciplina": row["disciplina"] or "SINAPI",
                "unidade": row["unidade"] or "NA",
                "ativo": True,
            }
        )

    compositions = []
    for row in read_csv(import_dir / "sinapi_composicoes.csv"):
        compositions.append(
            {
                "id": row["id"],
                "servico_id": row["servico_id"],
                "fonte_id": fonte_id,
                "codigo": row["codigo"],
                "descricao": row["descricao"],
                "unidade": row["unidade"] or "NA",
                "uf": row["uf"],
                "competencia": row["competencia"],
                "regime": row["regime"],
                "origem": "oficial",
                "custo_unitario": clean_decimal(row["custo_unitario"]) or "0",
                "ativo": True,
                "codigo_origem": row["composition_key"],
                "arquivo_origem": "SINAPI_Referência",
                "aba_origem": "CSD/CCD",
                "referencia_preco": row["referencia_preco"],
                "metadados": {
                    "grupo": row["grupo"],
                    "as_percent": clean_decimal(row["as_percent"]),
                    "composition_key": row["composition_key"],
                },
            }
        )

    items = []
    for row in read_csv(import_dir / "sinapi_itens.csv"):
        try:
            metadata = json.loads(row["metadados"] or "{}")
        except json.JSONDecodeError:
            metadata = {}
        items.append(
            {
                "id": row["id"],
                "composicao_id": row["composicao_id"],
                "tipo": row["tipo"],
                "codigo": row["codigo"],
                "descricao": row["descricao"],
                "unidade": row["unidade"] or "NA",
                "coeficiente": clean_decimal(row["coeficiente"]) or "0",
                "preco_unitario": clean_decimal(row["preco_unitario"]) or "0",
                "custo_parcial": clean_decimal(row["custo_parcial"]) or "0",
                "ordem": clean_int(row["ordem"]) or 0,
                "codigo_origem": row["item_key"],
                "arquivo_origem": "SINAPI_Referência",
                "aba_origem": "Analítico",
                "linha_origem": clean_int(row["linha_origem"]),
                "metadados": metadata,
            }
        )

    return services, compositions, items


def main() -> int:
    supabase_url = os.environ.get("SUPABASE_URL", DEFAULT_SUPABASE_URL).strip()
    service_key = env_required("SUPABASE_SERVICE_ROLE_KEY")
    import_dir = Path(os.environ.get("SINAPI_IMPORT_DIR", str(DEFAULT_IMPORT_DIR)))
    batch_size = int(os.environ.get("SINAPI_BATCH_SIZE", str(DEFAULT_BATCH_SIZE)))

    client = SupabaseRest(supabase_url, service_key)
    fonte_id = ensure_fonte(client)
    services, compositions, items = load_payloads(import_dir, fonte_id)

    print(f"Importando SINAPI de {import_dir}")
    print(f"Servicos: {len(services)} | Composicoes: {len(compositions)} | Itens: {len(items)}")

    client.upsert("orc_servicos", services, on_conflict="id", batch_size=batch_size)
    client.upsert("orc_composicoes", compositions, on_conflict="id", batch_size=batch_size)
    client.upsert("orc_composicao_itens", items, on_conflict="id", batch_size=batch_size)
    print("Importacao SINAPI concluida.")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(str(exc), file=sys.stderr)
        raise SystemExit(1)
