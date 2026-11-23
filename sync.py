import json
import os
from typing import Any

import gspread
from google.oauth2.service_account import Credentials
import httpx

PIERRE_BASE_URL = os.getenv("PIERRE_BASE_URL", "https://pierre.finance").rstrip("/")
SHEET_ID = os.environ["GOOGLE_SHEET_ID"]
WORKSHEET_NAME = os.getenv("GOOGLE_WORKSHEET_NAME", "Transacoes")

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


def carregar_contas_pierre() -> list[dict[str, str]]:
    """Carrega as chaves configuradas nas variáveis de ambiente."""
    contas = []
    if os.getenv("PIERRE_KEY_1"):
        contas.append({
            "pessoa": os.getenv("PIERRE_NOME_1", "Pessoa 1"),
            "api_key": os.environ["PIERRE_KEY_1"],
        })
    if os.getenv("PIERRE_KEY_2"):
        contas.append({
            "pessoa": os.getenv("PIERRE_NOME_2", "Pessoa 2"),
            "api_key": os.environ["PIERRE_KEY_2"],
        })
    if not contas:
        raise ValueError("Nenhuma PIERRE_KEY_1 ou PIERRE_KEY_2 foi configurada.")
    return contas


def conectar_planilha() -> gspread.Worksheet:
    """Autentica com a Service Account via JSON em variável de ambiente."""
    creds_raw = os.environ["GOOGLE_CREDENTIALS_JSON"]
    creds_dict = json.loads(creds_raw)
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    client = gspread.authorize(creds)
    spreadsheet = client.open_by_key(SHEET_ID)
    return spreadsheet.worksheet(WORKSHEET_NAME)


def buscar_transacoes_pierre(api_key: str) -> list[dict[str, Any]]:
    """Consulta o endpoint de transações do Pierre."""
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Accept": "application/json",
    }
    url = f"{PIERRE_BASE_URL}/api/v1/transactions"

    with httpx.Client(timeout=30.0) as client:
        response = client.get(url, headers=headers)
        response.raise_for_status()
        payload = response.json()

    if isinstance(payload, list):
        return payload
    if isinstance(payload, dict):
        for chave in ("transactions", "data", "items", "results"):
            if isinstance(payload.get(chave), list):
                return payload[chave]
    return []


def normalizar_transacao(tx: dict[str, Any], pessoa: str) -> list[Any] | None:
    """Converte uma transação do JSON para a linha padrão da planilha."""
    tx_id = str(tx.get("id") or tx.get("transactionId") or "").strip()
    if not tx_id:
        return None

    data = str(
        tx.get("date")
        or tx.get("createdAt")
        or tx.get("postedAt")
        or ""
    )[:10]

    descricao = str(
        tx.get("description")
        or tx.get("merchant")
        or tx.get("name")
        or "Sem descrição"
    ).strip()

    categoria_obj = tx.get("category")
    if isinstance(categoria_obj, dict):
        categoria = str(categoria_obj.get("name") or "Outros")
    else:
        categoria = str(categoria_obj or tx.get("categoryName") or "Outros")

    conta_obj = tx.get("account")
    if isinstance(conta_obj, dict):
        conta = str(conta_obj.get("name") or conta_obj.get("institution") or "Conta")
    else:
        conta = str(tx.get("accountName") or tx.get("accountId") or "Conta")

    valor_raw = tx.get("amount", 0)
    try:
        valor = float(valor_raw)
    except (TypeError, ValueError):
        valor = 0.0

    tipo = str(tx.get("type") or ("expense" if valor < 0 else "income")).lower()

    return [
        tx_id,
        pessoa,
        data,
        descricao,
        categoria,
        conta,
        tipo,
        valor,
    ]


def main() -> None:
    worksheet = conectar_planilha()

    # Coluna A contém os IDs já salvos
    ids_existentes = set(worksheet.col_values(1)[1:])
    novas_linhas: list[list[Any]] = []

    for conta_cfg in carregar_contas_pierre():
        pessoa = conta_cfg["pessoa"]
        api_key = conta_cfg["api_key"]

        transacoes = buscar_transacoes_pierre(api_key)
        for tx in transacoes:
            linha = normalizar_transacao(tx, pessoa)
            if linha is None:
                continue
            tx_id = linha[0]
            if tx_id not in ids_existentes:
                novas_linhas.append(linha)
                ids_existentes.add(tx_id)

    if novas_linhas:
        worksheet.append_rows(novas_linhas, value_input_option="USER_ENTERED")
        print(f"{len(novas_linhas)} nova(s) transação(ões) adicionada(s).")
    else:
        print("Nenhuma transação nova encontrada.")


if __name__ == "__main__":
    main()