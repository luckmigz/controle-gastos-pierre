from datetime import datetime
import json
import os
import re
from typing import Any

import gspread
from google.oauth2.service_account import Credentials
import httpx

PIERRE_BASE_URL = os.getenv("PIERRE_BASE_URL", "https://pierre.finance").rstrip("/")
PIERRE_START_DATE = os.getenv("PIERRE_START_DATE", "").strip() or None
SHEET_ID = os.environ["GOOGLE_SHEET_ID"].strip()
WORKSHEET_NAME = os.getenv("GOOGLE_WORKSHEET_NAME", "Gastos totais").strip()

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

KEYWORDS_CONTA_CORRENTE = (
    "pix ",
    "transferência",
    "transferencia",
    "compra no débito",
    "compra no debito",
    "compra débito",
    "compra debito",
    "resgate rdb",
    "aplicação rdb",
    "aplicacao rdb",
    "pagamento de fatura",
    "pagamento de imposto",
    "liquido de vencimento",
    "líquido de vencimento",
    "limite convertido",
)

REGRAS_CATEGORIA: list[tuple[str, tuple[str, ...]]] = [
    ("Tarifas Bancárias", ("iof", "anual", "tarifa")),
    (
        "Comida",
        (
            "padaria",
            "villaportugal",
            "ifd*",
            "oxxo",
            "rest ",
            "frangoassado",
            "minuto pa",
            "nema",
            "draft rock",
            "bolognani",
            "carrefour",
            "pao de acucar",
            "pão de açúcar",
            "panificadora",
        ),
    ),
    (
        "Compras",
        (
            "shein",
            "marcyn",
            "relvaco",
            "shop kids",
            "petisk",
            "lojaonlinede",
            "lojareversa",
            "firehouse",
            "jessica abud",
            "mp *pp",
            "shopee",
        ),
    ),
    ("Transporte", ("posto", "zul ")),
    ("Entretenimento", ("playstation", "eventim", "bolix", "cinemark", "steam")),
    (
        "Assinaturas",
        ("telegram", "google*one", "jim.com", "overgeared", "crunchyroll", "bumble"),
    ),
    ("Saúde", ("drogasil", "odonto", "iort", "peau sante")),
    ("Pagamentos", ("pix ", "transferência", "transferencia", "pagamento de fatura", "pagamento recebido")),
]


def carregar_contas_pierre() -> list[dict[str, str]]:
    contas = []
    if os.getenv("PIERRE_KEY_1"):
        contas.append({
            "pessoa": os.getenv("PIERRE_NOME_1", "Lucas"),
            "api_key": os.environ["PIERRE_KEY_1"].strip(),
        })
    if os.getenv("PIERRE_KEY_2"):
        contas.append({
            "pessoa": os.getenv("PIERRE_NOME_2", "Pessoa 2"),
            "api_key": os.environ["PIERRE_KEY_2"].strip(),
        })
    if not contas:
        raise ValueError("Configure PIERRE_KEY_1 e/ou PIERRE_KEY_2.")
    return contas


def conectar_planilha() -> gspread.Worksheet:
    creds_dict = json.loads(os.environ["GOOGLE_CREDENTIALS_JSON"])
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    client = gspread.authorize(creds)
    return client.open_by_key(SHEET_ID).worksheet(WORKSHEET_NAME)


def buscar_transacoes_pierre(
    api_key: str,
    start_date: str | None = None,
    end_date: str | None = None,
) -> list[dict[str, Any]]:
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Accept": "application/json",
    }
    url = f"{PIERRE_BASE_URL}/tools/api/get-transactions"
    params: dict[str, str] = {}
    if start_date:
        params["startDate"] = start_date
    if end_date:
        params["endDate"] = end_date

    with httpx.Client(timeout=30.0) as client:
        response = client.get(url, headers=headers, params=params)
        if response.status_code >= 400:
            print(f"Erro HTTP {response.status_code}: {response.text[:500]}")
            response.raise_for_status()
        payload = response.json()

    if isinstance(payload, list):
        return payload
    if isinstance(payload, dict):
        for chave in ("transactions", "data", "items", "results"):
            if isinstance(payload.get(chave), list):
                return payload[chave]
    return []


def extrair_parcelas(tx: dict[str, Any], descricao: str) -> tuple[Any, Any]:
    atual = tx.get("installmentNumber") or tx.get("currentInstallment")
    total = tx.get("totalInstallments") or tx.get("installments")
    if atual and total:
        return int(atual), int(total)

    match = re.search(r"\b(\d{1,2})/(\d{1,2})\b", descricao)
    if match:
        p_atual, p_total = int(match.group(1)), int(match.group(2))
        if 1 <= p_atual <= p_total <= 48:
            return p_atual, p_total
    return "", ""


def classificar_origem(tx: dict[str, Any], descricao: str) -> str:
    desc_lower = descricao.lower()
    if any(k in desc_lower for k in KEYWORDS_CONTA_CORRENTE):
        return "Conta Corrente"

    conta_obj = tx.get("account")
    tipo_conta = (
        str(conta_obj.get("type") or conta_obj.get("name") or "")
        if isinstance(conta_obj, dict)
        else str(tx.get("accountType") or tx.get("accountName") or "")
    ).lower()

    if any(k in tipo_conta for k in ("credit", "cartao", "cartão", "card")):
        return "Cartão de Crédito"
    if any(k in tipo_conta for k in ("checking", "corrente", "conta")):
        return "Conta Corrente"

    return "Cartão de Crédito"


def classificar_categoria(tx: dict[str, Any], descricao: str) -> str:
    texto = descricao.lower()
    for categoria, termos in REGRAS_CATEGORIA:
        if any(termo in texto for termo in termos):
            return categoria

    cat_obj = tx.get("category")
    raw_cat = (
        str(cat_obj.get("name") or "")
        if isinstance(cat_obj, dict)
        else str(cat_obj or tx.get("categoryName") or "")
    ).lower()

    if any(k in raw_cat for k in ("food", "grocer", "restaurant", "comida")):
        return "Comida"
    if any(k in raw_cat for k in ("shop", "compra", "retail")):
        return "Compras"
    if any(k in raw_cat for k in ("entertain", "leisure", "entretenimento")):
        return "Entretenimento"
    if any(k in raw_cat for k in ("subscri", "assinatura")):
        return "Assinaturas"
    if any(k in raw_cat for k in ("health", "saude", "saúde", "pharm")):
        return "Saúde"
    if any(k in raw_cat for k in ("transp", "fuel")):
        return "Transporte"
    if any(k in raw_cat for k in ("fee", "tarifa")):
        return "Tarifas Bancárias"
    if any(k in raw_cat for k in ("pay", "transfer", "pagamento")):
        return "Pagamentos"

    return "Outros"


def formatar_para_sua_planilha(tx: dict[str, Any], pessoa: str) -> list[Any] | None:
    tx_id = str(tx.get("id") or tx.get("transactionId") or "").strip()
    if not tx_id:
        return None

    descricao = str(
        tx.get("description") or tx.get("merchant") or tx.get("name") or "Sem descrição"
    ).strip()

    parcela_atual, parcela_total = extrair_parcelas(tx, descricao)

    data_iso = str(tx.get("date") or tx.get("createdAt") or tx.get("postedAt") or "")[:10]
    try:
        dt = datetime.strptime(data_iso, "%Y-%m-%d")
        data_br = dt.strftime("%d/%m/%Y")
        mes_ano = dt.strftime("%m/%Y")
    except ValueError:
        data_br = data_iso
        mes_ano = ""

    origem = classificar_origem(tx, descricao)
    categoria = classificar_categoria(tx, descricao)

    valor_raw = float(tx.get("amount", 0) or 0)
    if origem == "Cartão de Crédito":
        valor = round(abs(valor_raw), 2)
        valor_tratado = round(-abs(valor_raw), 2)
    else:
        valor = round(valor_raw, 2)
        valor_tratado = round(valor_raw, 2)

    return [
        descricao,      # A: Gasto
        parcela_atual,  # B: ParcelaAtual
        parcela_total,  # C: ParcelaTotal
        valor,          # D: Valor
        data_br,        # E: Data (DD/MM/AAAA)
        mes_ano,        # F: Mês/Ano (MM/AAAA)
        categoria,      # G: Categoria
        origem,         # H: Origem
        valor_tratado,  # I: Valor Tratado
        pessoa,         # J: Pessoa
        tx_id,          # K: ID_Pierre
    ]


def main() -> None:
    worksheet = conectar_planilha()
    ids_existentes = set(worksheet.col_values(11)[1:])
    novas_linhas: list[list[Any]] = []

    for conta_cfg in carregar_contas_pierre():
        pessoa = conta_cfg["pessoa"]
        transacoes = buscar_transacoes_pierre(
            conta_cfg["api_key"],
            start_date=PIERRE_START_DATE,
        )

        for tx in transacoes:
            linha = formatar_para_sua_planilha(tx, pessoa)
            if linha is None:
                continue
            tx_id = linha[10]
            if tx_id not in ids_existentes:
                novas_linhas.append(linha)
                ids_existentes.add(tx_id)

    if novas_linhas:
        worksheet.append_rows(novas_linhas, value_input_option="USER_ENTERED")
        print(f"{len(novas_linhas)} nova(s) transação(ões) adicionada(s) em '{WORKSHEET_NAME}'.")
    else:
        print("Nenhuma transação nova encontrada.")


if __name__ == "__main__":
    main()