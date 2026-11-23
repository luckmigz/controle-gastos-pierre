from datetime import datetime
import json
import os
import re
from typing import Any

import gspread
from google.oauth2.service_account import Credentials
import httpx

PIERRE_BASE_URL = os.getenv("PIERRE_BASE_URL", "https://pierre.finance").rstrip("/")
SHEET_ID = os.environ["GOOGLE_SHEET_ID"]
WORKSHEET_NAME = os.getenv("GOOGLE_WORKSHEET_NAME", "Gastos totais")

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# Mapeia categorias da API do Pierre para as 9 categorias do seu Dashboard
MAPA_CATEGORIAS = {
    "food": "Comida",
    "groceries": "Comida",
    "restaurant": "Comida",
    "alimentacao": "Comida",
    "shopping": "Compras",
    "compras": "Compras",
    "entertainment": "Entretenimento",
    "entretenimento": "Entretenimento",
    "leisure": "Entretenimento",
    "subscriptions": "Assinaturas",
    "assinaturas": "Assinaturas",
    "health": "Saúde",
    "saude": "Saúde",
    "pharmacy": "Saúde",
    "transport": "Transporte",
    "transporte": "Transporte",
    "fuel": "Transporte",
    "bank_fees": "Tarifas Bancárias",
    "tarifas": "Tarifas Bancárias",
    "payments": "Pagamentos",
    "pagamentos": "Pagamentos",
    "transfer": "Pagamentos",
}


def carregar_contas_pierre() -> list[dict[str, str]]:
    contas = []
    if os.getenv("PIERRE_KEY_1"):
        contas.append({
            "pessoa": os.getenv("PIERRE_NOME_1", "Lucas"),
            "api_key": os.environ["PIERRE_KEY_1"],
        })
    if os.getenv("PIERRE_KEY_2"):
        contas.append({
            "pessoa": os.getenv("PIERRE_NOME_2", "Pessoa 2"),
            "api_key": os.environ["PIERRE_KEY_2"],
        })
    if not contas:
        raise ValueError("Configure PIERRE_KEY_1 e/ou PIERRE_KEY_2.")
    return contas


def conectar_planilha() -> gspread.Worksheet:
    creds_dict = json.loads(os.environ["GOOGLE_CREDENTIALS_JSON"])
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    client = gspread.authorize(creds)
    return client.open_by_key(SHEET_ID).worksheet(WORKSHEET_NAME)


def buscar_transacoes_pierre(api_key: str) -> list[dict[str, Any]]:
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


def extrair_parcelas(tx: dict[str, Any], descricao: str) -> tuple[Any, Any]:
    """Tenta obter ParcelaAtual e ParcelaTotal do JSON ou da descrição (ex: 2/6)."""
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


def mapear_categoria(raw_cat: str) -> str:
    chave = raw_cat.strip().lower()
    if chave in MAPA_CATEGORIAS:
        return MAPA_CATEGORIAS[chave]
    for termo, destino in MAPA_CATEGORIAS.items():
        if termo in chave:
            return destino
    return "Outros"


def formatar_para_sua_planilha(tx: dict[str, Any], pessoa: str) -> list[Any] | None:
    tx_id = str(tx.get("id") or tx.get("transactionId") or "").strip()
    if not tx_id:
        return None

    descricao = str(
        tx.get("description") or tx.get("merchant") or tx.get("name") or "Sem descrição"
    ).strip()

    parcela_atual, parcela_total = extrair_parcelas(tx, descricao)

    # Converte data ISO (YYYY-MM-DD) para DD/MM/AAAA e MM/AAAA
    data_iso = str(tx.get("date") or tx.get("createdAt") or tx.get("postedAt") or "")[:10]
    try:
        dt = datetime.strptime(data_iso, "%Y-%m-%d")
        data_br = dt.strftime("%d/%m/%Y")
        mes_ano = dt.strftime("%m/%Y")
    except ValueError:
        data_br = data_iso
        mes_ano = ""

    # Categoria
    cat_obj = tx.get("category")
    raw_cat = cat_obj.get("name", "") if isinstance(cat_obj, dict) else str(cat_obj or tx.get("categoryName") or "")
    categoria = mapear_categoria(raw_cat)

    # Origem (Conta Corrente vs Cartão de Crédito)
    conta_obj = tx.get("account")
    tipo_conta = (
        str(conta_obj.get("type") or conta_obj.get("name") or "")
        if isinstance(conta_obj, dict)
        else str(tx.get("accountType") or tx.get("accountName") or "")
    ).lower()

    eh_cartao = any(k in tipo_conta for k in ("credit", "cartao", "cartão", "card"))
    origem = "Cartão de Crédito" if eh_cartao else "Conta Corrente"

    # Valor e Valor Tratado seguindo exatamente a regra da sua planilha
    valor_raw = float(tx.get("amount", 0) or 0)
    if eh_cartao:
        valor = abs(valor_raw)
        valor_tratado = -abs(valor_raw)
    else:
        valor = valor_raw
        valor_tratado = valor_raw

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

    # Coluna K (índice 11) guarda os IDs do Pierre já importados
    ids_existentes = set(worksheet.col_values(11)[1:])
    novas_linhas: list[list[Any]] = []

    for conta_cfg in carregar_contas_pierre():
        pessoa = conta_cfg["pessoa"]
        transacoes = buscar_transacoes_pierre(conta_cfg["api_key"])

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