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
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "").strip() or None
GEMINI_MODEL = os.getenv("GEMINI_MODEL", "gemini-2.5-flash").strip()

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

CATEGORIAS_VALIDAS = (
    "Comida",
    "Compras",
    "Entretenimento",
    "Assinaturas",
    "Saúde",
    "Transporte",
    "Tarifas Bancárias",
    "Pagamentos",
    "Outros",
)

SUBCATEGORIAS_VALIDAS = (
    "Delivery",
    "Mercado / Padaria",
    "Restaurantes / Lanches",
    "Moda / Vestuário",
    "Marketplaces / Variedades",
    "Casa / Pets / Eletrônicos",
    "Jogos / Streaming",
    "Serviços Digitais / Apps",
    "Combustível / Estacionamento",
    "Saúde / Farmácia",
    "Tarifas / Impostos",
    "Fatura Cartão",
    "Pix / Transferências",
    "Investimentos (RDB)",
    "Entradas / Recebimentos",
    "Outros",
)

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

REGRAS_CLASSIFICACAO: list[tuple[str, str, tuple[str, ...]]] = [
    ("Outros", "Investimentos (RDB)", ("rdb", "limite convertido")),
    ("Pagamentos", "Fatura Cartão", ("pagamento de fatura", "pagamento recebido")),
    (
        "Outros",
        "Entradas / Recebimentos",
        (
            "pix recebido",
            "transferência recebida",
            "transferencia recebida",
            "liquido de vencimento",
            "líquido de vencimento",
            "estorno de pagamento",
            "crédito de parcelamento",
        ),
    ),
    (
        "Pagamentos",
        "Pix / Transferências",
        (
            "pix enviado",
            "pix no crédito",
            "pix no credito",
            "transferência enviada",
            "transferencia enviada",
        ),
    ),
    ("Tarifas Bancárias", "Tarifas / Impostos", ("iof", "anual", "tarifa")),
    ("Comida", "Delivery", ("ifd*",)),
    (
        "Comida",
        "Mercado / Padaria",
        (
            "carrefour",
            "pao de acucar",
            "pão de açúcar",
            "padaria",
            "oxxo",
            "minuto pa",
            "hortus",
            "frutas",
            "panificadora",
        ),
    ),
    (
        "Comida",
        "Restaurantes / Lanches",
        (
            "outback",
            "burger king",
            "pizza hut",
            "mania de churrasco",
            "villaportugal",
            "costela",
            "bistro",
            "frangoassado",
            "nema",
            "draft rock",
            "bolognani",
            "quintal secreto",
            "mordidita",
            "lindt",
            "mmb comercio",
            "rancho dos orixas",
        ),
    ),
    (
        "Compras",
        "Moda / Vestuário",
        (
            "shein",
            "relvaco",
            "marcyn",
            "inditex",
            "corello",
            "lojareversa",
            "darknesssport",
            "romanticcrown",
            "uoustore",
            "renovadora de ca",
        ),
    ),
    (
        "Compras",
        "Marketplaces / Variedades",
        ("shopee", "mp *", "lojaonlinede", "shop kids"),
    ),
    (
        "Compras",
        "Casa / Pets / Eletrônicos",
        ("kabum", "petisk", "firehouse", "jessica abud"),
    ),
    (
        "Entretenimento",
        "Jogos / Streaming",
        ("playstation", "crunchyroll", "steam", "bolix", "cinemark", "eventim", "stray fawn"),
    ),
    (
        "Assinaturas",
        "Serviços Digitais / Apps",
        ("telegram", "google*one", "bumble", "overgeared", "vindi", "jim.com"),
    ),
    (
        "Transporte",
        "Combustível / Estacionamento",
        ("posto", "zul ", "lava rapido", "av dos estados"),
    ),
    (
        "Saúde",
        "Saúde / Farmácia",
        ("drogasil", "drogaria", "odonto", "iort", "peau sante", "farmapur"),
    ),
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


def classificar_por_regras(tx: dict[str, Any], descricao: str) -> tuple[str, str]:
    texto = descricao.lower()
    for categoria, subcategoria, termos in REGRAS_CLASSIFICACAO:
        if any(termo in texto for termo in termos):
            return categoria, subcategoria

    cat_obj = tx.get("category")
    raw_cat = (
        str(cat_obj.get("name") or "")
        if isinstance(cat_obj, dict)
        else str(cat_obj or tx.get("categoryName") or "")
    ).lower()

    mapa_fallback = [
        (("food", "grocer", "restaurant", "comida"), "Comida"),
        (("shop", "compra", "retail"), "Compras"),
        (("entertain", "leisure", "entretenimento"), "Entretenimento"),
        (("subscri", "assinatura"), "Assinaturas"),
        (("health", "saude", "saúde", "pharm"), "Saúde"),
        (("transp", "fuel"), "Transporte"),
        (("fee", "tarifa"), "Tarifas Bancárias"),
        (("pay", "transfer", "pagamento"), "Pagamentos"),
    ]
    for chaves, cat in mapa_fallback:
        if any(k in raw_cat for k in chaves):
            return cat, cat

    return "Outros", "Outros"


def classificar_lote_com_gemini(itens: list[dict[str, Any]]) -> dict[int, tuple[str, str]]:
    """Classifica em lote apenas as descrições que caíram em 'Outros'."""
    if not GEMINI_API_KEY or not itens:
        return {}

    prompt = (
        "Classifique cada transação bancária brasileira abaixo escolhendo EXATAMENTE uma "
        f"categoria dentre {list(CATEGORIAS_VALIDAS)} e uma subcategoria dentre "
        f"{list(SUBCATEGORIAS_VALIDAS)}.\n"
        "Responda SOMENTE com um array JSON no formato: "
        '[{"idx": 0, "categoria": "...", "subcategoria": "..."}].\n\n'
        f"Transações:\n{json.dumps(itens, ensure_ascii=False)}"
    )

    url = (
        f"https://generativelanguage.googleapis.com/v1beta/models/"
        f"{GEMINI_MODEL}:generateContent?key={GEMINI_API_KEY}"
    )
    body = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {
            "temperature": 0.1,
            "responseMimeType": "application/json",
        },
    }

    try:
        with httpx.Client(timeout=30.0) as client:
            resp = client.post(url, json=body)
            resp.raise_for_status()
            data = resp.json()

        texto_json = data["candidates"][0]["content"]["parts"][0]["text"]
        resultados = json.loads(texto_json)
        mapeado: dict[int, tuple[str, str]] = {}
        for item in resultados:
            idx = int(item["idx"])
            cat = item.get("categoria", "Outros")
            sub = item.get("subcategoria", "Outros")
            if cat not in CATEGORIAS_VALIDAS:
                cat = "Outros"
            if sub not in SUBCATEGORIAS_VALIDAS and sub not in CATEGORIAS_VALIDAS:
                sub = cat
            mapeado[idx] = (cat, sub)
        return mapeado
    except Exception as exc:
        print(f"Aviso: fallback Gemini indisponível ({exc}). Mantendo classificação padrão.")
        return {}


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
    categoria, subcategoria = classificar_por_regras(tx, descricao)

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
        subcategoria,   # L: Subcategoria
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

    # Fallback Gemini apenas para transações classificadas como "Outros"
    pendentes_ia = [
        {"idx": i, "descricao": linha[0], "valor": linha[8], "origem": linha[7]}
        for i, linha in enumerate(novas_linhas)
        if linha[6] == "Outros" and linha[11] == "Outros"
    ]
    if pendentes_ia:
        classificacoes_ia = classificar_lote_com_gemini(pendentes_ia)
        for idx, (cat_ia, sub_ia) in classificacoes_ia.items():
            novas_linhas[idx][6] = cat_ia
            novas_linhas[idx][11] = sub_ia

    if novas_linhas:
        worksheet.append_rows(novas_linhas, value_input_option="USER_ENTERED")
        print(f"{len(novas_linhas)} nova(s) transação(ões) adicionada(s) em '{WORKSHEET_NAME}'.")
    else:
        print("Nenhuma transação nova encontrada.")


if __name__ == "__main__":
    main()