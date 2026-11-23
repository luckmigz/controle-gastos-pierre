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

# CORREÇÃO: O modelo correto e atualizado na API é o 1.5-flash
GEMINI_MODEL = os.getenv("GEMINI_MODEL", "gemini-1.5-flash").strip()

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

CATEGORIAS_VALIDAS = (
    "Comida", "Compras", "Entretenimento", "Assinaturas", "Saúde", 
    "Transporte", "Tarifas Bancárias", "Pagamentos", "Outros"
)

SUBCATEGORIAS_VALIDAS = (
    "Delivery", "Mercado / Padaria", "Restaurantes / Lanches", "Moda / Vestuário",
    "Marketplaces / Variedades", "Casa / Pets / Eletrônicos", "Jogos / Streaming",
    "Serviços Digitais / Apps", "Combustível / Estacionamento", "Saúde / Farmácia",
    "Tarifas / Impostos", "Fatura Cartão", "Pix / Transferências",
    "Investimentos (RDB)", "Entradas / Recebimentos", "Outros"
)

KEYWORDS_CONTA_CORRENTE = (
    "pix ", "transferência", "transferencia", "compra no débito", "compra no debito",
    "compra débito", "compra debito", "resgate rdb", "aplicação rdb", "aplicacao rdb",
    "pagamento de fatura", "pagamento de imposto", "liquido de vencimento",
    "líquido de vencimento", "limite convertido"
)

REGRAS_CLASSIFICACAO: list[tuple[str, str, tuple[str, ...]]] = [
    ("Outros", "Investimentos (RDB)", ("rdb", "limite convertido")),
    ("Pagamentos", "Fatura Cartão", ("pagamento de fatura", "pagamento recebido")),
    ("Outros", "Entradas / Recebimentos", ("pix recebido", "transferência recebida", "transferencia recebida", "liquido de vencimento", "líquido de vencimento", "estorno de pagamento", "crédito de parcelamento")),
    ("Pagamentos", "Pix / Transferências", ("pix enviado", "pix no crédito", "pix no credito", "transferência enviada", "transferencia enviada")),
    ("Tarifas Bancárias", "Tarifas / Impostos", ("iof", "anual", "tarifa")),
    ("Comida", "Delivery", ("ifd*",)),
    ("Comida", "Mercado / Padaria", ("carrefour", "pao de acucar", "pão de açúcar", "padaria", "oxxo", "minuto pa", "hortus", "frutas", "panificadora")),
    ("Comida", "Restaurantes / Lanches", ("outback", "burger king", "pizza hut", "mania de churrasco", "villaportugal", "costela", "bistro", "frangoassado", "nema", "draft rock", "bolognani", "quintal secreto", "mordidita", "lindt", "mmb comercio", "rancho dos orixas")),
    ("Compras", "Moda / Vestuário", ("shein", "relvaco", "marcyn", "inditex", "corello", "lojareversa", "darknesssport", "romanticcrown", "uoustore", "renovadora de ca")),
    ("Compras", "Marketplaces / Variedades", ("shopee", "mp *", "lojaonlinede", "shop kids")),
    ("Compras", "Casa / Pets / Eletrônicos", ("kabum", "petisk", "firehouse", "jessica abud")),
    ("Entretenimento", "Jogos / Streaming", ("playstation", "crunchyroll", "steam", "bolix", "cinemark", "eventim", "stray fawn")),
    ("Assinaturas", "Serviços Digitais / Apps", ("telegram", "google*one", "bumble", "overgeared", "vindi", "jim.com")),
    ("Transporte", "Combustível / Estacionamento", ("posto", "zul ", "lava rapido", "av dos estados")),
    ("Saúde", "Saúde / Farmácia", ("drogasil", "drogaria", "odonto", "iort", "peau sante", "farmapur")),
]


def carregar_contas_pierre() -> list[dict[str, str]]:
    contas = []
    if os.getenv("PIERRE_KEY_1"):
        contas.append({"pessoa": os.getenv("PIERRE_NOME_1", "Lucas"), "api_key": os.environ["PIERRE_KEY_1"].strip()})
    if os.getenv("PIERRE_KEY_2"):
        contas.append({"pessoa": os.getenv("PIERRE_NOME_2", "Raíssa"), "api_key": os.environ["PIERRE_KEY_2"].strip()})
    if not contas:
        raise ValueError("Configure PIERRE_KEY_1 e/ou PIERRE_KEY_2.")
    return contas


def conectar_planilha() -> gspread.Worksheet:
    creds_dict = json.loads(os.environ["GOOGLE_CREDENTIALS_JSON"])
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    client = gspread.authorize(creds)
    return client.open_by_key(SHEET_ID).worksheet(WORKSHEET_NAME)


def buscar_transacoes_pierre(api_key: str, start_date: str | None = None) -> list[dict[str, Any]]:
    headers = {"Authorization": f"Bearer {api_key}", "Accept": "application/json"}
    url = f"{PIERRE_BASE_URL}/tools/api/get-transactions"
    params = {"startDate": start_date} if start_date else {}

    with httpx.Client(timeout=30.0) as client:
        response = client.get(url, headers=headers, params=params)
        response.raise_for_status()
        payload = response.json()

    if isinstance(payload, list): return payload
    if isinstance(payload, dict):
        for chave in ("transactions", "data", "items", "results"):
            if isinstance(payload.get(chave), list): return payload[chave]
    return []


def extrair_parcelas(tx: dict[str, Any], descricao: str) -> tuple[Any, Any]:
    atual = tx.get("installmentNumber") or tx.get("currentInstallment")
    total = tx.get("totalInstallments") or tx.get("installments")
    if atual and total: return int(atual), int(total)

    match = re.search(r"\b(\d{1,2})/(\d{1,2})\b", descricao)
    if match: return int(match.group(1)), int(match.group(2))
    return "", ""


def classificar_origem(tx: dict[str, Any], descricao: str) -> str:
    if any(k in descricao.lower() for k in KEYWORDS_CONTA_CORRENTE): return "Conta Corrente"
    conta_obj = tx.get("account")
    tipo_conta = (str(conta_obj.get("type") or conta_obj.get("name") or "") if isinstance(conta_obj, dict) else str(tx.get("accountType") or tx.get("accountName") or "")).lower()
    if any(k in tipo_conta for k in ("credit", "cartao", "cartão", "card")): return "Cartão de Crédito"
    if any(k in tipo_conta for k in ("checking", "corrente", "conta")): return "Conta Corrente"
    return "Cartão de Crédito"


def classificar_por_regras(tx: dict[str, Any], descricao: str) -> tuple[str, str]:
    texto = descricao.lower()
    for categoria, subcategoria, termos in REGRAS_CLASSIFICACAO:
        if any(termo in texto for termo in termos): return categoria, subcategoria

    cat_obj = tx.get("category")
    raw_cat = (str(cat_obj.get("name") or "") if isinstance(cat_obj, dict) else str(cat_obj or tx.get("categoryName") or "")).lower()
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
        if any(k in raw_cat for k in chaves): return cat, cat
    return "Outros", "Outros"


def classificar_lote_com_gemini(itens: list[dict[str, Any]]) -> list[dict[str, str]]:
    if not GEMINI_API_KEY or not itens: return []

    prompt = (
        "Classifique cada transação bancária brasileira abaixo escolhendo EXATAMENTE uma "
        f"categoria dentre {list(CATEGORIAS_VALIDAS)} e uma subcategoria dentre {list(SUBCATEGORIAS_VALIDAS)}.\n"
        "Responda SOMENTE com um array JSON válido no formato:\n"
        '[{"id_ref": "valor_passado", "categoria": "...", "subcategoria": "..."}]\n\n'
        f"Transações:\n{json.dumps(itens, ensure_ascii=False)}"
    )

    url = f"https://generativelanguage.googleapis.com/v1beta/models/{GEMINI_MODEL}:generateContent?key={GEMINI_API_KEY}"
    body = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {"temperature": 0.1, "responseMimeType": "application/json"}
    }

    try:
        with httpx.Client(timeout=45.0) as client:
            resp = client.post(url, json=body)
            resp.raise_for_status()
            data = resp.json()

        texto_json = data["candidates"][0]["content"]["parts"][0]["text"]
        return json.loads(texto_json)
    except Exception as exc:
        print(f"Aviso: fallback Gemini indisponível ({exc}).")
        return []


def formatar_nova_linha(tx: dict[str, Any], pessoa: str) -> list[Any] | None:
    tx_id = str(tx.get("id") or tx.get("transactionId") or "").strip()
    if not tx_id: return None
    descricao = str(tx.get("description") or tx.get("merchant") or tx.get("name") or "Sem descrição").strip()
    parcela_atual, parcela_total = extrair_parcelas(tx, descricao)
    data_iso = str(tx.get("date") or tx.get("createdAt") or tx.get("postedAt") or "")[:10]
    
    try:
        dt = datetime.strptime(data_iso, "%Y-%m-%d")
        data_br, mes_ano = dt.strftime("%d/%m/%Y"), dt.strftime("%m/%Y")
    except ValueError:
        data_br, mes_ano = data_iso, ""

    origem = classificar_origem(tx, descricao)
    categoria, subcategoria = classificar_por_regras(tx, descricao)
    valor_raw = float(tx.get("amount", 0) or 0)
    valor = round(abs(valor_raw), 2) if origem == "Cartão de Crédito" else round(valor_raw, 2)
    valor_tratado = round(-abs(valor_raw), 2) if origem == "Cartão de Crédito" else round(valor_raw, 2)

    return [descricao, parcela_atual, parcela_total, valor, data_br, mes_ano, categoria, origem, valor_tratado, pessoa, tx_id, subcategoria]


def main() -> None:
    worksheet = conectar_planilha()
    
    # 1. Analisa os dados existentes para achar "Outros" e evitar duplicatas
    dados_planilha = worksheet.get_all_values()
    ids_existentes = set()
    linhas_outros_antigas = []
    
    for i, row in enumerate(dados_planilha, start=1):
        if i == 1: continue
        if len(row) > 10 and row[10].strip():
            ids_existentes.add(row[10].strip())
            
        cat = row[6].strip() if len(row) > 6 else ""
        subcat = row[11].strip() if len(row) > 11 else ""
        if cat == "Outros" and subcat == "Outros":
            linhas_outros_antigas.append({"id_ref": f"row_{i}", "descricao": row[0], "origem": row[7] if len(row) > 7 else ""})

    # 2. Puxa transações novas
    novas_linhas = []
    novas_linhas_outros = []
    
    for conta_cfg in carregar_contas_pierre():
        for tx in buscar_transacoes_pierre(conta_cfg["api_key"], start_date=PIERRE_START_DATE):
            linha = formatar_nova_linha(tx, conta_cfg["pessoa"])
            if linha and linha[10] not in ids_existentes:
                novas_linhas.append(linha)
                ids_existentes.add(linha[10])
                if linha[6] == "Outros" and linha[11] == "Outros":
                    novas_linhas_outros.append({"id_ref": f"new_{len(novas_linhas)-1}", "descricao": linha[0], "origem": linha[7]})

    # 3. Processa TUDO que for "Outros" na Inteligência Artificial (Novas + Antigas)
    todos_pendentes_ia = novas_linhas_outros + linhas_outros_antigas
    atualizacoes_planilha = []

    if todos_pendentes_ia:
        print(f"Enviando {len(todos_pendentes_ia)} registros 'Outros' para reclassificação com Gemini...")
        
        # Divide em lotes para não estourar a janela de contexto da API
        for i in range(0, len(todos_pendentes_ia), 50):
            lote = todos_pendentes_ia[i:i+50]
            classificacoes = classificar_lote_com_gemini(lote)
            
            for res in classificacoes:
                id_ref, cat_ia, sub_ia = res.get("id_ref"), res.get("categoria", "Outros"), res.get("subcategoria", "Outros")
                if not id_ref: continue
                
                # Garante que ele escolha uma categoria válida (correção de segurança)
                if cat_ia not in CATEGORIAS_VALIDAS: cat_ia = "Outros"
                if sub_ia not in SUBCATEGORIAS_VALIDAS and sub_ia not in CATEGORIAS_VALIDAS: sub_ia = cat_ia
                
                if str(id_ref).startswith("new_"):
                    idx = int(str(id_ref).split("_")[1])
                    novas_linhas[idx][6] = cat_ia
                    novas_linhas[idx][11] = sub_ia
                elif str(id_ref).startswith("row_"):
                    row_idx = int(str(id_ref).split("_")[1])
                    atualizacoes_planilha.append(gspread.Cell(row=row_idx, col=7, value=cat_ia))   # Coluna G (Categoria)
                    atualizacoes_planilha.append(gspread.Cell(row=row_idx, col=12, value=sub_ia))  # Coluna L (Subcategoria)

    # 4. Salva as informações na Planilha do Google
    if novas_linhas:
        worksheet.append_rows(novas_linhas, value_input_option="USER_ENTERED")
        print(f"{len(novas_linhas)} nova(s) transação(ões) adicionada(s).")
    else:
        print("Nenhuma transação nova encontrada.")

    if atualizacoes_planilha:
        worksheet.update_cells(atualizacoes_planilha)
        print(f"{len(atualizacoes_planilha) // 2} transação(ões) antigas foram reclassificadas pela IA com sucesso!")


if __name__ == "__main__":
    main()