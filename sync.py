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

# Modelo configurável pela variável GEMINI_MODEL do repositório (a família 1.5 foi descontinuada pelo Google).
GEMINI_MODEL = os.getenv("GEMINI_MODEL", "").strip() or "gemini-2.5-flash"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

CATEGORIAS_VALIDAS = (
    "Moradia / Contas", "Comida", "Compras", "Entretenimento", "Assinaturas", "Saúde",
    "Transporte", "Tarifas Bancárias", "Outros"
)

SUBCATEGORIAS_VALIDAS = (
    "Delivery", "Mercado / Padaria", "Restaurantes / Lanches", "Moda / Vestuário",
    "Marketplaces / Variedades", "Casa / Pets / Eletrônicos", "Jogos / Streaming",
    "Serviços Digitais / Apps", "Combustível / Estacionamento", "Saúde / Farmácia",
    "Tarifas / Impostos", "Fatura Cartão", "Pix / Transferências",
    "Investimentos (RDB)", "Entradas / Recebimentos", "Aluguel", "Contas da casa",
    "Seguros", "Transporte", "Transferência interna", "Empréstimo", "Outros"
)

KEYWORDS_CONTA_CORRENTE = (
    "pix ", "transferência", "transferencia", "compra no débito", "compra no debito",
    "compra débito", "compra debito", "resgate rdb", "aplicação rdb", "aplicacao rdb",
    "pagamento de fatura", "pagamento de imposto", "liquido de vencimento",
    "líquido de vencimento", "limite convertido", "salário", "salario", "consignado",
    "cofrinho", "rendimento", "crédito em conta", "credito em conta", "devolução de recebimento",
    "devolucao de recebimento"
)

# Lançamentos de cartão que são créditos (entram positivos), não compras.
KEYWORDS_CREDITO_CARTAO = (
    "pagamento recebido", "estorno", "reembolso", "crédito de parcelamento", "credito de parcelamento"
)

# Nomes do casal: Pix/transferências entre eles (ou para si mesmo) não são gasto nem renda.
NOMES_CASAL = tuple(
    n.strip().lower() for n in os.getenv("NOMES_CASAL", "lucas miguel de matos negri,raissa sanches crevelenti,raíssa sanches crevelenti").split(",") if n.strip()
)

REGRAS_CLASSIFICACAO: list[tuple[str, str, tuple[str, ...]]] = [
    ("Empréstimos", "Empréstimo", ("consignado",)),
    ("Investimentos", "Investimentos (RDB)", ("rdb", "cofrinho")),
    ("Transferências internas", "Transferência interna", ("limite convertido", "valor adicionado", "reembolso de pix no crédito", "reembolso de pix no credito")),
    ("Pagamento de fatura", "Fatura Cartão", ("pagamento de fatura", "pagamento recebido")),
    ("Moradia / Contas", "Aluguel", ("imoveis", "imóveis", "imobiliaria", "imobiliária")),
    ("Moradia / Contas", "Contas da casa", ("sabesp", "eletropaulo", "enel", "claro", "nucel", "vivo", "tim ")),
    ("Moradia / Contas", "Seguros", ("seguro",)),
    ("Comida", "Delivery", ("ifood", "99 food")),
    ("Transporte", "Transporte", ("uber", "lalamove", "99app", "ituran", "unidas", "mercadocar")),
    ("Outros", "Entradas / Recebimentos", ("pix recebido", "transferência recebida", "transferencia recebida", "liquido de vencimento", "líquido de vencimento", "estorno de pagamento", "crédito de parcelamento")),
    ("Outros", "Pix / Transferências", ("pix enviado", "pix no crédito", "pix no credito", "transferência enviada", "transferencia enviada")),
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
        (("pay", "transfer", "pagamento"), "Outros"),
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


def classificar_tipo(descricao: str, valor_tratado: float, subcategoria: str) -> tuple[str, str]:
    """Define a coluna Tipo usada pelo Painel e pelo artefato. Retorna (tipo, observação)."""
    texto = descricao.lower()
    if "consignado" in texto:
        return "Empréstimo", ""
    if any(k in texto for k in ("aplicação rdb", "aplicacao rdb", "resgate rdb", "cofrinho")):
        return "Investimento", ""
    if "rendimento" in texto:
        return "Receita", ""
    if subcategoria == "Fatura Cartão" or any(k in texto for k in ("pagamento de fatura",)) or texto.strip() == "pagamento recebido":
        return "Pagamento de fatura", ""
    if any(k in texto for k in ("valor adicionado", "limite convertido", "reembolso de pix no crédito", "reembolso de pix no credito")):
        return "Transferência interna", ""
    if any(n in texto for n in NOMES_CASAL) and any(k in texto for k in ("pix", "transfer")):
        return "Transferência interna", ""
    if any(k in texto for k in ("estorno", "reembolso", "devolução", "devolucao")):
        return "Despesa", "Estorno/reembolso: abate dos gastos"
    if valor_tratado > 0 and any(k in texto for k in ("salário", "salario", "vencimento", "pix recebido", "transferência recebida", "transferencia recebida", "crédito em conta", "credito em conta")):
        return "Receita", ""
    if valor_tratado < 0:
        return "Despesa", ""
    return "Revisar", "Entrada positiva sem regra: confirmar o que é"


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
    if origem == "Cartão de Crédito":
        credito = any(k in descricao.lower() for k in KEYWORDS_CREDITO_CARTAO)
        valor = round(abs(valor_raw), 2)
        valor_tratado = round(abs(valor_raw), 2) if credito else round(-abs(valor_raw), 2)
    else:
        valor = round(valor_raw, 2)
        valor_tratado = round(valor_raw, 2)

    tipo, obs = classificar_tipo(descricao, valor_tratado, subcategoria)
    # Mantém Categoria/Subcategoria coerentes com o Tipo para o que não é gasto.
    categoria, subcategoria = {
        "Receita": ("Receitas", "Entradas / Recebimentos"),
        "Transferência interna": ("Transferências internas", "Transferência interna"),
        "Investimento": ("Investimentos", "Investimentos (RDB)"),
        "Pagamento de fatura": ("Pagamento de fatura", "Fatura Cartão"),
        "Empréstimo": ("Empréstimos", "Empréstimo"),
    }.get(tipo, (categoria, subcategoria))
    return [descricao, parcela_atual, parcela_total, valor, data_br, mes_ano, categoria, origem, valor_tratado, pessoa, tx_id, subcategoria, tipo, obs]


def chave_duplicata(descricao: str, data_br: str, valor_tratado: Any, pessoa: str) -> str:
    """Chave para achar o mesmo lançamento que o Pierre devolveu com outro ID."""
    try:
        if isinstance(valor_tratado, (int, float)):
            v = f"{abs(float(valor_tratado)):.2f}"
        else:
            txt = str(valor_tratado).replace("R$", "").replace("−", "-").replace(" ", "").strip()
            if "," in txt:
                txt = txt.replace(".", "").replace(",", ".")
            v = f"{abs(float(txt or 0)):.2f}"
    except ValueError:
        v = str(valor_tratado)
    return "|".join([re.sub(r"\s+", " ", descricao.lower()).strip(), data_br.strip(), v, pessoa.strip().lower()])


def main() -> None:
    worksheet = conectar_planilha()
    
    # 1. Analisa os dados existentes para achar "Outros" e evitar duplicatas
    dados_planilha = worksheet.get_all_values()
    ids_existentes = set()
    chaves_existentes = set()
    linhas_outros_antigas = []
    
    for i, row in enumerate(dados_planilha, start=1):
        if i == 1: continue
        if len(row) > 10 and row[10].strip():
            ids_existentes.add(row[10].strip())
        if len(row) > 9:
            chaves_existentes.add(chave_duplicata(row[0], row[4], row[8], row[9]))
            
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
                chave = chave_duplicata(linha[0], linha[4], linha[8], linha[9])
                if chave in chaves_existentes:
                    # Mesmo lançamento com outro ID: entra marcado para revisão, sem contar nos totais.
                    linha[12], linha[13] = "Revisar", "Possível duplicata (mesma descrição, data, valor e pessoa)"
                chaves_existentes.add(chave)
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