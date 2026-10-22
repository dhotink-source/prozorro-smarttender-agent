import os
import io
import re
import time
import html
import json
import zipfile
import sqlite3
import subprocess
import tempfile
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
import docx
from concurrent.futures import ThreadPoolExecutor, as_completed
from google import genai
from google.genai import types

# 1. Конфигурация ключей из GitHub Secrets
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY")
GOOGLE_WEB_APP_URL = os.environ.get("GOOGLE_WEB_APP_URL")
SECRET_BRIDGE_KEY = os.environ.get("SECRET_BRIDGE_KEY", "FUEL_TENDER_SECRET_2026")

PROZORRO_API_BASE = "https://public-api.prozorro.gov.ua/api/2.5"

# Порог бюджета: 2,4 млн грн (2 400 000 грн)
MIN_AMOUNT_UAH = 2400000
TARGET_CPV_PREFIXES = ("09130000", "09132000", "09134200")

# Дата начала глубокого ретроспективного поиска (20 октября 2026)
BACKFILL_TARGET_ISO = "2026-10-20T00:00:00"

GEMINI_MODELS_FALLBACK = (
    "gemini-3.0-flash",
    "gemini-3-flash-preview",
    "gemini-flash-latest",
    "gemini-2.5-pro",
)

PROCUREMENT_METHODS = {
    "aboveThreshold": "Открытые торги с особенностями",
    "aboveThresholdUA": "Открытые торги",
    "aboveThresholdEU": "Открытые торги с публикацией на английском языке",
    "priceQuotation": "Запрос ценовых предложений",
    "negotiation": "Переговорная процедура",
    "negotiation.quick": "Переговорная процедура (ускоренная)",
    "belowThreshold": "Упрощенная / допороговая закупка"
}

# 2. HTTP-сессия с пулом соединений
http_session = requests.Session()
retries = Retry(total=3, backoff_factor=0.5, status_forcelist=[500, 502, 503, 504])
adapter = HTTPAdapter(pool_connections=20, pool_maxsize=20, max_retries=retries)
http_session.mount("https://", adapter)
http_session.mount("http://", adapter)

# 3. База данных проверенных тендеров, курсоров и очереди гарантированного повтора (retry_queue)
conn = sqlite3.connect("processed_tenders.db")
cursor = conn.cursor()
cursor.execute("CREATE TABLE IF NOT EXISTS tenders (id TEXT PRIMARY KEY)")
cursor.execute("CREATE TABLE IF NOT EXISTS sync_state (key TEXT PRIMARY KEY, val TEXT)")
cursor.execute("CREATE TABLE IF NOT EXISTS retry_queue (id TEXT PRIMARY KEY)")
conn.commit()

client = genai.Client(api_key=GEMINI_API_KEY)

def is_already_processed(tender_id: str) -> bool:
    cursor.execute("SELECT 1 FROM tenders WHERE id = ?", (tender_id,))
    return cursor.fetchone() is not None

def mark_as_processed(tender_id: str):
    cursor.execute("INSERT OR IGNORE INTO tenders (id) VALUES (?)", (tender_id,))
    cursor.execute("DELETE FROM retry_queue WHERE id = ?", (tender_id,))
    conn.commit()

def add_to_retry_queue(tender_id: str):
    cursor.execute("INSERT OR IGNORE INTO retry_queue (id) VALUES (?)", (tender_id,))
    conn.commit()

def get_retry_queue_ids() -> list:
    cursor.execute("SELECT id FROM retry_queue")
    return [r[0] for r in cursor.fetchall()]

def get_sync_val(key: str) -> str:
    cursor.execute("SELECT val FROM sync_state WHERE key = ?", (key,))
    row = cursor.fetchone()
    return row[0] if row else ""

def set_sync_val(key: str, val: str):
    cursor.execute("INSERT OR REPLACE INTO sync_state (key, val) VALUES (?, ?)", (key, val))
    conn.commit()

# 4. Формирование прямой ссылки на закупку на площадке SmartTender.biz
def build_smarttender_url(tender_ua_id: str) -> str:
    gateway_url = f"https://smarttender.biz/government/details/{tender_ua_id}"
    try:
        resp = http_session.head(gateway_url, allow_redirects=True, timeout=8)
        if resp.status_code == 200 and "smarttender.biz" in resp.url:
            return resp.url
    except Exception:
        pass
    return gateway_url

# 5. Строгий математический расчет для Пункта 4 (ТОЛЬКО с НДС по формуле * 1.2)
def calculate_budget_with_vat_strict(tender_data: dict) -> str:
    budget_raw = float(tender_data.get("value", {}).get("amount", 0) or 0)
    items = tender_data.get("items", [])
    
    fuel_items = [
        i for i in items
        if str(i.get("classification", {}).get("id", "")).startswith(TARGET_CPV_PREFIXES)
    ]
    target_items = fuel_items if fuel_items else items

    if len(target_items) == 1:
        qty = float(target_items[0].get("quantity", 0) or 0)
        unit_raw = str(target_items[0].get("unit", {}).get("name", "ед.")).lower()
        if "літр" in unit_raw or "литр" in unit_raw or unit_raw == "л":
            unit_label = "литр"
        elif "тон" in unit_raw or unit_raw == "т":
            unit_label = "тонну"
        elif "кілогр" in unit_raw or "килогр" in unit_raw or unit_raw == "кг":
            unit_label = "кг"
        else:
            unit_label = unit_raw

        if qty > 0:
            unit_price_with_vat = (budget_raw / qty) * 1.2
            return f"{unit_price_with_vat:,.2f} грн/{unit_label} с НДС".replace(",", " ")

    total_with_vat = budget_raw * 1.2
    return f"{total_with_vat:,.2f} грн с НДС".replace(",", " ")

# 6. Чтение файлов .docx и старых файлов .doc (Word 97-2003)
def parse_docx_bytes(file_bytes: bytes, filename: str) -> str:
    try:
        doc = docx.Document(io.BytesIO(file_bytes))
        lines = [f"=== ДОКУМЕНТ: {filename} ==="]
        for p in doc.paragraphs:
            txt = p.text.strip()
            if txt:
                lines.append(txt)
        for table in doc.tables:
            lines.append("--- Таблица ---")
            for row in table.rows:
                cells = []
                last_txt = None
                for cell in row.cells:
                    c_txt = " ".join(cell.text.split())
                    if c_txt and c_txt != last_txt:
                        cells.append(c_txt)
                        last_txt = c_txt
                if cells:
                    lines.append(" | ".join(cells))
        return "\n".join(lines)[:45000]
    except Exception as e:
        return f"Ошибка чтения DOCX {filename}: {e}"

def parse_old_doc_bytes(file_bytes: bytes, filename: str) -> str:
    try:
        with tempfile.NamedTemporaryFile(suffix=".doc", delete=True) as tmp:
            tmp.write(file_bytes)
            tmp.flush()
            res = subprocess.run(["antiword", tmp.name], capture_output=True, text=True, timeout=10)
            if res.returncode == 0 and res.stdout.strip():
                return f"=== ДОКУМЕНТ (DOC): {filename} ===\n" + res.stdout[:45000]
    except Exception:
        pass
    return ""

# 7. Выгрузка актуальных редакций всех вложений (.docx, .doc, .pdf, .zip)
def extract_all_attachments(documents_json: list) -> tuple[str, list]:
    latest_docs = {}
    for d in documents_json:
        if d.get("documentOf") not in ("tender", "lot", "item"):
            continue
        doc_id = d.get("id")
        if not doc_id:
            continue
        if doc_id not in latest_docs or d.get("dateModified", "") > latest_docs[doc_id].get("dateModified", ""):
            latest_docs[doc_id] = d

    valid_docs = []
    for d in latest_docs.values():
        title = d.get("title", "").lower()
        if title.endswith((".p7s", ".asic", ".sig", ".asice")):
            continue
        valid_docs.append(d)

    def doc_priority(doc_item):
        t = doc_item.get("title", "").lower()
        if any(k in t for k in ("догов", "контракт", "техніч", "документац", "додаток", "специфікац", "вимог")):
            return 0
        return 1

    valid_docs.sort(key=doc_priority)

    extracted_texts = []
    pdf_parts = []

    for d in valid_docs[:8]:
        title = d.get("title", "document")
        url = d.get("url", "")
        if not url:
            continue

        try:
            resp = http_session.get(url, timeout=25)
            if resp.status_code != 200:
                continue
            content = resp.content
            low_title = title.lower()

            if low_title.endswith(".docx"):
                extracted_texts.append(parse_docx_bytes(content, title))
            elif low_title.endswith(".doc"):
                doc_txt = parse_old_doc_bytes(content, title)
                if doc_txt:
                    extracted_texts.append(doc_txt)
            elif low_title.endswith(".pdf"):
                if len(content) < 15 * 1024 * 1024 and len(pdf_parts) < 3:
                    pdf_parts.append(types.Part.from_bytes(data=content, mime_type="application/pdf"))
                    extracted_texts.append(f"[Прикреплен PDF-документ: {title}]")
            elif low_title.endswith(".zip"):
                with zipfile.ZipFile(io.BytesIO(content)) as zf:
                    for inner_name in zf.namelist()[:8]:
                        low_inner = inner_name.lower()
                        if low_inner.endswith((".p7s", ".sig")):
                            continue
                        inner_bytes = zf.read(inner_name)
                        if low_inner.endswith(".docx"):
                            extracted_texts.append(parse_docx_bytes(inner_bytes, f"{title}/{inner_name}"))
                        elif low_inner.endswith(".doc"):
                            doc_txt = parse_old_doc_bytes(inner_bytes, f"{title}/{inner_name}")
                            if doc_txt:
                                extracted_texts.append(doc_txt)
                        elif low_inner.endswith(".pdf") and len(inner_bytes) < 15 * 1024 * 1024 and len(pdf_parts) < 3:
                            pdf_parts.append(types.Part.from_bytes(data=inner_bytes, mime_type="application/pdf"))
                            extracted_texts.append(f"[Прикреплен PDF из архива: {inner_name}]")
        except Exception as e:
            extracted_texts.append(f"Ошибка загрузки файла {title}: {e}")

    return "\n\n".join(extracted_texts), pdf_parts

# 8. Когнитивная экспертиза через Gemini 3.0 Flash (с защитой от 503 UNAVAILABLE и 429)
def analyze_tender_with_gemini(tender_data: dict, docs_text: str, pdf_parts: list) -> dict:
    val = tender_data.get("value", {})
    budget_amount = val.get("amount", 0)
    guarantee_amount = tender_data.get("guarantee", {}).get("amount", 0)
    method_type = tender_data.get("procurementMethodType", "")
    method_readable = PROCUREMENT_METHODS.get(method_type, "Открытые торги")
    
    items_json = json.dumps(tender_data.get("items", []), ensure_ascii=False)
    lots_json = json.dumps(tender_data.get("lots", []), ensure_ascii=False)
    milestones_json = json.dumps(tender_data.get("milestones", []), ensure_ascii=False)

    prompt = f"""
    ПАСПОРТ ЗАКУПКИ ИЗ PROZORRO:
    ID закупки: {tender_data.get('tenderID', '')}
    Заголовок: {tender_data.get('title', '')}
    Описание: {tender_data.get('description', '')}
    Вид процедуры: {method_readable}
    Бюджет закупки в системе: {budget_amount} UAH
    Тендерное обеспечение в паспорте Prozorro: {guarantee_amount} UAH
    Лоты: {lots_json}
    Номенклатура и объемы (items): {items_json}
    Электронные условия оплаты и поставки (milestones): {milestones_json}

    ПОЛНЫЙ ТЕКСТ ТЕНДЕРНОЙ ДОКУМЕНТАЦИИ И ПРОЕКТА ДОГОВОРА:
    {docs_text[:95000]}
    """

    contents = [prompt] + pdf_parts
    gen_config = types.GenerateContentConfig(
        system_instruction=(
            "Ты — ведущий тендерный юрист и аналитик оптового поставщика нефтепродуктов в Украине. "
            "Проанализируй паспорт закупки, поля milestones/items/lots и прикрепленный проект договора/ТД.\n"
            "КРИТЕРИЙ ОТБОРА (is_relevant = true):\n"
            "1. Основной предмет закупки (или хотя бы одного отдельного лота в многолотовой закупке) — поставка дизельного топлива или бензинов НАЛИВОМ: "
            "АВТОМОБИЛЬНЫМ ТРАНСПОРТОМ (бензовозами/автоцистернами в резервуары заказчика или самовывозом) ИЛИ ЖЕЛЕЗНОДОРОЖНЫМ ТРАНСПОРТОМ (вагон-цистернами).\n"
            "2. ВАЖНО: Если основной предмет закупки — талоны, топливные карты, скретч-карты, бланки-разрешения или требуется наличие сети АЗС для заправки машин "
            "(даже если в договоре есть резервный пункт о наливе на случай остановки работы АЗС) — строго верни is_relevant = false!\n"
            "3. В поле subject_and_procedure НЕ указывай '(Постановление КМУ №1178)' — пиши только предмет закупки и чистый вид процедуры.\n"
            "Все ответы формируй на русском языке (прямые цитаты из украинского договора приводи дословно в кавычках с указанием номера пункта договора)."
        ),
        response_mime_type="application/json",
        response_schema={
            "type": "OBJECT",
            "properties": {
                "is_relevant": {"type": "BOOLEAN"},
                "subject_and_procedure": {
                    "type": "STRING",
                    "description": "Пункт 2: Предмет закупки и вид торгов БЕЗ упоминания Постановления КМУ №1178"
                },
                "fuel_grade_and_volume": {
                    "type": "STRING",
                    "description": "Пункт 3: Марка топлива и точный объем с единицами измерения (тонны / кг / литры)"
                },
                "payment_terms": {
                    "type": "STRING",
                    "description": "Пункт 5: Условия оплаты (отсрочка в банковских/календарных днях, по факту поставки или предоплата)"
                },
                "is_formula_pricing": {
                    "type": "BOOLEAN",
                    "description": "true, если цена определяется формульным ценообразованием; false, если цена фиксированная"
                },
                "formula_details": {
                    "type": "STRING",
                    "description": "Пункт 6: Если цена формульная — точная математическая формула и описание всех ее коэффициентов согласно договору. Если фиксированная — напиши 'Цена фиксированная'"
                },
                "delivery_mode_and_location": {
                    "type": "STRING",
                    "description": "Пункт 8: Способ отгрузки (Автотранспорт / ЖД вагон-цистерны), базис Инкотермс и точный адрес передачи товара или ж/д станция назначения"
                },
                "delivery_and_contract_term": {
                    "type": "STRING",
                    "description": "Пункт 9: Срок поставки отдельной партии по заявке, общий период поставки и срок действия договора"
                },
                "bid_security_uah": {
                    "type": "STRING",
                    "description": "Пункт 10: Обеспечение тендерного предложения — размер гарантии в сумме (грн), или 'Не требуется'"
                },
                "contract_security_details": {
                    "type": "STRING",
                    "description": "Пункт 11: Обеспечение выполнения договора — вид обеспечения (гарантия / депозит на счет заказчика) и размер в % от суммы договора, или 'Не требуется'"
                },
                "features_pitfalls_and_risks": {
                    "type": "STRING",
                    "description": "Для кнопки 'Особенности тендера' (Блок 1): Подводные камни по договору, особенности приемки и риски убытков поставщика с номерами пунктов и прямыми цитатами"
                },
                "features_penalties": {
                    "type": "STRING",
                    "description": "Для кнопки 'Особенности тендера' (Блок 2): Ответственность поставщика (штрафные санкции, пеня) — размер и за что вменяется, с номерами пунктов договора и прямыми цитатами"
                },
                "features_price_increase_rules": {
                    "type": "STRING",
                    "description": "Для кнопки 'Особенности тендера' (Блок 3): Условия повышения цены (на основании каких документов можно поднять цену, условия и ограничения) с номерами пунктов договора и прямыми цитатами"
                }
            },
            "required": [
                "is_relevant", "subject_and_procedure", "fuel_grade_and_volume",
                "payment_terms", "is_formula_pricing", "formula_details",
                "delivery_mode_and_location", "delivery_and_contract_term",
                "bid_security_uah", "contract_security_details",
                "features_pitfalls_and_risks", "features_penalties", "features_price_increase_rules"
            ]
        },
        temperature=0.0
    )

    last_err = None
    for model_name in GEMINI_MODELS_FALLBACK:
        for attempt in range(3):
            try:
                response = client.models.generate_content(
                    model=model_name,
                    contents=contents,
                    config=gen_config
                )
                return json.loads(response.text)
            except Exception as err:
                last_err = err
                err_str = str(err)
                # Если модель не найдена (404) — сразу переключаемся на следующую в списке
                if "404" in err_str or "NOT_FOUND" in err_str:
                    break
                # Если временный пик нагрузки (503 UNAVAILABLE или 429) — ждем 4 секунды и повторяем
                if "503" in err_str or "UNAVAILABLE" in err_str or "429" in err_str or "RESOURCE_EXHAUSTED" in err_str:
                    time.sleep(4 * (attempt + 1))
                    continue
                break
    raise last_err

# 9. Сборка карточки из 12 пунктов и передача в Google Sheets CRM (17 колонок)
def dispatch_to_crm(tender_data: dict, ai: dict):
    tender_ua_id = tender_data.get("tenderID", "Не указан")
    entity = tender_data.get("procuringEntity", {})
    cust_name = entity.get("name", "Не указан")
    cust_edrpou = str(entity.get("identifier", {}).get("id", "Не указан"))
    customer_str = f"{cust_name} (ЕГРПОУ: {cust_edrpou})"

    method_type = tender_data.get("procurementMethodType", "")
    procedure_clean = PROCUREMENT_METHODS.get(method_type, "Открытые торги")

    subject_and_proc = re.sub(
        r"\s*\(?Постановление КМУ №\s*1178\)?",
        "",
        ai.get("subject_and_procedure", ""),
        flags=re.IGNORECASE
    ).strip()

    budget_display_str = calculate_budget_with_vat_strict(tender_data)

    deadline = tender_data.get("tenderPeriod", {}).get("endDate", "Не указан")[:16].replace("T", " ")
    tender_url = build_smarttender_url(tender_ua_id)

    is_formula = bool(ai.get("is_formula_pricing"))
    formula_raw = ai.get("formula_details", "")
    if is_formula:
        pricing_html = (
            f"Цена формульная (нажмите на спойлер ниже или кнопку под сообщением):\n"
            f"<tg-spoiler>{html.escape(formula_raw)}</tg-spoiler>"
        )
    else:
        pricing_html = "Цена фиксированная"

    card_html = (
        f"<b>НОВЫЙ ТЕНДЕР (ПОСТАВКА НАЛИВОМ: АВТО / ЖД)</b>\n"
        f"<b>ID закупки:</b> <code>{html.escape(tender_ua_id)}</code>\n"
        f"-----------------------------------------\n"
        f"<b>1) Заказчик:</b> {html.escape(customer_str)}\n\n"
        f"<b>2) Предмет закупки и вид торгов:</b> {html.escape(subject_and_proc)}\n\n"
        f"<b>3) Марка топлива и объем:</b> {html.escape(ai['fuel_grade_and_volume'])}\n\n"
        f"<b>4) Бюджет закупки:</b> {html.escape(budget_display_str)}\n\n"
        f"<b>5) Условия оплаты:</b> {html.escape(ai['payment_terms'])}\n\n"
        f"<b>6) Ценообразование:</b> {pricing_html}\n\n"
        f"<b>7) Срок подачи тендерного предложения:</b> до {html.escape(deadline)}\n\n"
        f"<b>8) Условия поставки и адрес передачи товара (Авто / ЖД):</b> {html.escape(ai['delivery_mode_and_location'])}\n\n"
        f"<b>9) Срок поставки и срок действия договора:</b> {html.escape(ai['delivery_and_contract_term'])}\n\n"
        f"<b>10) Обеспечение тендерного предложения:</b> {html.escape(ai['bid_security_uah'])}\n\n"
        f"<b>11) Обеспечение договора:</b> {html.escape(ai['contract_security_details'])}\n\n"
        f"<b>12) Ссылка на тендер:</b> <a href=\"{html.escape(tender_url)}\">{html.escape(tender_url)}</a>\n"
        f"-----------------------------------------\n"
        f"<i>Ответьте (Reply) свободным текстом на это сообщение, чтобы ИИ внес вашу цену, статус или результат в таблицу.</i>"
    )

    features_detailed_text = (
        f"1. ПОДВОДНЫЕ КАМНИ, ПРИЕМКА И РИСКИ УБЫТКОВ ПОСТАВЩИКА:\n"
        f"{ai['features_pitfalls_and_risks']}\n\n"
        f"2. ОТВЕТСТВЕННОСТЬ ПОСТАВЩИКА (ШТРАФНЫЕ САНКЦИИ И ПЕНЯ):\n"
        f"{ai['features_penalties']}\n\n"
        f"3. УСЛОВИЯ И ОГРАНИЧЕНИЯ ДЛЯ ПОВЫШЕНИЯ ЦЕНЫ (ДОКУМЕНТЫ И ПУНКТЫ ДОГОВОРА):\n"
        f"{ai['features_price_increase_rules']}"
    )

    payload = {
        "source": "python_agent",
        "secret": SECRET_BRIDGE_KEY,
        "tender": {
            "tenderID": tender_ua_id,
            "customerName": cust_name,
            "customerEdrpou": cust_edrpou,
            "procedureType": procedure_clean,
            "budgetDisplay": budget_display_str,
            "fuelSummary": ai["fuel_grade_and_volume"],
            "deliveryAndTerms": f"{ai['delivery_mode_and_location']} | Сроки: {ai['delivery_and_contract_term']}",
            "paymentAndGuarantees": (
                f"Оплата: {ai['payment_terms']} | "
                f"Тендерная гарантия: {ai['bid_security_uah']} | "
                f"Обеспечение договора: {ai['contract_security_details']}"
            ),
            "deadline": deadline,
            "tenderUrl": tender_url,
            "isFormula": is_formula,
            "formulaText": formula_raw,
            "featuresDetailedText": features_detailed_text,
            "telegramCardHtml": card_html
        }
    }

    resp = http_session.post(GOOGLE_WEB_APP_URL, json=payload, timeout=25)
    print(f" -> Ответ Google Таблицы: {resp.text[:120]}")

# 10. Загрузка карточки тендера с защитой и автоповтором при коде 429
def fetch_tender_detail(t_id: str) -> dict:
    for attempt in range(4):
        try:
            r = http_session.get(f"{PROZORRO_API_BASE}/tenders/{t_id}", timeout=12)
            if r.status_code == 200:
                return r.json().get("data", {})
            elif r.status_code == 429:
                time.sleep(1.5 * (attempt + 1))
                continue
        except Exception:
            pass
    return {}

# 11. Оценка активных тендеров (max_workers=15 + сохранение в retry_queue при временных сбоях)
def evaluate_active_tenders(active_ids: list) -> tuple[int, int, int]:
    matched_budget = 0
    matched_fuel_cpv = 0
    sent_to_crm = 0
    fuel_tenders_to_analyze = []

    with ThreadPoolExecutor(max_workers=15) as executor:
        future_to_id = {executor.submit(fetch_tender_detail, tid): tid for tid in active_ids}
        for future in as_completed(future_to_id):
            tid = future_to_id[future]
            tender = future.result()
            if not tender:
                add_to_retry_queue(tid)
                continue

            status = tender.get("status", "")
            if status not in ("active.enquiries", "active.tendering"):
                mark_as_processed(tid)
                continue

            amount = float(tender.get("value", {}).get("amount", 0) or 0)
            if amount < MIN_AMOUNT_UAH:
                mark_as_processed(tid)
                continue

            matched_budget += 1

            items = tender.get("items", [])
            has_fuel_cpv = any(
                str(i.get("classification", {}).get("id", "")).startswith(TARGET_CPV_PREFIXES)
                for i in items
            )
            if not has_fuel_cpv:
                mark_as_processed(tid)
                continue

            matched_fuel_cpv += 1
            fuel_tenders_to_analyze.append(tender)

    for tender in fuel_tenders_to_analyze:
        t_id = tender["id"]
        tender_ua_id = tender.get("tenderID", t_id)
        amount = float(tender.get("value", {}).get("amount", 0) or 0)
        print(f"Топливный тендер >= 2.4 млн: {tender_ua_id} | {amount:,.2f} грн | {tender.get('title', '')[:70]}")

        try:
            print(f" -> Скачивание документации и экспертиза Gemini 3.0 Flash для {tender_ua_id}...")
            docs_text, pdf_parts = extract_all_attachments(tender.get("documents", []))
            ai_verdict = analyze_tender_with_gemini(tender, docs_text, pdf_parts)

            if ai_verdict.get("is_relevant"):
                print(f" -> ПОДТВЕРЖДЕН НАЛИВ! Отправка в Google Sheets и Telegram: {tender_ua_id}")
                dispatch_to_crm(tender, ai_verdict)
                sent_to_crm += 1
            else:
                print(f" -> Отклонен ИИ после чтения ТД (талоны/карты/АЗС): {tender_ua_id}")

            mark_as_processed(t_id)
            mark_as_processed(tender_ua_id)
        except Exception as e:
            print(f" -> Ошибка при анализе {tender_ua_id} (добавлен в очередь гарантированного повтора): {e}")
            add_to_retry_queue(t_id)

    return matched_budget, matched_fuel_cpv, sent_to_crm

# 12. Одноразовая точечная допроверка дневного среза 20 октября (10:15 - 20:00), где в запуске #7 был лимит 429
def recover_oct20_daytime_slice_if_needed() -> tuple[int, int, int, int]:
    if get_sync_val("oct20_daytime_recovered") == "1":
        return 0, 0, 0, 0

    print("=== Допроверка дневного среза 20.10.2026 (включая UA-2026-10-20-011345-a и UA-2026-10-20-000609-a) ===")
    # Стартуем с утра 20 октября (00:00) вверх по ленте до вечера 20 октября (20:30)
    url = f"{PROZORRO_API_BASE}/tenders?offset=2026-10-20T00:00:00+03:00&limit=100&opt_fields=status"
    candidate_ids = []
    scanned = 0

    for _ in range(220):
        if not url:
            break
        resp = http_session.get(url, timeout=15).json()
        data = resp.get("data", [])
        if not data:
            break
        last_ts = data[-1].get("dateModified", "")
        for item in data:
            scanned += 1
            tid = item.get("id")
            st = item.get("status", "")
            if st in ("active.enquiries", "active.tendering") and not is_already_processed(tid):
                candidate_ids.append(tid)
        if last_ts > "2026-10-20T20:30:00":
            break
        url = resp.get("next_page", {}).get("uri")

    print(f"Найдено непроверенных активных карточек за 20.10.2026: {len(candidate_ids)}")
    b, f, s = evaluate_active_tenders(candidate_ids)
    set_sync_val("oct20_daytime_recovered", "1")
    return scanned, len(candidate_ids), b, s

# 13. Главный цикл непрерывного мониторинга
def run_monitoring():
    total_scanned = 0
    total_active = 0
    total_budget = 0
    total_fuel = 0
    total_sent = 0

    # Шаг 0: Проверка очереди отложенных из-за ошибки 503 тендеров
    retry_ids = get_retry_queue_ids()
    if retry_ids:
        print(f"=== Повторная проверка отложенных тендеров из очереди retry_queue ({len(retry_ids)} шт.) ===")
        rb, rf, rs = evaluate_active_tenders(retry_ids)
        total_active += len(retry_ids)
        total_budget += rb
        total_fuel += rf
        total_sent += rs

    # Шаг 1: Проверка свежих изменений с вершины ленты Prozorro
    last_top_ts = get_sync_val("last_top_ts")
    new_top_ts = ""
    next_url = f"{PROZORRO_API_BASE}/tenders?descending=1&limit=100&opt_fields=status"
    
    fresh_active_ids = []
    max_head_pages = 60

    print("=== ФАЗА 1: Проверка свежих изменений ленты Prozorro ===")
    for page in range(1, max_head_pages + 1):
        if not next_url:
            break
        resp = http_session.get(next_url, timeout=15).json()
        items_list = resp.get("data", [])
        if not items_list:
            break

        if page == 1 and items_list:
            new_top_ts = items_list[0].get("dateModified", "")

        reached_previous_window = False
        for item in items_list:
            total_scanned += 1
            t_id = item.get("id")
            status = item.get("status", "")
            mod_ts = item.get("dateModified", "")

            if status in ("active.enquiries", "active.tendering") and not is_already_processed(t_id):
                fresh_active_ids.append(t_id)

            if last_top_ts and mod_ts < last_top_ts:
                reached_previous_window = True

        next_url = resp.get("next_page", {}).get("uri")
        if reached_previous_window:
            print(f"Окно синхронизации сомкнуто со штампом прошлого запуска ({last_top_ts}) на стр. {page}.")
            break
        if not last_top_ts and page >= 25:
            if not get_sync_val("backfill_next_url") and next_url:
                set_sync_val("backfill_next_url", next_url)
            break

    if new_top_ts:
        set_sync_val("last_top_ts", new_top_ts)

    print(f"Найдено свежих активных тендеров: {len(fresh_active_ids)}")
    b1, f1, s1 = evaluate_active_tenders(fresh_active_ids)
    total_active += len(fresh_active_ids)
    total_budget += b1
    total_fuel += f1
    total_sent += s1

    # Шаг 2: Одноразовая допроверка карточек за 20 октября (включая UA-2026-10-20-000609-a и КП "КИЇВПАСТРАНС")
    sc2, ac2, b2, s2 = recover_oct20_daytime_slice_if_needed()
    total_scanned += sc2
    total_active += ac2
    total_budget += b2
    total_sent += s2

    print("=========================================")
    print("ИТОГИ ЦИКЛА МОНИТОРИНГА:")
    print(f"- Всего просмотрено записей в ленте: {total_scanned}")
    print(f"- Из них проверено активных карточек: {total_active}")
    print(f"- Активных тендеров с бюджетом >= 2.4 млн грн: {total_budget}")
    print(f"- Подтверждено наливом и отправлено в Telegram: {total_sent}")
    print("=========================================")

if __name__ == "__main__":
    run_monitoring()
