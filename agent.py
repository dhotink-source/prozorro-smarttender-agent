import os
import io
import re
import sys
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

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(line_buffering=True)

# 1. Конфигурация ключей из GitHub Secrets
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "").strip()
GOOGLE_WEB_APP_URL = os.environ.get("GOOGLE_WEB_APP_URL")
SECRET_BRIDGE_KEY = os.environ.get("SECRET_BRIDGE_KEY", "FUEL_TENDER_SECRET_2026")

PROZORRO_API_BASE = "https://public-api.prozorro.gov.ua/api/2.5"

# Порог бюджета: 2,4 млн грн
MIN_AMOUNT_UAH = 2400000

# Коды ДК 021:2015: общая группа 09130000, бензины 09132..., дизельное топливо 09134...
TARGET_CPV_PREFIXES = ("09130000", "09132", "09134")
EXCLUDED_CPV_PREFIXES = ("09131", "09133", "09135")

MAX_GEMINI_ANALYSES_PER_RUN = 15

# Модели в Vertex AI (us-central1 / global)
GEMINI_MODELS_FALLBACK = (
    "gemini-2.5-flash",
    "gemini-2.0-flash",
    "gemini-2.5-pro",
)

# Жесткий фильтр исключения: керосин (гас), авиатопливо (JET A-1, РТ, ТС-1), газ, сжиженный газ, мазут, масла
EXCLUDED_FUEL_REGEX = re.compile(
    r"(\bгас\b|\bгасу\b|гасов[еоиій]|газ\b|\bгазу\b|газов[еоиій]|скраплен|пропан|бутан|метан|"
    r"авіаційн|реактивн|газотурбін|jet\s*a|джет|\bрт\b|\bтс-?1\b|мазут|олив[аи]|мастил|пічн[ео]\s+палив)",
    re.IGNORECASE
)

# Обязательный фильтр целевого топлива: бензин или дизельное топливо в любых проявлениях
MANDATORY_FUEL_REGEX = re.compile(
    r"(дизельн|паливо\s+дизельн|пальне\s+дизельн|\bдп\b|дп-|бензин|\bа-?95\b|\bа-?92\b|\bа-?98\b|\bа-?80\b)",
    re.IGNORECASE
)

# Фильтр для отсева явных талонов, топливных карт и скретч-карт
COUPON_CARD_REGEX = re.compile(
    r"(талон|скретч|смарт-?карт|паливн[іи]х?\s+карт|бланк[іи]?в?-дозвол|заправк[аи]\s+на\s+азс|відпуск\s+з\s+азс)",
    re.IGNORECASE
)
BULK_KEYWORD_REGEX = re.compile(r"(налив|автоцистерн|бензовоз|вагон-?цистерн)", re.IGNORECASE)

# Признаки ставки НДС 0% (Постановление КМУ №178 от 02.03.2022 и прямое указание ПДВ 0%)
ZERO_VAT_REGEX = re.compile(
    r"(постанов[аиуою]\s+кму\s*№?\s*178|від\s+02\.03\.2022\s*№?\s*178|пдв\s*0\s*%|0\s*%\s*пдв|"
    r"нульов[оа]ю\s+ставк[оа]ю|за\s+ставкою\s+0\s*%)",
    re.IGNORECASE
)

PROCUREMENT_METHODS = {
    "aboveThreshold": "Открытые торги с особенностями",
    "aboveThresholdUA": "Открытые торги",
    "aboveThresholdEU": "Открытые торги с публикацией на английском языке",
    "priceQuotation": "Запрос ценовых предложений",
    "negotiation": "Переговорная процедура",
    "negotiation.quick": "Переговорная процедура (ускоренная)",
    "belowThreshold": "Упрощенная закупка",
    "reporting": "Отчет о заключенном договоре"
}

# 2. HTTP-сессия с пулом соединений
http_session = requests.Session()
retries = Retry(total=3, backoff_factor=0.5, status_forcelist=[500, 502, 503, 504])
adapter = HTTPAdapter(pool_connections=20, pool_maxsize=20, max_retries=retries)
http_session.mount("https://", adapter)
http_session.mount("http://", adapter)

# 3. База данных проверенных тендеров и инициализация клиента Vertex AI
conn = sqlite3.connect("processed_tenders.db")
cursor = conn.cursor()
cursor.execute("CREATE TABLE IF NOT EXISTS tenders (id TEXT PRIMARY KEY)")
cursor.execute("CREATE TABLE IF NOT EXISTS sync_state (key TEXT PRIMARY KEY, val TEXT)")
cursor.execute("CREATE TABLE IF NOT EXISTS retry_queue (id TEXT PRIMARY KEY)")
conn.commit()

cursor.execute("DELETE FROM retry_queue WHERE id LIKE '%004773%'")
conn.commit()

if GEMINI_API_KEY.startswith("{"):
    creds_info = json.loads(GEMINI_API_KEY)
    project_id = creds_info.get("project_id", "ultra-dominion-510720-s6")
    if creds_info.get("type") == "service_account":
        from google.oauth2 import service_account
        creds = service_account.Credentials.from_service_account_info(
            creds_info,
            scopes=["https://www.googleapis.com/auth/cloud-platform"]
        )
    else:
        from google.oauth2.credentials import Credentials
        creds = Credentials.from_authorized_user_info(
            creds_info,
            scopes=["https://www.googleapis.com/auth/cloud-platform"]
        )
    client = genai.Client(
        vertexai=True,
        project=project_id,
        location="us-central1",
        credentials=creds
    )
else:
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

# 4. Проверка номенклатуры: отсев авиатоплива/керосина/газа и проверка наличия ДТ/бензина
def is_valid_diesel_or_gasoline_tender(tender: dict) -> tuple[bool, str]:
    title_and_desc = f"{tender.get('title', '')} {tender.get('description', '')}"
    items = tender.get("items", [])

    valid_fuel_items = []
    for item in items:
        cpv = str(item.get("classification", {}).get("id", ""))
        item_desc = str(item.get("description", ""))

        if cpv.startswith(EXCLUDED_CPV_PREFIXES):
            continue
        if not cpv.startswith(TARGET_CPV_PREFIXES):
            continue
        if EXCLUDED_FUEL_REGEX.search(item_desc) and not MANDATORY_FUEL_REGEX.search(item_desc):
            continue
        valid_fuel_items.append(item_desc)

    if not valid_fuel_items:
        return False, "Нет позиций бензина или ДТ по целевым кодам ДК"

    combined_items_text = " ".join(valid_fuel_items)
    full_check_text = f"{title_and_desc} {combined_items_text}"

    if EXCLUDED_FUEL_REGEX.search(full_check_text) and not MANDATORY_FUEL_REGEX.search(combined_items_text):
        return False, "Отсеян по стоп-словам (гас / газ / авиатопливо JET A-1)"

    if not MANDATORY_FUEL_REGEX.search(full_check_text):
        return False, "Не найдено упоминаний бензина или дизельного топлива"

    return True, "OK"

# 5. Умная проверка явных талонов/карт с полной защитой многолотовых закупок
def is_strictly_coupons_or_cards(tender: dict) -> bool:
    full_header = f"{tender.get('title', '')} {tender.get('description', '')}"
    items_desc = " ".join(str(i.get("description", "")) for i in tender.get("items", []))
    combined_text = f"{full_header} {items_desc}"

    if BULK_KEYWORD_REGEX.search(combined_text):
        return False

    lots = tender.get("lots", [])
    if len(lots) > 1:
        active_lots = [l for l in lots if l.get("status", "active") == "active"]
        if active_lots:
            return all(
                bool(COUPON_CARD_REGEX.search(f"{l.get('title', '')} {l.get('description', '')}"))
                for l in active_lots
            )

    return bool(COUPON_CARD_REGEX.search(full_header))

# 6. Формирование прямой ссылки на закупку на площадке SmartTender.biz
def build_smarttender_url(tender_ua_id: str) -> str:
    gateway_url = f"https://smarttender.biz/government/details/{tender_ua_id}"
    try:
        resp = http_session.head(gateway_url, allow_redirects=True, timeout=8)
        if resp.status_code == 200 and "smarttender.biz" in resp.url:
            return resp.url
    except Exception:
        pass
    return gateway_url

# 7. Форматирование даты и времени: "ЧЧ:ММ, ДД-ММ-ГГГГ"
def format_deadline_strict(iso_str: str) -> str:
    if not iso_str:
        return "Не указан"
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})T(\d{2}:\d{2})", str(iso_str))
    if m:
        year, month, day, time_hm = m.group(1), m.group(2), m.group(3), m.group(4)
        return f"{time_hm}, {day}-{month}-{year}"
    return str(iso_str)[:16].replace("T", " ")

# 8. Гарантированное указание источника при ссылке на пункты: "(п. 4.1)" -> "(п. 4.1 Договора)"
def enforce_clause_source(text: str, default_source: str = "Договора") -> str:
    if not text:
        return ""
    pattern = re.compile(r"\((п\.?\s*[\d\.,\s\-а-яіїєґА-ЯІЇЄҐ]+)\)")
    source_keywords = ("догов", "контракт", "тд", "документац", "додат", "прилож", "специф", "оголош", "закон", "постанов")

    def replacer(match):
        inner = match.group(1).strip()
        low_inner = inner.lower()
        if any(k in low_inner for k in source_keywords):
            return f"({inner})"
        return f"({inner} {default_source})"

    return pattern.sub(replacer, text)

# 9. Точный расчет бюджета с проверкой НДС (0% по Постановлению КМУ №178 или 20%)
def calculate_budget_with_vat_verified(tender_data: dict, ai: dict, docs_text: str) -> str:
    val = tender_data.get("value", {})
    budget_raw = float(val.get("amount", 0) or 0)
    prozorro_vat_included = bool(val.get("valueAddedTaxIncluded", False))

    ai_vat_rate = ai.get("vat_rate_percent", 20)
    is_zero_vat = (ai_vat_rate == 0) or bool(ZERO_VAT_REGEX.search(docs_text))

    if is_zero_vat:
        multiplier = 1.0
        vat_label = "с НДС (0%)"
    else:
        multiplier = 1.0 if prozorro_vat_included else 1.2
        vat_label = "с НДС"

    items = tender_data.get("items", [])
    fuel_items = [
        i for i in items
        if str(i.get("classification", {}).get("id", "")).startswith(TARGET_CPV_PREFIXES)
        and not str(i.get("classification", {}).get("id", "")).startswith(EXCLUDED_CPV_PREFIXES)
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
            unit_price_with_vat = (budget_raw / qty) * multiplier
            return f"{unit_price_with_vat:,.2f} грн/{unit_label} {vat_label}".replace(",", " ")

    total_with_vat = budget_raw * multiplier
    return f"{total_with_vat:,.2f} грн {vat_label}".replace(",", " ")

# 10. Чтение файлов .docx и старых файлов .doc (Word 97-2003)
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

# 11. Выгрузка актуальных редакций вложений
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
        if any(k in t for k in ("догов", "контракт", "техніч", "документац", "додаток", "специфікац", "вимог", "оголош")):
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

# 12. Когнитивная экспертиза через Gemini (Vertex AI)
def analyze_tender_with_gemini(tender_data: dict, docs_text: str, pdf_parts: list) -> dict:
    val = tender_data.get("value", {})
    budget_amount = val.get("amount", 0)
    vat_included_flag = val.get("valueAddedTaxIncluded", False)
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
    Бюджет закупки в системе Prozorro: {budget_amount} UAH (флаг valueAddedTaxIncluded в системе: {vat_included_flag})
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
            "1. Ищем ТОЛЬКО закупки БЕНЗИНА или ДИЗЕЛЬНОГО ТОПЛИВА (бензин, дизельне паливо, паливо дизельне, ДП) НАЛИВОМ: "
            "АВТОМОБИЛЬНЫМ ТРАНСПОРТОМ (бензовозами/автоцистернами в резервуары заказчика или самовывозом) ИЛИ ЖЕЛЕЗНОДОРОЖНЫМ ТРАНСПОРТОМ (вагон-цистернами).\n"
            "2. Если закупается авиационное топливо (JET A-1, РТ, ТС-1), керосин (гас, гасове паливо), газ, сжиженный газ (газ скраплений), мазут или масла — СТРОГО верни is_relevant = false!\n"
            "3. Если основной предмет закупки — талоны, топливные карты, скретч-карты, бланки-разрешения или розничная заправка на АЗС — СТРОГО верни is_relevant = false!\n"
            "4. ПРОВЕРКА НДС (vat_rate_percent): Внимательно проверь раздел 'Цена договора' в проекте договора! "
            "Если заказчик (например, Нацгвардия, ВСУ, ГПСУ, ГСЧС) закупает топливо с НДС 0% (ПДВ 0%, в т.ч. согласно Постановлению КМУ №178 от 02.03.2022) или без НДС — укажи vat_rate_percent = 0. "
            "Если применяется стандартная ставка НДС 20% — укажи vat_rate_percent = 20.\n"
            "5. ОБЯЗАТЕЛЬНОЕ ПРАВИЛО ССЫЛОК НА ПУНКТЫ: Везде указывай номер пункта и точный источник (например: '(п. 4.1 Договора)', '(п. 4 Оголошення)', '(п. 3.2 Тендерной документации)').\n"
            "Все ответы формируй на русском языке."
        ),
        response_mime_type="application/json",
        response_schema={
            "type": "OBJECT",
            "properties": {
                "is_relevant": {"type": "BOOLEAN"},
                "vat_rate_percent": {
                    "type": "INTEGER",
                    "description": "Ставка НДС по договору: 0 (если ПДВ 0% по Постановлению КМУ №178 или без ПДВ) или 20 (если стандартный НДС 20%)"
                },
                "subject_grade_and_volume": {
                    "type": "STRING",
                    "description": "Пункт 1 (Предмет закупки и объем): Полное наименование товара, марка топлива (например: Дизельное топливо ДП (Л) – Евро-5) и точное количество с единицами измерения и ссылкой на пункт Договора/Оголошення"
                },
                "payment_terms": {
                    "type": "STRING",
                    "description": "Пункт 3: Условия оплаты с обязательным указанием пункта и документа-источника, например: (п. 4.1 Договора)"
                },
                "is_formula_pricing": {
                    "type": "BOOLEAN",
                    "description": "true, если цена определяется формульным ценообразованием; false, если цена фиксированная"
                },
                "formula_details": {
                    "type": "STRING",
                    "description": "Пункт 4: Если цена формульная — точная математическая формула и описание всех коэффициентов с указанием пункта и источника. Если фиксированная — напиши 'Цена фиксированная'"
                },
                "delivery_mode_and_location": {
                    "type": "STRING",
                    "description": "Пункт 6: Способ отгрузки, базис поставки и точный адрес передачи товара с указанием пункта и источника (например: п. 5.2 Договора)"
                },
                "delivery_and_contract_term": {
                    "type": "STRING",
                    "description": "Пункт 7: Срок поставки партии, общий период поставки и срок действия договора с указанием пунктов и источника"
                },
                "bid_security_uah": {
                    "type": "STRING",
                    "description": "Пункт 8: Обеспечение тендерного предложения — размер гарантии в сумме (грн) с указанием пункта ТД, или 'Не требуется'"
                },
                "contract_security_details": {
                    "type": "STRING",
                    "description": "Пункт 9: Обеспечение выполнения договора — вид обеспечения и размер в % от суммы договора с указанием пункта Договора/ТД, или 'Не требуется'"
                },
                "features_pitfalls_and_risks": {
                    "type": "STRING",
                    "description": "Для кнопки 'Особенности тендера (Риски поставщика)' (Блок 1): Подводные камни, особенности приемки и риски убытков поставщика с номерами пунктов, источником и прямыми цитатами"
                },
                "features_penalties": {
                    "type": "STRING",
                    "description": "Для кнопки 'Особенности тендера (Риски поставщика)' (Блок 2): Ответственность поставщика (штрафы, пеня) с номерами пунктов, источником и прямыми цитатами"
                },
                "features_price_increase_rules": {
                    "type": "STRING",
                    "description": "Для кнопки 'Особенности тендера (Риски поставщика)' (Блок 3): Условия повышения цены с номерами пунктов, источником и прямыми цитатами"
                }
            },
            "required": [
                "is_relevant", "vat_rate_percent", "subject_grade_and_volume",
                "payment_terms", "is_formula_pricing", "formula_details",
                "delivery_mode_and_location", "delivery_and_contract_term",
                "bid_security_uah", "contract_security_details",
                "features_pitfalls_and_risks", "features_penalties", "features_price_increase_rules"
            ]
        },
        temperature=0.0
    )

    primary_err = None
    for model_name in GEMINI_MODELS_FALLBACK:
        for attempt in range(2):
            try:
                response = client.models.generate_content(
                    model=model_name,
                    contents=contents,
                    config=gen_config
                )
                return json.loads(response.text)
            except Exception as err:
                err_str = str(err)
                if primary_err is None and "404" not in err_str:
                    primary_err = err
                if "404" in err_str or "NOT_FOUND" in err_str:
                    break
                if "503" in err_str or "UNAVAILABLE" in err_str or "429" in err_str:
                    time.sleep(4 * (attempt + 1))
                    continue
                break
    raise (primary_err or Exception("Все модели Gemini временно недоступны"))

# 13. Сборка карточки в новом формате и передача в Google Sheets CRM
def dispatch_to_crm(tender_data: dict, ai: dict, docs_text: str):
    tender_ua_id = tender_data.get("tenderID", "Не указан")
    entity = tender_data.get("procuringEntity", {})
    cust_name = entity.get("name", "Не указан")
    cust_edrpou = str(entity.get("identifier", {}).get("id", "Не указан"))
    customer_str = f"{cust_name} (ЕГРПОУ: {cust_edrpou})"

    method_type = tender_data.get("procurementMethodType", "")
    procedure_clean = PROCUREMENT_METHODS.get(method_type, "Открытые торги")

    budget_display_str = calculate_budget_with_vat_verified(tender_data, ai, docs_text)

    raw_end_date = tender_data.get("tenderPeriod", {}).get("endDate", "")
    deadline_formatted = format_deadline_strict(raw_end_date)

    tender_url = build_smarttender_url(tender_ua_id)

    subject_vol_clean = enforce_clause_source(ai.get("subject_grade_and_volume", ""), "Договора")
    payment_terms_clean = enforce_clause_source(ai.get("payment_terms", ""), "Договора")
    delivery_loc_clean = enforce_clause_source(ai.get("delivery_mode_and_location", ""), "Договора")
    delivery_term_clean = enforce_clause_source(ai.get("delivery_and_contract_term", ""), "Договора")
    bid_sec_clean = enforce_clause_source(ai.get("bid_security_uah", ""), "Тендерной документации")
    contract_sec_clean = enforce_clause_source(ai.get("contract_security_details", ""), "Договора")
    formula_raw = enforce_clause_source(ai.get("formula_details", ""), "Договора")

    pitfalls_clean = enforce_clause_source(ai.get("features_pitfalls_and_risks", ""), "Договора")
    penalties_clean = enforce_clause_source(ai.get("features_penalties", ""), "Договора")
    price_inc_clean = enforce_clause_source(ai.get("features_price_increase_rules", ""), "Договора")

    is_formula = bool(ai.get("is_formula_pricing"))
    if is_formula:
        pricing_html = (
            f"Цена формульная (нажмите на спойлер ниже или кнопку под сообщением):\n"
            f"<tg-spoiler>{html.escape(formula_raw)}</tg-spoiler>"
        )
    else:
        pricing_html = "Цена фиксированная"

    card_html = (
        f"<b>Тендер:</b> <code>{html.escape(tender_ua_id)}</code> ({html.escape(procedure_clean)})\n"
        f"-----------------------------------------\n"
        f"<b>{html.escape(customer_str)}</b>\n\n"
        f"<b>1) Предмет закупки и объем:</b> {html.escape(subject_vol_clean)}\n\n"
        f"<b>2) Бюджет закупки:</b> {html.escape(budget_display_str)}\n\n"
        f"<b>3) Условия оплаты:</b> {html.escape(payment_terms_clean)}\n\n"
        f"<b>4) Ценообразование:</b> {pricing_html}\n\n"
        f"<b>5) Срок подачи тендерного предложения:</b> до {html.escape(deadline_formatted)}\n\n"
        f"<b>6) Условия поставки и адрес передачи товара:</b> {html.escape(delivery_loc_clean)}\n\n"
        f"<b>7) Срок поставки и срок действия договора:</b> {html.escape(delivery_term_clean)}\n\n"
        f"<b>8) Обеспечение тендерного предложения:</b> {html.escape(bid_sec_clean)}\n\n"
        f"<b>9) Обеспечение договора:</b> {html.escape(contract_sec_clean)}\n\n"
        f"<b>10) Ссылка на тендер:</b> <a href=\"{html.escape(tender_url)}\">{html.escape(tender_url)}</a>\n"
        f"-----------------------------------------\n"
        f"<i>Ответьте (Reply) свободным текстом на это сообщение, чтобы ИИ внес вашу цену, статус или результат в таблицу.</i>"
    )

    features_detailed_text = (
        f"1. ПОДВОДНЫЕ КАМНИ, ПРИЕМКА И РИСКИ УБЫТКОВ ПОСТАВЩИКА:\n"
        f"{pitfalls_clean}\n\n"
        f"2. ОТВЕТСТВЕННОСТЬ ПОСТАВЩИКА (ШТРАФНЫЕ САНКЦИИ И ПЕНЯ):\n"
        f"{penalties_clean}\n\n"
        f"3. УСЛОВИЯ И ОГРАНИЧЕНИЯ ДЛЯ ПОВЫШЕНИЯ ЦЕНЫ (ДОКУМЕНТЫ И ПУНКТЫ ДОГОВОРА):\n"
        f"{price_inc_clean}"
    )

    payload = {
        "source": "python_agent",
        "secret": SECRET_BRIDGE_KEY,
        "tender": {
            "prozorroInternalId": tender_data.get("id", tender_ua_id),
            "tenderID": tender_ua_id,
            "customerName": cust_name,
            "customerEdrpou": cust_edrpou,
            "procedureType": procedure_clean,
            "budgetDisplay": budget_display_str,
            "fuelSummary": subject_vol_clean,
            "deliveryAndTerms": f"{delivery_loc_clean} | Сроки: {delivery_term_clean}",
            "paymentAndGuarantees": (
                f"Оплата: {payment_terms_clean} | "
                f"Тендерная гарантия: {bid_sec_clean} | "
                f"Обеспечение договора: {contract_sec_clean}"
            ),
            "deadline": deadline_formatted,
            "tenderUrl": tender_url,
            "isFormula": is_formula,
            "formulaText": formula_raw,
            "featuresDetailedText": features_detailed_text,
            "telegramCardHtml": card_html
        }
    }

    resp = http_session.post(GOOGLE_WEB_APP_URL, json=payload, timeout=25)
    print(f" -> Ответ Google Таблицы: {resp.text[:120]}")

# 14. Проверка результатов торгов по тендерам «Принимаем участие»
def check_results_for_participating_tenders():
    try:
        req_payload = {
            "source": "python_agent",
            "secret": SECRET_BRIDGE_KEY,
            "action": "get_tracked_tenders"
        }
        r = http_session.post(GOOGLE_WEB_APP_URL, json=req_payload, timeout=20)
        if r.status_code != 200:
            return
        resp_data = r.json()
        tracked_list = resp_data.get("trackedTenders", [])
        if not tracked_list:
            return

        print(f"=== Проверка результатов торгов по тендерам «Принимаем участие» ({len(tracked_list)} шт.) ===")
        for item in tracked_list:
            tender_ua_id = item.get("tenderID", "")
            lookup_id = item.get("prozorroInternalId") or tender_ua_id
            already_notified_stage = item.get("notifiedStage", "")

            t_data = fetch_tender_detail(lookup_id)
            if not t_data:
                continue

            awards = t_data.get("awards", [])
            bids = t_data.get("bids", [])
            method_type = t_data.get("procurementMethodType", "")

            event_stage = ""
            event_title = ""
            winner_name = ""
            winner_edrpou = ""
            price_str = ""

            active_awards = [a for a in awards if a.get("status") == "active"]
            pending_awards = [a for a in awards if a.get("status") == "pending"]

            if active_awards:
                aw = active_awards[-1]
                supp = (aw.get("suppliers") or [{}])[0]
                winner_name = supp.get("name", "Не указано")
                winner_edrpou = str(supp.get("identifier", {}).get("id", "Не указан"))
                val = aw.get("value", {})
                amt = float(val.get("amount", 0) or 0)
                vat_inc = val.get("valueAddedTaxIncluded", False)
                price_str = f"{amt:,.2f} грн ({'с НДС' if vat_inc else 'без НДС'})".replace(",", " ")
                event_stage = f"active_award_{aw.get('id', '')}"
                event_title = "АКЦЕПТ / ПРИЗНАН ПОБЕДИТЕЛЬ ТЕНДЕРА"

            elif pending_awards:
                aw = pending_awards[-1]
                supp = (aw.get("suppliers") or [{}])[0]
                winner_name = supp.get("name", "Не указано")
                winner_edrpou = str(supp.get("identifier", {}).get("id", "Не указан"))
                val = aw.get("value", {})
                amt = float(val.get("amount", 0) or 0)
                vat_inc = val.get("valueAddedTaxIncluded", False)
                price_str = f"{amt:,.2f} грн ({'с НДС' if vat_inc else 'без НДС'})".replace(",", " ")
                event_stage = f"pending_award_{aw.get('id', '')}"
                if method_type == "priceQuotation":
                    event_title = "ЗАПРОС ЦЕНОВЫХ ПРЕДЛОЖЕНИЙ: НАИМЕНЬШАЯ ЦЕНА (НА РАССМОТРЕНИИ)"
                else:
                    event_title = "ОПРЕДЕЛЕН ЛИДЕР С НАИМЕНЬШЕЙ ЦЕНОЙ (КВАЛИФИКАЦИЯ)"

            elif bids:
                valid_bids = [
                    b for b in bids
                    if b.get("status", "active") in ("active", "pending")
                    and b.get("value", {}).get("amount") is not None
                    and b.get("tenderers")
                ]
                if valid_bids:
                    valid_bids.sort(key=lambda x: float(x["value"]["amount"]))
                    best_bid = valid_bids[0]
                    supp = best_bid["tenderers"][0]
                    winner_name = supp.get("name", "Не указано")
                    winner_edrpou = str(supp.get("identifier", {}).get("id", "Не указан"))
                    val = best_bid.get("value", {})
                    amt = float(val.get("amount", 0) or 0)
                    vat_inc = val.get("valueAddedTaxIncluded", False)
                    price_str = f"{amt:,.2f} грн ({'с НДС' if vat_inc else 'без НДС'})".replace(",", " ")
                    event_stage = f"lowest_bid_{best_bid.get('id', '')}"
                    event_title = "РАСКРЫТЫ ПРЕДЛОЖЕНИЯ: УЧАСТНИК С САМОЙ НИЗКОЙ ЦЕНОЙ"

            if event_stage and event_stage != already_notified_stage:
                notify_payload = {
                    "source": "python_agent",
                    "secret": SECRET_BRIDGE_KEY,
                    "action": "notify_tender_result",
                    "tenderID": tender_ua_id,
                    "eventStage": event_stage,
                    "eventTitle": event_title,
                    "winnerName": winner_name,
                    "winnerEdrpou": winner_edrpou,
                    "winningPrice": price_str
                }
                nr = http_session.post(GOOGLE_WEB_APP_URL, json=notify_payload, timeout=20)
                print(f" -> Отправлено уведомление о результатах {tender_ua_id} ({winner_name}): {nr.text[:100]}")
    except Exception as e:
        print(f"Ошибка при проверке результатов активных тендеров: {e}")

# 15. Загрузка карточки тендера с защитой и автоповтором при коде 429
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

# 16. Оценка активных тендеров
def evaluate_active_tenders(active_ids: list, gemini_budget_left: int) -> tuple[int, int, int, int]:
    matched_budget = 0
    matched_fuel_cpv = 0
    sent_to_crm = 0
    gemini_used = 0
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

            is_valid_fuel, _ = is_valid_diesel_or_gasoline_tender(tender)
            if not is_valid_fuel:
                mark_as_processed(tid)
                continue

            matched_fuel_cpv += 1
            fuel_tenders_to_analyze.append(tender)

    for tender in fuel_tenders_to_analyze:
        t_id = tender["id"]
        tender_ua_id = tender.get("tenderID", t_id)
        amount = float(tender.get("value", {}).get("amount", 0) or 0)
        print(f"Целевой тендер (Бензин/ДТ >= 2.4 млн): {tender_ua_id} | {amount:,.2f} грн | {tender.get('title', '')[:70]}")

        if is_strictly_coupons_or_cards(tender):
            print(f" -> Мгновенно отсеян (явные талоны/топливные карты без налива): {tender_ua_id}")
            mark_as_processed(t_id)
            mark_as_processed(tender_ua_id)
            continue

        if gemini_used >= gemini_budget_left:
            print(f" -> Достигнут лимит ИИ на 1 цикл ({MAX_GEMINI_ANALYSES_PER_RUN} шт.), {tender_ua_id} проверим в следующем цикле.")
            add_to_retry_queue(t_id)
            continue

        try:
            if gemini_used > 0:
                time.sleep(2)
            print(f" -> Скачивание документации и экспертиза Gemini (Vertex AI) для {tender_ua_id}...")
            docs_text, pdf_parts = extract_all_attachments(tender.get("documents", []))
            ai_verdict = analyze_tender_with_gemini(tender, docs_text, pdf_parts)
            gemini_used += 1

            if ai_verdict.get("is_relevant"):
                print(f" -> ПОДТВЕРЖДЕН НАЛИВ! Отправка в Google Sheets и Telegram: {tender_ua_id}")
                dispatch_to_crm(tender, ai_verdict, docs_text)
                sent_to_crm += 1
            else:
                print(f" -> Отклонен ИИ после чтения ТД (талоны/карты/АЗС/не ДТ и не бензин): {tender_ua_id}")

            mark_as_processed(t_id)
            mark_as_processed(tender_ua_id)
        except Exception as e:
            gemini_used += 1
            print(f" -> Ошибка при анализе {tender_ua_id} (сохранен в очередь retry_queue): {e}")
            add_to_retry_queue(t_id)

    return matched_budget, matched_fuel_cpv, sent_to_crm, gemini_used

# 17. Главный цикл мониторинга
def run_monitoring():
    total_scanned = 0
    total_active = 0
    total_budget = 0
    total_fuel = 0
    total_sent = 0
    gemini_left = MAX_GEMINI_ANALYSES_PER_RUN

    retry_ids = get_retry_queue_ids()
    if retry_ids:
        print(f"=== Проверка отложенных тендеров из очереди retry_queue ({len(retry_ids)} шт.) ===")
        rb, rf, rs, g_used = evaluate_active_tenders(retry_ids, gemini_left)
        gemini_left -= g_used
        total_active += len(retry_ids)
        total_budget += rb
        total_fuel += rf
        total_sent += rs

    last_top_ts = get_sync_val("last_top_ts")
    new_top_ts = ""
    next_url = f"{PROZORRO_API_BASE}/tenders?descending=1&limit=100&opt_fields=status"
    
    fresh_active_ids = []
    max_head_pages = 15

    print("=== Проверка свежих изменений ленты Prozorro ===")
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

    if new_top_ts:
        set_sync_val("last_top_ts", new_top_ts)

    print(f"Найдено свежих активных тендеров: {len(fresh_active_ids)}")
    b1, f1, s1, _ = evaluate_active_tenders(fresh_active_ids, gemini_left)
    total_active += len(fresh_active_ids)
    total_budget += b1
    total_fuel += f1
    total_sent += s1

    check_results_for_participating_tenders()

    remaining_in_queue = len(get_retry_queue_ids())
    print("=========================================")
    print("ИТОГИ ЦИКЛА МОНИТОРИНГА:")
    print(f"- Всего просмотрено записей в ленте: {total_scanned}")
    print(f"- Из них проверено активных карточек: {total_active}")
    print(f"- Активных тендеров с бюджетом >= 2.4 млн грн: {total_budget}")
    print(f"- Из них целевых (бензин / ДТ): {total_fuel}")
    print(f"- Подтверждено наливом и отправлено в Telegram: {total_sent}")
    print(f"- Осталось в очереди retry_queue на следующий цикл: {remaining_in_queue}")
    print("=========================================")

if __name__ == "__main__":
    run_monitoring()
