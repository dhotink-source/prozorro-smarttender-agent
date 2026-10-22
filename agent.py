import os
import io
import re
import html
import json
import zipfile
import sqlite3
import requests
import docx
from google import genai
from google.genai import types

# 1. Конфигурация ключей из GitHub Secrets
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY")
GOOGLE_WEB_APP_URL = os.environ.get("GOOGLE_WEB_APP_URL")
SECRET_BRIDGE_KEY = os.environ.get("SECRET_BRIDGE_KEY", "FUEL_TENDER_SECRET_2026")

PROZORRO_API_BASE = "https://public-api.prozorro.gov.ua/api/2.5"
MIN_AMOUNT_UAH = 3000000
TARGET_CPV_PREFIXES = ("09130000", "09132000", "09134200")

# Быстрый фильтр для отсева очевидных талонов, карт и бланков в заголовке
STOP_WORDS_REGEX = re.compile(
    r"(талон|скретч|паливн[іи]х?\s+карт|картк|бланк[іи]?в?-дозвол|через\s+мережу\s+азс|на\s+азс|заправк[аи]\s+на\s+азс)",
    re.IGNORECASE
)

# Справочник видов торгов Prozorro
PROCUREMENT_METHODS = {
    "aboveThreshold": "Открытые торги с особенностями (Постановление КМУ №1178)",
    "aboveThresholdUA": "Открытые торги (украинские)",
    "aboveThresholdEU": "Открытые торги с публикацией на английском языке (EU)",
    "priceQuotation": "Запрос ценовых предложений (Prozorro Market / ЗЦП)",
    "negotiation": "Переговорная процедура",
    "negotiation.quick": "Переговорная процедура (ускоренная)",
    "belowThreshold": "Допороговая закупка"
}

# 2. Локальная база данных проверенных тендеров
conn = sqlite3.connect("processed_tenders.db")
cursor = conn.cursor()
cursor.execute("CREATE TABLE IF NOT EXISTS tenders (id TEXT PRIMARY KEY)")
conn.commit()

client = genai.Client(api_key=GEMINI_API_KEY)

def is_already_processed(tender_id: str) -> bool:
    cursor.execute("SELECT 1 FROM tenders WHERE id = ?", (tender_id,))
    return cursor.fetchone() is not None

def mark_as_processed(tender_id: str):
    cursor.execute("INSERT OR IGNORE INTO tenders (id) VALUES (?)", (tender_id,))
    conn.commit()

# 3. Формирование прямой ссылки на закупку на площадке SmartTender.biz
def build_smarttender_url(tender_ua_id: str) -> str:
    # Официальный шлюз SmartTender для прямого перехода к тендеру по идентификатору Prozorro
    gateway_url = f"https://smarttender.biz/government/details/{tender_ua_id}"
    try:
        resp = requests.head(gateway_url, allow_redirects=True, timeout=8)
        if resp.status_code == 200 and "smarttender.biz" in resp.url:
            return resp.url
    except Exception:
        pass
    return gateway_url

# 4. Чтение файлов .docx (абзацы + таблицы без дублирования объединенных ячеек)
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

# 5. Выгрузка актуальных редакций файлов ТД и проекта договора (.docx, .pdf, .zip)
def extract_all_attachments(documents_json: list) -> tuple[str, list]:
    latest_docs = {}
    for d in documents_json:
        if d.get("documentOf") != "tender":
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
        # Наивысший приоритет проекту договора, техзаданию и тендерной документации
        if any(k in t for k in ("догов", "контракт", "техніч", "документац", "додаток", "специфікац", "вимог")):
            return 0
        return 1

    valid_docs.sort(key=doc_priority)

    extracted_texts = []
    pdf_parts = []

    for d in valid_docs[:7]:
        title = d.get("title", "document")
        url = d.get("url", "")
        if not url:
            continue

        try:
            resp = requests.get(url, timeout=25)
            if resp.status_code != 200:
                continue
            content = resp.content
            low_title = title.lower()

            if low_title.endswith(".docx"):
                extracted_texts.append(parse_docx_bytes(content, title))
            elif low_title.endswith(".pdf"):
                if len(content) < 15 * 1024 * 1024 and len(pdf_parts) < 3:
                    pdf_parts.append(types.Part.from_bytes(data=content, mime_type="application/pdf"))
                    extracted_texts.append(f"[Прикреплен PDF-документ: {title}]")
            elif low_title.endswith(".zip"):
                with zipfile.ZipFile(io.BytesIO(content)) as zf:
                    for inner_name in zf.namelist()[:7]:
                        low_inner = inner_name.lower()
                        if low_inner.endswith((".p7s", ".sig")):
                            continue
                        inner_bytes = zf.read(inner_name)
                        if low_inner.endswith(".docx"):
                            extracted_texts.append(parse_docx_bytes(inner_bytes, f"{title}/{inner_name}"))
                        elif low_inner.endswith(".pdf") and len(inner_bytes) < 15 * 1024 * 1024 and len(pdf_parts) < 3:
                            pdf_parts.append(types.Part.from_bytes(data=inner_bytes, mime_type="application/pdf"))
                            extracted_texts.append(f"[Прикреплен PDF из архива: {inner_name}]")
        except Exception as e:
            extracted_texts.append(f"Ошибка загрузки файла {title}: {e}")

    return "\n\n".join(extracted_texts), pdf_parts

# 6. Глубокая юридическая и коммерческая экспертиза тендера через Gemini 2.5 Flash
def analyze_tender_with_gemini(tender_data: dict, docs_text: str, pdf_parts: list) -> dict:
    val = tender_data.get("value", {})
    budget_amount = val.get("amount", 0)
    vat_included = val.get("valueAddedTaxIncluded", True)
    guarantee_amount = tender_data.get("guarantee", {}).get("amount", 0)
    method_type = tender_data.get("procurementMethodType", "")
    method_readable = PROCUREMENT_METHODS.get(method_type, method_type)
    items_json = json.dumps(tender_data.get("items", []), ensure_ascii=False)

    prompt = f"""
    ПАСПОРТ ЗАКУПКИ ИЗ PROZORRO:
    Заголовок: {tender_data.get('title', '')}
    Описание: {tender_data.get('description', '')}
    Вид процедуры в системе: {method_readable} ({method_type})
    Бюджет закупки: {budget_amount} UAH ({'с НДС' if vat_included else 'без НДС'})
    Тендерное обеспечение в паспорте Prozorro: {guarantee_amount} UAH
    Номенклатура и объемы (items): {items_json}

    ПОЛНЫЙ ТЕКСТ ТЕНДЕРНОЙ ДОКУМЕНТАЦИИ И ПРОЕКТА ДОГОВОРА:
    {docs_text[:95000]}
    """

    contents = [prompt] + pdf_parts

    response = client.models.generate_content(
        model="gemini-2.5-flash",
        contents=contents,
        config=types.GenerateContentConfig(
            system_instruction=(
                "Ты — ведущий тендерный юрист и аналитик оптового поставщика нефтепродуктов в Украине. "
                "Проанализируй тендерную документацию и проект договора. "
                "КРИТЕРИЙ ОТБОРА (is_relevant = true): Закупается дизельное топливо или бензины НАЛИВОМ — "
                "как АВТОМОБИЛЬНЫМ ТРАНСПОРТОМ (бензовозами/автоцистернами в резервуары заказчика или самовывозом автоцистернами заказчика), "
                "так и ЖЕЛЕЗНОДОРОЖНЫМ ТРАНСПОРТОМ (железнодорожными вагон-цистернами на станцию назначения). "
                "Если закупаются талоны, топливные карты, скретч-карты, бланки-разрешения или розничная заправка через АЗС — строго верни is_relevant = false. "
                "Все ответы формируй на русском языке (прямые цитаты из украинского договора приводи дословно в кавычках с указанием номера пункта договора)."
            ),
            response_mime_type="application/json",
            response_schema={
                "type": "OBJECT",
                "properties": {
                    "is_relevant": {"type": "BOOLEAN"},
                    "subject_and_procedure": {
                        "type": "STRING",
                        "description": "Пункт 2: Предмет закупки и точный вид торгов (например: Открытые торги с особенностями / Запрос ценовых предложений)"
                    },
                    "fuel_grade_and_volume": {
                        "type": "STRING",
                        "description": "Пункт 3: Марка топлива (ДП Евро-5, А-95 и т.д.) и точный объем с единицами измерения (тонны / кг / литры)"
                    },
                    "unit_price_calculation": {
                        "type": "STRING",
                        "description": "Пункт 4: Расчетная предельная цена за 1 единицу (литр/тонну/кг) с НДС исходя из бюджета закупки и объема (если несколько позиций — указать расчет по ним)"
                    },
                    "payment_terms": {
                        "type": "STRING",
                        "description": "Пункт 5: Условия оплаты (отсрочка в банковских/календарных днях, по факту поставки или предоплата, перечень документов для оплаты)"
                    },
                    "is_formula_pricing": {
                        "type": "BOOLEAN",
                        "description": "true, если цена определяется формульным ценообразованием; false, если цена фиксированная"
                    },
                    "formula_details": {
                        "type": "STRING",
                        "description": "Пункт 6: Если цена формульная — точная математическая формула и описание всех ее коэффициентов (Platts/Argus, курс НБУ, акциз, маржа) согласно договору. Если фиксированная — напиши 'Цена фиксированная'"
                    },
                    "delivery_mode_and_location": {
                        "type": "STRING",
                        "description": "Пункт 8: Способ отгрузки (Автотранспорт / ЖД вагон-цистерны), базис Инкотермс (DDP/EXW/FCA/CPT) и точный адрес передачи товара или ж/д станция назначения (с кодом станции, если указан)"
                    },
                    "delivery_and_contract_term": {
                        "type": "STRING",
                        "description": "Пункт 9: Срок поставки отдельной партии по заявке (часов/дней), общий период поставки и срок действия договора (до какой даты)"
                    },
                    "bid_security_uah": {
                        "type": "STRING",
                        "description": "Пункт 10: Обеспечение тендерного предложения — точный размер гарантии в сумме (грн) и срок ее действия"
                    },
                    "contract_security_details": {
                        "type": "STRING",
                        "description": "Пункт 11: Обеспечение выполнения договора — вид обеспечения (банковская гарантия / депозит на счет заказчика) и размер в % от суммы договора (если не требуется — указать 'Не требуется')"
                    },
                    "features_pitfalls_and_risks": {
                        "type": "STRING",
                        "description": "Для кнопки 'Особенности тендера' (Блок 1): Подводные камни по договору, особенности приемки по количеству и качеству (риски убытков поставщика: простой авто/жд цистерн, экспертиза за счет поставщика, температурный пересчет в кг/литры и др.) с номерами пунктов и прямыми цитатами"
                    },
                    "features_penalties": {
                        "type": "STRING",
                        "description": "Для кнопки 'Особенности тендера' (Блок 2): Ответственность поставщика (штрафные санкции, пеня, оперативно-хозяйственные санкции) — точный размер в % или сумме и за что вменяется, с номерами пунктов договора и прямыми цитатами"
                    },
                    "features_price_increase_rules": {
                        "type": "STRING",
                        "description": "Для кнопки 'Особенности тендера' (Блок 3): Условия повышения цены за единицу товара (на основании каких документов можно поднять цену — ТПП, Держзовнішінформ, ДП Прозорро, Platts; какие условия и ограничения по частоте/срокам/процентам колебания) с номерами пунктов договора и прямыми цитатами"
                    }
                },
                "required": [
                    "is_relevant", "subject_and_procedure", "fuel_grade_and_volume",
                    "unit_price_calculation", "payment_terms", "is_formula_pricing",
                    "formula_details", "delivery_mode_and_location", "delivery_and_contract_term",
                    "bid_security_uah", "contract_security_details",
                    "features_pitfalls_and_risks", "features_penalties", "features_price_increase_rules"
                ]
            },
            temperature=0.0
        )
    )
    return json.loads(response.text)

# 7. Сборка карточки из 12 пунктов и передача в Google Sheets CRM
def dispatch_to_crm(tender_data: dict, ai: dict):
    tender_ua_id = tender_data.get("tenderID", "Не указан")
    entity = tender_data.get("procuringEntity", {})
    cust_name = entity.get("name", "Не указан")
    cust_edrpou = entity.get("identifier", {}).get("id", "Не указан")
    customer_str = f"{cust_name} (ЕГРПОУ: {cust_edrpou})"

    val = tender_data.get("value", {})
    budget_num = val.get("amount", 0)
    vat_str = "с НДС" if val.get("valueAddedTaxIncluded", True) else "без НДС"
    budget_full_str = f"{budget_num:,.2f} грн ({vat_str}) | Расчет за ед.: {ai['unit_price_calculation']}"

    deadline = tender_data.get("tenderPeriod", {}).get("endDate", "Не указан")[:16].replace("T", " ")

    # Получаем прямую ссылку на SmartTender.biz для колонки I и пункта 12
    smarttender_url = build_smarttender_url(tender_ua_id)

    # Форматирование пункта 6 (Спойлер для формулы или "Цена фиксированная")
    is_formula = bool(ai.get("is_formula_pricing"))
    formula_raw = ai.get("formula_details", "")
    if is_formula:
        pricing_html = (
            f"Цена формульная (нажмите на спойлер ниже или кнопку под сообщением):\n"
            f"<tg-spoiler>{html.escape(formula_raw)}</tg-spoiler>"
        )
    else:
        pricing_html = "Цена фиксированная"

    # Сборка сообщения строго по 12 обязательным пунктам ТЗ (HTML разметка)
    card_html = (
        f"<b>НОВЫЙ ТЕНДЕР (ПОСТАВКА НАЛИВОМ: АВТО / ЖД)</b>\n"
        f"<b>ID закупки:</b> <code>{html.escape(tender_ua_id)}</code>\n"
        f"-----------------------------------------\n"
        f"<b>1) Заказчик:</b> {html.escape(customer_str)}\n\n"
        f"<b>2) Предмет закупки и вид торгов:</b> {html.escape(ai['subject_and_procedure'])}\n\n"
        f"<b>3) Марка топлива и объем:</b> {html.escape(ai['fuel_grade_and_volume'])}\n\n"
        f"<b>4) Бюджет закупки (исходя из цены за ед. с НДС):</b> {html.escape(budget_full_str)}\n\n"
        f"<b>5) Условия оплаты:</b> {html.escape(ai['payment_terms'])}\n\n"
        f"<b>6) Ценообразование:</b> {pricing_html}\n\n"
        f"<b>7) Срок подачи тендерного предложения:</b> до {html.escape(deadline)}\n\n"
        f"<b>8) Условия поставки и адрес передачи товара (Авто / ЖД):</b> {html.escape(ai['delivery_mode_and_location'])}\n\n"
        f"<b>9) Срок поставки и срок действия договора:</b> {html.escape(ai['delivery_and_contract_term'])}\n\n"
        f"<b>10) Обеспечение тендерного предложения:</b> {html.escape(ai['bid_security_uah'])}\n\n"
        f"<b>11) Обеспечение договора:</b> {html.escape(ai['contract_security_details'])}\n\n"
        f"<b>12) Ссылка на SmartTender.biz:</b> <a href=\"{html.escape(smarttender_url)}\">{html.escape(smarttender_url)}</a>\n"
        f"-----------------------------------------\n"
        f"<i>Ответьте (Reply) свободным текстом на это сообщение, чтобы ИИ внес вашу цену, статус или результат в таблицу.</i>"
    )

    # Полный текст для кнопки "Особенности тендера" (сохраняется в колонку M и выводится по кнопке)
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
            "customerAndMethod": f"{customer_str} | {ai['subject_and_procedure']}",
            "budgetAndUnitPrice": budget_full_str,
            "fuelSummary": ai["fuel_grade_and_volume"],
            "deliveryAndTerms": f"{ai['delivery_mode_and_location']} | Сроки: {ai['delivery_and_contract_term']}",
            "paymentAndGuarantees": (
                f"Оплата: {ai['payment_terms']} | "
                f"Тендерная гарантия: {ai['bid_security_uah']} | "
                f"Обеспечение договора: {ai['contract_security_details']}"
            ),
            "deadline": deadline,
            "smartTenderUrl": smarttender_url,
            "isFormula": is_formula,
            "formulaText": formula_raw,
            "featuresDetailedText": features_detailed_text,
            "telegramCardHtml": card_html
        }
    }

    requests.post(GOOGLE_WEB_APP_URL, json=payload, timeout=25)

# 8. Главный цикл мониторинга Prozorro (с пагинацией на 15 страниц = 1500 тендеров)
def run_monitoring():
    next_url = f"{PROZORRO_API_BASE}/tenders?descending=1&limit=100&opt_fields=status,value"
    pages_to_scan = 15  # Проверяем 1500 последних тендеров в ленте Prozorro
    
    total_scanned = 0
    matched_budget = 0
    matched_fuel_cpv = 0
    sent_to_crm = 0

    for page in range(1, pages_to_scan + 1):
        if not next_url:
            break
            
        print(f"Чтение страницы {page}/{pages_to_scan} ленты Prozorro...")
        resp = requests.get(next_url, timeout=15).json()
        items_list = resp.get("data", [])
        next_url = resp.get("next_page", {}).get("uri")

        for item in items_list:
            total_scanned += 1
            t_id = item["id"]
            status = item.get("status", "")
            amount = item.get("value", {}).get("amount", 0)

            # 1. Фильтр по статусу и бюджету от 3 млн грн
            if status not in ("active.enquiries", "active.tendering") or amount < MIN_AMOUNT_UAH:
                continue
            if is_already_processed(t_id):
                continue

            matched_budget += 1

            # 2. Загрузка полной карточки тендера
            detail_resp = requests.get(f"{PROZORRO_API_BASE}/tenders/{t_id}", timeout=15).json()
            tender = detail_resp.get("data", {})
            tender_ua_id = tender.get("tenderID", t_id)

            # 3. Фильтр по кодам ДК 021:2015 (Нефть, бензин, дизельное топливо)
            items = tender.get("items", [])
            has_fuel_cpv = any(
                i.get("classification", {}).get("id", "").startswith(TARGET_CPV_PREFIXES)
                for i in items
            )
            if not has_fuel_cpv:
                mark_as_processed(t_id)
                continue

            matched_fuel_cpv += 1
            full_title = f"{tender.get('title', '')} {tender.get('description', '')}"
            print(f"Найден топливный тендер >= 3 млн: {tender_ua_id} | {amount:,.2f} грн | {tender.get('title', '')[:60]}")

            # 4. Быстрый отсев явных талонов и карт в заголовке
            if STOP_WORDS_REGEX.search(full_title):
                print(f" -> Отсеян по стоп-словам (талоны/карты/АЗС) в заголовке: {tender_ua_id}")
                mark_as_processed(t_id)
                continue

            # 5. Выгрузка документации и глубокий анализ через Gemini API
            print(f" -> Скачивание вложений и запуск экспертизы Gemini для {tender_ua_id}...")
            docs_text, pdf_parts = extract_all_attachments(tender.get("documents", []))
            ai_verdict = analyze_tender_with_gemini(tender, docs_text, pdf_parts)

            if ai_verdict.get("is_relevant"):
                print(f" -> ПОДТВЕРЖДЕН НАЛИВ! Отправка в Google Sheets и Telegram: {tender_ua_id}")
                dispatch_to_crm(tender, ai_verdict)
                sent_to_crm += 1
            else:
                print(f" -> Отклонен ИИ после чтения ТД (не налив): {tender_ua_id}")

            mark_as_processed(t_id)

    print("=========================================")
    print(f"ИТОГИ ЦИКЛА МОНИТОРИНГА:")
    print(f"- Всего просмотрено тендеров в ленте: {total_scanned}")
    print(f"- Активных тендеров с бюджетом >= 3 млн грн: {matched_budget}")
    print(f"- Из них по кодам топлива (0913...): {matched_fuel_cpv}")
    print(f"- Подтверждено наливом и отправлено в Telegram: {sent_to_crm}")
    print("=========================================")

if __name__ == "__main__":
    run_monitoring()
