import base64
import json
import streamlit as st
from openai import OpenAI

def get_ai_client() -> OpenAI:
    """Инициализация клиента DashScope через совместимый интерфейс OpenAI."""
    return OpenAI(
        api_key=st.secrets["DASHSCOPE_API_KEY"],
        base_url="https://dashscope.aliyuncs.com/compatible-mode/v1"
    )

def parse_json_response(content: str) -> dict:
    """Очищает ответ модели от markdown-разметки и парсит JSON."""
    text = content.strip()
    if text.startswith("```"):
        text = text.split("\n", 1)[1].rsplit("\n", 1)[0]
        if text.startswith("json"):
            text = text[4:].strip()
    try:
        return json.loads(text)
    except Exception:
        return {}

def extract_batch_info_from_image(image_bytes: bytes) -> dict:
    """Извлекает данные маркировочного стикера/коробки с фото в JSON."""
    client = get_ai_client()
    b64_image = base64.b64encode(image_bytes).decode("utf-8")
    
    prompt = """Ты — QC инспектор. Распознай шиппинг-марку/стикер коробки с фото.
Верни ТОЛЬКО валидный JSON со следующими ключами:
{
  "order_no": "номер заказа/PO или пустая строка",
  "item_name": "название товара или пустая строка",
  "sku": "артикул/SKU или пустая строка",
  "supplier": "поставщик/фабрика или пустая строка",
  "carton_qty": "количество в коробке/партии или пустая строка",
  "gross_weight": "вес брутто или пустая строка"
}"""

    response = client.chat.completions.create(
        model="qwen3-vl-plus",
        messages=[{
            "role": "user",
            "content": [
                {"type": "text", "text": prompt},
                {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{b64_image}"}}
            ]
        }],
        temperature=0.01
    )
    return parse_json_response(response.choices[0].message.content)

def digitize_spec_sheet(images_bytes_list: list) -> dict:
    """
    Оцифровывает эталонную размерную сетку (Size Spec) с 1 или 2 фото/техпака
    через qwen3-vl-plus и объединяет их в единую структуру.
    """
    client = get_ai_client()
    
    prompt = """Ты — технический аудитор текстильного производства.
Перед тобой фото размерной спецификации (Size Spec / POM Table), разбитой на 1 или 2 страницы/изображения.
Объедини данные со всех предоставленных изображений в единый строгий JSON.

ПРАВИЛА:
1. Точки замера (POM): код (code), наименование (description). Если на фото нет буквенного кода, проставь A, B, C... по порядку.
2. Допуски (Tolerances): tol_minus и tol_plus в абсолютных положительных числах (например, 1.0 или 0.5). Если допуск "+/- 1", то tol_minus=1.0, tol_plus=1.0.
3. Единый размерный ряд для всех найденных размеров (S, M, L, XL, XXL и т.д.) и их эталонные (Target) значения.
4. Исключай дублирующиеся строки, если одно изображение перекрывает другое.
5. unit: "cm" или "inch".

ВЕРНИ ТОЛЬКО ЧИСТЫЙ JSON:
{
  "unit": "cm",
  "sizes": ["S", "M", "L", "XL"],
  "poms": [
    {
      "code": "A",
      "description": "Total Length",
      "tol_minus": 1.0,
      "tol_plus": 1.0,
      "specs": {"S": 70.0, "M": 72.0, "L": 74.0, "XL": 76.0}
    }
  ]
}"""

    user_content = [{"type": "text", "text": prompt}]
    
    for img_bytes in images_bytes_list:
        b64_img = base64.b64encode(img_bytes).decode("utf-8")
        user_content.append({
            "type": "image_url",
            "image_url": {"url": f"data:image/jpeg;base64,{b64_img}"}
        })

    response = client.chat.completions.create(
        model="qwen3-vl-plus",
        messages=[{"role": "user", "content": user_content}],
        temperature=0.01
    )
    return parse_json_response(response.choices[0].message.content)

def polish_qc_text(raw_text: str, field_type: str = "comments") -> str:
    """Приводит черновик инспектора к строгому техническому тону (QC English)."""
    if not raw_text.strip():
        return ""
    
    client = get_ai_client()
    prompt = f"""Ты — старший сюрвейер международной инспекционной компании.
Отредактируй черновой текст инспектора для поля '{field_type}'.
ТРЕБОВАНИЯ:
1. Строгий, лаконичный английский язык (Technical QC English).
2. Сохрани ВСЕ факты, цифры и артикулы ровно так, как указал инспектор.
3. КАТЕГОРИЧЕСКИ ЗАПРЕЩЕНО добавлять вымышленные факты или менять смысл.
4. Верни только отредактированный текст без вступительных фраз.

Исходный текст:
{raw_text}"""

    response = client.chat.completions.create(
        model="qwen3-vl-plus",
        messages=[{"role": "user", "content": prompt}],
        temperature=0.1
    )
    return response.choices[0].message.content.strip()
