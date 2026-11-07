import subprocess
import sys

def install(package):
    subprocess.check_call([sys.executable, "-m", "pip", "install", package])

try:
    import streamlit as st
except ImportError:
    install("streamlit")
    import streamlit as st

try:
    from openai import OpenAI
except ImportError:
    install("openai")
    from openai import OpenAI

try:
    from PIL import Image
except ImportError:
    install("Pillow")
    from PIL import Image

try:
    from pptx import Presentation
except ImportError:
    install("python-pptx")
    from pptx import Presentation

import base64
import io

st.set_page_config(page_title="Стоматологический Координатор", page_icon="🦷", layout="centered")

st.title("🦷 Автоматический план стоматологического лечения")
st.write("Загрузите снимок, документ или презентацию (PPTX) вместе с описанием проблемы, чтобы получить детальный план лечения по протоколам.")

api_key = st.text_input("Введите ваш OpenAI API Key:", type="password")

SYSTEM_PROMPT = """
Роль: Ты — опытный стоматолог-координатор и ассистент врача-стоматолога в современной клинике. Твоя задача — на основе введенных врачом данных (жалобы, анамнез, результаты осмотра, КТ/рентгена, презентаций и прикрепленных изображений) составить подробный, структурированный и логичный план лечения.

Правила работы:
1. Строго по данным: Опирайся только на переданную информацию (текст, файлы и фото). Не придумывай несуществующие симптомы или диагнозы. Если данных недостаточно, укажи это в блоке «Необходимая додиагностика».
2. Логическая последовательность: Всегда выстраивай этапы лечения по медицинским протоколам:
   • Этап 1: Острая боль и санация (профгигиена, лечение кариеса, пульпита, периодонтита, удаление безнадежных зубов).
   • Этап 2: Подготовка и ортодонтия (установка брекет-систем, элайнеров, костная пластика, синус-лифтинг — если применимо).
   • Этап 3: Имплантация и протезирование (установка имплантов, коронок, виниров — после или в процессе ортодонтического лечения).
   • Этап 4: Профилактика и уход (рекомендации, уход за ортодонтическими конструкциями и график профосмотров).
3. Нумерация зубов: Используй международную систему Виола (двухцифровую: 11–48).
4. Альтернативы: Если возможно несколько вариантов лечения (например, брекеты vs элайнеры; мост vs имплант), предложи основной и альтернативный варианты, указав плюсы и минусы.
5. Тон общения: Профессиональный, понятный для пациента, с расшифровкой терминов в скобках.

Формат вывода (строго соблюдай структуру):
• Предварительный диагноз / Клиническая картина: (краткое резюме по зубам и прикусу).
• Этап 1. Санация и терапия: (список манипуляций по каждому зубу).
• Этап 2. Ортодонтия (брекеты и подготовка): (установка брекет-системы, коррекция).
• Этап 3. Хирургия и имплантация: (если требуется, в т.ч. удаление зубов мудрости).
• Этап 4. Ортопедия (протезирование): (если требуется после ортодонтии).
• Альтернативные варианты лечения: (если есть).
• Ориентировочные сроки и количество визитов: (примерный прогноз).
• Рекомендации: (уход, гигиена с брекетами, ограничения, анализы).

Ограничение: В конце каждого ответа добавляй дисклеймер: «План сформирован автоматически на основе первичных данных и подлежит утверждению лечащим врачом после очного осмотра».
"""

# Добавили поддержку pptx и pdf в загрузчик
uploaded_file = st.file_uploader(
    "Загрузите снимок, презентацию (PPTX) или документ", 
    type=["png", "jpg", "jpeg", "pptx", "pdf"]
)

extracted_text_from_file = ""

if uploaded_file is not None:
    file_extension = uploaded_file.name.split('.')[-1].lower()
    
    if file_extension in ["png", "jpg", "jpeg"]:
        st.image(uploaded_file, caption="Загруженное изображение / снимок", use_column_width=True)
    elif file_extension == "pptx":
        st.info("📊 Презентация PowerPoint успешно загружена. Извлекаем текст со слайдов...")
        try:
            prs = Presentation(io.BytesIO(uploaded_file.getvalue()))
            slide_texts = []
            for i, slide in enumerate(prs.slides):
                slide_text = f"--- Слайд {i+1} ---\n"
                for shape in slide.shapes:
                    if shape.has_text_frame:
                        for paragraph in shape.text_frame.paragraphs:
                            slide_text += paragraph.text + "\n"
                slide_texts.append(slide_text)
            extracted_text_from_file = "\n".join(slide_texts)
            st.text_area("Извлеченный текст из презентации:", extracted_text_from_file, height=150)
        except Exception as e:
            st.error(f"Не удалось прочитать презентацию: {e}")
    elif file_extension == "pdf":
        st.info("📄 PDF документ загружен.")

doctor_notes = st.text_area("Опишите жалобы, данные осмотра и укажите проблемные зубы:", placeholder="Например: Жалобы на скученность во фронтальном отделе. На зубе 36 кариес...")

def encode_image(image_file):
    return base64.b64encode(image_file.getvalue()).decode('utf-8')

if st.button("🚀 Сгенерировать план лечения"):
    if not api_key:
        st.error("Пожалуйста, введите ваш API-ключ OpenAI.")
    elif not doctor_notes and not uploaded_file:
        st.warning("Пожалуйста, введите описание или загрузите файл/снимок.")
    else:
        with st.spinner("ИИ анализирует данные и формирует протокол лечения..."):
            try:
                client = OpenAI(api_key=api_key)
                
                # Формируем общий текст запроса, включая данные из презентации, если она была загружена
                full_query = doctor_notes
                if extracted_text_from_file:
                    full_query += f"\n\nДанные из прикрепленной презентации (PPTX):\n{extracted_text_from_file}"

                content_messages = [{"type": "text", "text": full_query}]
                
                # Если загружена именно картинка, передаем её в OpenAI Vision
                if uploaded_file is not None:
                    file_extension = uploaded_file.name.split('.')[-1].lower()
                    if file_extension in ["png", "jpg", "jpeg"]:
                        base64_img = encode_image(uploaded_file)
                        content_messages.append({
                            "type": "image_url",
                            "image_url": {"url": f"data:image/jpeg;base64,{base64_img}"}
                        })

                response = client.chat.completions.create(
                    model="gpt-4o",
                    messages=[
                        {"role": "system", "content": SYSTEM_PROMPT},
                        {"role": "user", "content": content_messages}
                    ],
                    max_tokens=2000
                )

                plan_result = response.choices[0].message.content
                st.success("План успешно составлен!")
                st.markdown("---")
                st.markdown(plan_result)

            except Exception as e:
                st.error(f"Произошла ошибка при обращении к ИИ: {e}")
