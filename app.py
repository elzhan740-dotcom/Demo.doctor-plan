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
    from google import genai
except ImportError:
    install("google-genai")
    from google import genai

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

import io

st.set_page_config(page_title="Стоматологический Эксперт-Координатор", page_icon="🦷", layout="centered")

st.title("🦷 Ортодонтический и клинический план лечения (Gemini)")
st.write("Загрузите снимки, документы или презентацию (PPTX) для получения экспертного протокола лечения с подробной биомеханикой.")

api_key = st.text_input("Введите ваш Google Gemini API Key:", type="password")

SYSTEM_PROMPT = """
Роль: Ты — ведущий профессор, доктор медицинских наук, заведующий кафедрой ортодонтии и эксперт высшей категории. Ты обучаешь своего ординатора (ученика-врача) составлению идеальных, пошаговых и высокоточных планов ортодонтического лечения. 

Правила работы и мышления эксперта:
1. Максимальная ортодонтическая точность и биомеханика: Никогда не пиши общими фразами вроде «установить брекеты» или «исправить прикус». Ты обязан расписать пошаговую биомеханику:
   - Куда конкретно устанавливается пружина (например: раскрывающая нитиноловая пружина Open Coil Spring на дуге 0.016x0.022 SS между зубами 13 и 15 по системе Виола для создания места под зуб 14).
   - Какие тяги и эластики необходимы (например: межчелюстные тяги II класса 3/16 4.5 oz от клыка в/ч к моляру н/ч).
   - Применение аппаратуры и опор (накладки, дистализаторы, аппараты Forsus/Herbst, минивинты/TADs) с четким указанием векторов сил и анатомических ориентиров.
2. Глубокий цефалометрический анализ: Анализируй углы (ANB, SNB, SNA, ML-NSL, Ar-Go-Me) и положение резцов. Объясняй ученику причинно-следственные связи: почему при данном скелетном классе и типе роста мы выбираем именно эту механику.
3. Обучающий тон наставника: Обращайся к врачу как строгий, но мудрый профессор: «Коллега, обратите внимание на...», «Для коррекции этого угла нам потребуется...», «Обязательно установите пружину на этапе...».
4. Пошаговая структура по международной системе Виола (11-48).

Формат вывода ответа (строго соблюдай разделы):
• Разбор цефалометрии и диагноз (клинический разбор профессора).
• Этап 1. Терапевтическая подготовка и санация (какие зубы пролечить/удалить перед ортодонтией).
• Этап 2. Нивелирование и выравнивание (последовательность дуг, стартовые манипуляции).
• Этап 3. Биомеханика и коррекция прикуса (детализированные указания: где именно ставить пружины, накладки, тяги, векторы сил, дистализацию или винты).
• Этап 4. Завершение, юстировка окклюзии и ретенция.
• Рекомендации и риски для ординатора.

Дисклеймер в конце: «План сформирован экспертной обучающей системой на основе представленных данных и подлежит обязательному утверждению лечащим врачом после очного осмотра».
"""

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
        st.info("📊 Презентация PowerPoint успешно загружена. Извлекаем данные слайдов...")
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

doctor_notes = st.text_area("Опишите жалобы, данные осмотра или дополнительные задачи:", placeholder="Например: Жалобы на дистальный прикус, скученность. Проверьте вектор тяги и пружины...")

if st.button("🚀 Сгенерировать план лечения"):
    if not api_key:
        st.error("Пожалуйста, введите ваш Google Gemini API Key.")
    elif not doctor_notes and not uploaded_file:
        st.warning("Пожалуйста, введите описание или загрузите файл/снимок.")
    else:
        with st.spinner("Профессор-ортодонт анализирует биомеханику и формирует клинический протокол..."):
            try:
                client = genai.Client(api_key=api_key)
                
                full_query = SYSTEM_PROMPT + "\n\nДанные от врача/ординатора:\n" + doctor_notes
                if extracted_text_from_file:
                    full_query += f"\n\nДанные из презентации (ТРГ/КТ/Карточка):\n{extracted_text_from_file}"

                contents = [full_query]
                
                if uploaded_file is not None:
                    file_extension = uploaded_file.name.split('.')[-1].lower()
                    if file_extension in ["png", "jpg", "jpeg"]:
                        img = Image.open(uploaded_file)
                        contents.append(img)

                response = client.models.generate_content(
                    model='gemini-3.8-flash',
                    contents=contents
                )

                st.success("Экспертный план лечения составлен!")
                st.markdown("---")
                st.markdown(response.text)

            except Exception as e:
                st.error(f"Произошла ошибка при обращении к Gemini: {e}")
