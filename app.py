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
    import google.genai as genai
except ImportError:
    install("google-genai")
    import google.genai as genai

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

# Настройка страницы
st.set_page_config(
    page_title="OrthoAI Pro — Клинический Координатор", 
    page_icon="🦷", 
    layout="wide",
    initial_sidebar_state="expanded"
)

# Кастомный CSS для крутого визуального стиля
st.markdown("""
    <style>
    .main {
        background-color: #0e1117;
    }
    .stButton>button {
        width: 100%;
        background: linear-gradient(45deg, #2b5876, #4e4376);
        color: white;
        font-weight: bold;
        border-radius: 8px;
        border: none;
        padding: 0.6rem 1rem;
        transition: 0.3s;
    }
    .stButton>button:hover {
        background: linear-gradient(45deg, #4e4376, #2b5876);
        box-shadow: 0 4px 12px rgba(78, 67, 118, 0.4);
    }
    .metric-card {
        background-color: #1e2530;
        padding: 15px;
        border-radius: 10px;
        border-left: 5px solid #4e4376;
    }
    </style>
""", unsafe_allow_html=True)

# Боковая панель (Sidebar) для настроек
with st.sidebar:
    st.image("https://img.icons8.com/color/96/tooth.png", width=70)
    st.title("OrthoAI Pro")
    st.markdown("---")
    
    st.subheader("🔑 Авторизация")
    api_key = st.text_input("Google Gemini API Key:", type="password", placeholder="AIzaSy...")
    
    st.markdown("---")
    st.subheader("⚙️ Параметры протокола")
    appliance_type = st.selectbox(
        "Основная аппаратура:",
        ["Брекет-система (Прямая дуга / MBT)", "Элайнеры (Прозрачные каппы)", "Комбинированное лечение", "Аппарат Forsus / Herbst / Дистализаторы"]
    )
    
    cepha_school = st.selectbox(
        "Цефалометрический стандарт:",
        ["Школа Шварца / Стенаера", "Анализ Риккетса", "Джакобс / ТРГ боковая проекция"]
    )
    
    st.markdown("---")
    st.info("💡 **Совет для ординатора:** Загружайте презентации с ТРГ-расчетами и фотографии окклюзии для максимальной точности биомеханики.")

# Главный экран
st.title("🦷 Интеллектуальный ортодонтический координатор")
st.markdown("Экспертная система анализа клинических данных, ТРГ и составления пошаговой биомеханики лечения под кураторством ИИ-профессора.")

# Колонки для разделения интерфейса
col1, col2 = st.columns([1, 1], gap="large")

with col1:
    st.markdown("### 📁 Загрузка данных пациента")
    uploaded_file = st.file_uploader(
        "Перетащите снимок, ТРГ-презентацию (PPTX) или документ", 
        type=["png", "jpg", "jpeg", "pptx", "pdf"]
    )
    
    extracted_text_from_file = ""

    if uploaded_file is not None:
        file_extension = uploaded_file.name.split('.')[-1].lower()
        
        if file_extension in ["png", "jpg", "jpeg"]:
            st.image(uploaded_file, caption="Загруженный снимок / фото", use_column_width=True)
        elif file_extension == "pptx":
            st.success("📊 Презентация успешно прочитана!")
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
                with st.expander("📄 Посмотреть извлеченный текст из презентации"):
                    st.text_area("", extracted_text_from_file, height=200)
            except Exception as e:
                st.error(f"Ошибка чтения презентации: {e}")
        elif file_extension == "pdf":
            st.info("📄 PDF документ прикреплен к анализу.")

with col2:
    st.markdown("### 📝 Клинические жалобы и анамнез")
    doctor_notes = st.text_area(
        "Опишите данные осмотра, углы, скученность или пожелания к лечению:", 
        placeholder="Например: Пациент с дистальным прикусом, класс II. Скученность во фронтальном отделе. Требуется расчет биомеханики с пружинами...",
        height=210
    )
    
    generate_btn = st.button("🚀 Сгенерировать экспертный план лечения")

# Системный промпт профессора
SYSTEM_PROMPT = f"""
Роль: Ты — ведущий профессор, доктор медицинских наук, заведующий кафедрой ортодонтии. Ты обучаешь своего ординатора составлению идеальных, пошаговых планов лечения.
Выбранный тип аппаратуры: {appliance_type}.
Цефалометрическая школа: {cepha_school}.

Правила анализа:
1. Биомеханика: Расписывай конкретные точки приложения сил, векторы тяг, где именно ставятся пружины (например, нитиноловая пружина Open Coil между зубами 13 и 15 по Виола), межчелюстные эластики, винты или аппараты.
2. Цефалометрия: Учитывай углы (ANB, SNB, SNA, ML-NSL) и объясняй их клиническое значение.
3. Структура ответа: Разбор цефалометрии, Этап 1 (Санация), Этап 2 (Нивелирование), Этап 3 (Биомеханика и коррекция прикуса), Этап 4 (Ретенция).
4. Тон: Профессиональный наставник.
Дисклеймер в конце: «План сформирован экспертной системой и подлежит утверждению лечащим врачом».
"""

# Генерация результата
if generate_btn:
    if not api_key:
        st.sidebar.error("⚠️ Введите API Key в боковой панели слева!")
        st.error("Пожалуйста, укажите ваш Google Gemini API Key в боковой панели.")
    elif not doctor_notes and not uploaded_file:
        st.warning("⚠️ Пожалуйста, добавьте описание или загрузите файл с данными пациента.")
    else:
        with st.spinner("🔬 Профессор-ортодонт рассчитывает биомеханику и составляет протокол..."):
            try:
                client = genai.Client(api_key=api_key)
                
                full_query = SYSTEM_PROMPT + "\n\nДанные от ординатора:\n" + doctor_notes
                if extracted_text_from_file:
                    full_query += f"\n\nДанные из слайдов презентации:\n{extracted_text_from_file}"

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

                st.markdown("---")
                st.success("✅ Экспертный клинический протокол успешно составлен!")
                
                # Вывод результата в красивом блоке
                st.markdown(response.text)
                
                # Кнопка скачивания отчета
                st.download_button(
                    label="📥 Скачать план лечения (.txt)",
                    data=response.text,
                    file_name="ortho_treatment_plan.txt",
                    mime="text/plain"
                )

            except Exception as e:
                st.error(f"Произошла ошибка при обращении к API: {e}")
