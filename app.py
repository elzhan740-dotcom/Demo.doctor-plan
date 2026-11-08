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
    import streamlit as st

try:
    from pptx import Presentation
except ImportError:
    install("python-pptx")
    import streamlit as st

import io
import time
import streamlit.components.v1 as components

# Настройка страницы
st.set_page_config(
    page_title="OrthoAI Pro — Clinical Suite 3D", 
    page_icon="🦷", 
    layout="wide",
    initial_sidebar_state="expanded"
)

# Профессиональный интерфейс в темных тонах
st.markdown("""
    <style>
    .stApp {
        background-color: #07090e;
        color: #e2e8f0;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
        overflow-y: auto !important;
    }
    [data-testid="stSidebar"] {
        background-color: #0b0f17;
        border-right: 1px solid #1e293b;
    }
    .stButton>button {
        width: 100%;
        background: linear-gradient(135deg, #00f2fe 0%, #4facfe 100%);
        color: #07090e;
        font-weight: 700;
        border-radius: 10px;
        border: none;
        padding: 0.75rem 1rem;
        font-size: 16px;
        box-shadow: 0 4px 14px rgba(79, 172, 254, 0.3);
        transition: all 0.3s ease;
    }
    .stButton>button:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 20px rgba(79, 172, 254, 0.5);
    }
    textarea {
        background-color: #0e1626 !important;
        color: #f8fafc !important;
        border: 1px solid #1e293b !important;
        border-radius: 10px !important;
    }
    h1, h2, h3 {
        color: #f8fafc;
        font-weight: 700;
        letter-spacing: -0.5px;
    }
    .block-container {
        padding-bottom: 5rem;
    }
    </style>
""", unsafe_allow_html=True)

if "generated_result" not in st.session_state:
    st.session_state.generated_result = None

# Боковая панель
with st.sidebar:
    st.image("https://img.icons8.com/color/96/tooth.png", width=60)
    st.markdown("### **OrthoAI Suite** v4.1")
    st.markdown("<span style='color: #38ef7d; font-size: 12px;'>● 3D Дуга Челюсти Активна</span>", unsafe_allow_html=True)
    st.markdown("---")
    
    api_key = st.text_input("🔑 Google Gemini API Key:", type="password", placeholder="Введите ключ...")
    
    st.markdown("---")
    st.markdown("### ⚙️ Протокол лечения")
    appliance_type = st.selectbox(
        "Аппаратура:",
        ["Брекет-система (MBT / Прямая дуга)", "Элайнеры (Прозрачные каппы)", "Комбинированная биомеханика"]
    )
    
    cepha_school = st.selectbox(
        "Цефалометрия:",
        ["Школа Шварца / Стенаера", "Анализ Риккетса", "Джакобс / ТРГ"]
    )

# Главный экран
st.title("🦷 Клинический Ортодонтический Координатор")
st.markdown("Профессиональный модуль анализа ТРГ, расчета биомеханики и 3D-осмотра зубных рядов по системе Виола (FDI).")

col1, col2 = st.columns([1, 1], gap="large")

with col1:
    st.markdown("### 📁 Материалы пациента")
    uploaded_file = st.file_uploader(
        "Загрузите ТРГ-снимок, презентацию (PPTX) или документ", 
        type=["png", "jpg", "jpeg", "pptx", "pdf"]
    )
    
    extracted_text_from_file = ""
    if uploaded_file is not None:
        file_ext = uploaded_file.name.split('.')[-1].lower()
        if file_ext in ["png", "jpg", "jpeg"]:
            st.image(uploaded_file, caption="Загруженный снимок", use_column_width=True)
        elif file_ext == "pptx":
            st.success("📊 Данные из презентации извлечены!")
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
                with st.expander("📄 Просмотреть текст слайдов"):
                    st.text_area("", extracted_text_from_file, height=150)
            except Exception as e:
                st.error(f"Ошибка чтения: {e}")

with col2:
    st.markdown("### 📝 Клиническая задача")
    doctor_notes = st.text_area(
        "Опишите диагноз и требования к лечению:", 
        placeholder="Например: Пошаговая инструкция коррекции прикуса, точное расположение пружин по номерам зубов...",
        height=210
    )
    
    generate_btn = st.button("🚀 Запустить расчет и 3D-моделирование")

SYSTEM_PROMPT = f"""
Роль: Ты — ведущий профессор ортодонтии, доктор медицинских наук. Дай пошаговый и клинически точный протокол лечения без лишней воды.
Аппаратура: {appliance_type}.
Цефалометрия: {cepha_school}.

Правила:
1. Подробная биомеханика: укажи точное место для пружин (Open Coil Spring между какими зубами по системе Виола 11-48), вектор сил, межчелюстные тяги.
2. Формат вывода:
   - Анализ цефалометрии и диагноз.
   - Этап 1. Санация.
   - Этап 2. Нивелирование.
   - Этап 3. Биомеханика и прикус (детально по пружинам/тягам).
   - Этап 4. Ретенция.
   - 3D-Финал и прогноз.
Дисклеймер в конце обязательно.
"""

if generate_btn:
    if not api_key:
        st.error("⚠️ Введите Google Gemini API Key в боковой панели!")
    elif not doctor_notes and not uploaded_file:
        st.warning("⚠️ Загрузите файл или введите описание задачи!")
    else:
        with st.spinner("🔬 Профессор рассчитывает биомеханику и генерирует 3D-схему..."):
            try:
                client = genai.Client(api_key=api_key)
                
                full_query = SYSTEM_PROMPT + "\n\nДанные ординатора:\n" + doctor_notes
                if extracted_text_from_file:
                    full_query += f"\n\nДанные PPTX:\n{extracted_text_from_file}"

                contents = [full_query]
                if uploaded_file is not None and uploaded_file.name.split('.')[-1].lower() in ["png", "jpg", "jpeg"]:
                    contents.append(Image.open(uploaded_file))

                response = None
                max_retries = 3
                for attempt in range(max_retries):
                    try:
                        response = client.models.generate_content(
                            model='gemini-3.8-flash',
                            contents=contents
                        )
                        break
                    except Exception as api_err:
                        if ("503" in str(api_err) or "unavailable" in str(api_err).lower()) and attempt < max_retries - 1:
                            time.sleep(3)
                            continue
                        else:
                            raise api_err

                if response:
                    st.session_state.generated_result = response.text

            except Exception as e:
                st.error(f"⚠️ Ошибка при обращении к API: {e}")

# Вывод результатов и 3D-модели
if st.session_state.generated_result:
    st.markdown("---")
    st.success("✅ Протокол сформирован!")
    
    tab1, tab2 = st.tabs(["📋 Экспертный протокол", "🎮 3D Дуга челюсти (Система Виола)"])
    
    with tab1:
        st.markdown(st.session_state.generated_result)
        st.download_button(
            "📥 Скачать план лечения (.txt)", 
            data=st.session_state.generated_result, 
            file_name="ortho_protocol.txt"
        )

    with tab2:
        st.markdown("### 🧊 Объемная 3D-модель зубных рядов (Анатомическая дуга)")
        st.markdown("<span style='color: #94a3b8; font-size: 14px;'>🖱️ Зажмите ЛКМ и вращайте челюсть | 📜 Колесико мыши: Приближение</span>", unsafe_allow_html=True)
        
        native_3d_code = """
        <!DOCTYPE html>
        <html>
        <head>
            <style>
                body { margin: 0; background-color: #0b0f17; overflow: hidden; display: flex; justify-content: center; align-items: center; height: 620px; }
                canvas { background: #07090e; border-radius: 12px; border: 1px solid #1e293b; cursor: grab; }
                canvas:active { cursor: grabbing; }
                .hud { position: absolute; top: 15px; left: 15px; color: #00f2fe; font-family: monospace; font-size: 13px; pointer-events: none; background: rgba(0,0,0,0.7); padding: 6px 12px; border-radius: 6px; border: 1px solid #1e293b; }
            </style>
        </head>
        <body>
            <div class="hud">SYSTEM: 3D_JAW_ARCH // FDI_VIOLA</div>
            <canvas id="c" width="850" height="580"></canvas>
            <script>
                const canvas = document.getElementById('c');
                const ctx = canvas.getContext('2d');

                let rotX = 0.6;
                let rotY = 0.3;
                let zoom = 1.0;
                let isDragging = false;
                let lastX = 0, lastY = 0;

                canvas.addEventListener('mousedown', e => { isDragging = true; lastX = e.clientX; lastY = e.clientY; });
                window.addEventListener('mouseup', () => isDragging = false);
                window.addEventListener('mousemove', e => {
                    if (!isDragging) return;
                    rotY += (e.clientX - lastX) * 0.008;
                    rotX += (e.clientY - lastY) * 0.008;
                    lastX = e.clientX;
                    lastY = e.clientY;
                });
                canvas.addEventListener('wheel', e => {
                    e.preventDefault();
                    zoom += e.deltaY * -0.001;
                    zoom = Math.max(0.6, Math.min(2.5, zoom));
                });

                const upperFDI = [18, 17, 16, 15, 14, 13, 12, 11, 21, 22, 23, 24, 25, 26, 27, 28];
                const lowerFDI = [48, 47, 46, 45, 44, 43, 42, 41, 31, 32, 33, 34, 35, 36, 37, 38];

                function getArch() {
                    let arch = [];
                    // Верхняя челюсть (параболическая дуга)
                    for (let i = 0; i < 16; i++) {
                        let t = (i / 15) * 2 - 1;
                        let x = t * 130;
                        let z = (Math.pow(t, 2) * -90) + 50;
                        arch.push({ x, y: -45, z, fdi: upperFDI[i], jaw: 'upper' });
                    }
                    // Нижняя челюсть
                    for (let i = 0; i < 16; i++) {
                        let t = (i / 15) * 2 - 1;
                        let x = t * 120;
                        let z = (Math.pow(t, 2) * -80) + 40;
                        arch.push({ x, y: 45, z, fdi: lowerFDI[i], jaw: 'lower' });
                    }
                    return arch;
                }
                const teeth = getArch();

                function draw() {
                    ctx.clearRect(0, 0, canvas.width, canvas.height);
                    ctx.save();
                    ctx.translate(canvas.width / 2, canvas.height / 2);

                    if (!isDragging) rotY += 0.004;

                    let projected = teeth.map(t => {
                        let x1 = t.x * Math.cos(rotY) - t.z * Math.sin(rotY);
                        let z1 = t.x * Math.sin(rotY) + t.z * Math.cos(rotY);
                        let y1 = t.y;

                        let y2 = y1 * Math.cos(rotX) - z1 * Math.sin(rotX);
                        let z2 = y1 * Math.sin(rotX) + z1 * Math.cos(rotX);
                        let x2 = x1;

                        let fov = 500;
                        let scale = (fov / (fov + z2 + 300)) * zoom;
                        return {
                            x: x2 * scale,
                            y: y2 * scale,
                            z: z2,
                            scale: scale,
                            fdi: t.fdi,
                            jaw: t.jaw
                        };
                    });

                    projected.sort((a, b) => a.z - b.z);

                    projected.forEach(p => {
                        ctx.save();
                        ctx.translate(p.x, p.y);
                        ctx.scale(p.scale, p.scale);

                        // Форма зуба
                        ctx.fillStyle = '#f1f5f9';
                        ctx.strokeStyle = '#0284c7';
                        ctx.lineWidth = 2;

                        ctx.beginPath();
                        ctx.roundRect(-12, -16, 24, 32, 6);
                        ctx.fill();
                        ctx.stroke();

                        // Брекет
                        ctx.fillStyle = '#00f2fe';
                        ctx.fillRect(-4, -4, 8, 8);

                        // Номер Виола
                        ctx.fillStyle = '#38ef7d';
                        ctx.font = 'bold 11px monospace';
                        ctx.textAlign = 'center';
                        ctx.fillText(p.fdi, 0, p.jaw === 'upper' ? -24 : 26);

                        ctx.restore();
                    });

                    ctx.restore();
                    requestAnimationFrame(draw);
                }
                draw();
            </script>
        </body>
        </html>
        """
        components.html(native_3d_code, height=640)
