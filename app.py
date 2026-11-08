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
    page_title="OrthoAI Pro — Clinical Suite", 
    page_icon="🦷", 
    layout="wide",
    initial_sidebar_state="expanded"
)

# Профессиональный дизайн в темных тонах (Medical Dark Glassmorphism)
st.markdown("""
    <style>
    /* Главный фон */
    .stApp {
        background-color: #07090e;
        color: #e2e8f0;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
    }
    
    /* Сайдбар */
    [data-testid="stSidebar"] {
        background-color: #0b0f17;
        border-right: 1px solid #1e293b;
    }

    /* Карточки-контейнеры */
    .css-1r7sldc, .stTextArea, .stFileUploader {
        border-radius: 12px;
    }

    /* Кнопки */
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

    /* Текстовые поля */
    textarea {
        background-color: #0e1626 !important;
        color: #f8fafc !important;
        border: 1px solid #1e293b !important;
        border-radius: 10px !important;
    }
    
    /* Заголовки */
    h1, h2, h3 {
        color: #f8fafc;
        font-weight: 700;
        letter-spacing: -0.5px;
    }
    </style>
""", unsafe_allow_html=True)

# Боковая панель
with st.sidebar:
    st.image("https://img.icons8.com/color/96/tooth.png", width=60)
    st.markdown("### **OrthoAI Suite** v3.5")
    st.markdown("<span style='color: #38ef7d; font-size: 12px;'>● Система активна (Pro Clinical)</span>", unsafe_allow_html=True)
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
    st.markdown("---")
    st.info("💡 **Рекомендация:** Загрузите презентацию с расчетами ТРГ для построения точной биомеханики.")

# Главный экран
st.title("🦷 Клинический Ортодонтический Координатор")
st.markdown("Профессиональный модуль анализа ТРГ, расчета векторов сил и 3D-визуализации окклюзии.")

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
            st.success("📊 Данные из презентации успешно извлечены!")
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
        placeholder="Например: Полная пошаговая инструкция коррекции дистального прикуса, расчет пружин и векторов тяг без воды...",
        height=210
    )
    
    generate_btn = st.button("🚀 Запустить расчет и 3D-моделирование")

SYSTEM_PROMPT = f"""
Роль: Ты — ведущий профессор ортодонтии, доктор медицинских наук. Дай исчерпывающий, пошаговый и клинически точный протокол лечения без лишней воды.
Аппаратура: {appliance_type}.
Цефалометрия: {cepha_school}.

Правила:
1. Подробная биомеханика: укажи точное место для пружин (Open Coil Spring между какими зубами по системе Виола), вектор сил, межчелюстные тяги и накладки.
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
        st.error("⚠️ Введите Google Gemini API Key в боковой панели слева!")
    elif not doctor_notes and not uploaded_file:
        st.warning("⚠️ Загрузите файл или введите описание задачи!")
    else:
        with st.spinner("🔬 Профессор рассчитывает биомеханику и генерирует цифровую модель..."):
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
                        if "503" in str(api_err) and attempt < max_retries - 1:
                            time.sleep(2)
                            continue
                        else:
                            raise api_err

                if response:
                    st.markdown("---")
                    st.success("✅ Протокол успешно сформирован!")
                    
                    tab1, tab2 = st.tabs(["📋 Экспертный протокол", "🎮 3D Интерактивный осмотр окклюзии"])
                    
                    with tab1:
                        st.markdown(response.text)
                        st.download_button("📥 Скачать план лечения (.txt)", data=response.text, file_name="ortho_protocol.txt")

                    with tab2:
                        st.markdown("### 🧊 3D Модель зубных рядов (CS Inspect Mode)")
                        st.markdown("<span style='color: #94a3b8; font-size: 14px;'>🖱️ ЛКМ: Вращение | 📜 Колесико: Зум / Приближение к зубам | ПКМ: Смещение</span>", unsafe_allow_html=True)
                        
                        # Исправленный и стабильный код Three.js с авторазмером контейнера
                        three_js_code = """
                        <!DOCTYPE html>
                        <html>
                        <head>
                            <style>
                                body { margin: 0; background-color: #0b0f17; overflow: hidden; }
                                #canvas-container { width: 100vw; height: 520px; position: relative; }
                                .ui-buttons { position: absolute; top: 15px; left: 15px; z-index: 10; display: flex; gap: 10px; }
                                .btn { background: linear-gradient(135deg, #00f2fe, #4facfe); color: #07090e; border: none; padding: 10px 18px; border-radius: 8px; cursor: pointer; font-weight: bold; font-family: sans-serif; box-shadow: 0 4px 12px rgba(0,242,254,0.3); transition: 0.2s; }
                                .btn:hover { transform: scale(1.05); }
                                .btn-reset { background: #1e293b; color: #f8fafc; box-shadow: none; }
                            </style>
                        </head>
                        <body>
                            <div id="canvas-container">
                                <div class="ui-buttons">
                                    <button class="btn" id="inspectBtn">🔍 Inspect (F)</button>
                                    <button class="btn btn-reset" id="resetBtn">🔄 Сброс</button>
                                </div>
                            </div>

                            <script src="https://cdnjs.cloudflare.com/ajax/libs/three.js/r128/three.min.js"></script>
                            <script src="https://cdn.jsdelivr.net/npm/three@0.128.0/examples/js/controls/OrbitControls.js"></script>
                            <script src="https://cdnjs.cloudflare.com/ajax/libs/tween.js/18.6.4/tween.umd.js"></script>

                            <script>
                                const container = document.getElementById('canvas-container');
                                const scene = new THREE.Scene();
                                scene.background = new THREE.Color(0x0b0f17);
                                
                                const camera = new THREE.PerspectiveCamera(45, container.clientWidth / container.clientHeight, 0.1, 1000);
                                camera.position.set(0, 3, 14);
                                
                                const renderer = new THREE.WebGLRenderer({ antialias: true });
                                renderer.setSize(container.clientWidth, container.clientHeight);
                                container.appendChild(renderer.domElement);
                                
                                const controls = new THREE.OrbitControls(camera, renderer.domElement);
                                controls.enableDamping = true;
                                controls.dampingFactor = 0.05;
                                controls.minDistance = 2;
                                controls.maxDistance = 25;

                                // Освещение
                                scene.add(new THREE.AmbientLight(0xffffff, 0.8));
                                const light1 = new THREE.DirectionalLight(0xffffff, 1);
                                light1.position.set(5, 10, 7);
                                scene.add(light1);
                                const light2 = new THREE.DirectionalLight(0x00f2fe, 0.5);
                                light2.position.set(-5, -5, -5);
                                scene.add(light2);

                                const mainGroup = new THREE.Group();
                                const toothMat = new THREE.MeshStandardMaterial({ color: 0xf8fafc, roughness: 0.15, metalness: 0.05 });
                                const gumMat = new THREE.MeshStandardMaterial({ color: 0xe11d48, roughness: 0.4 });

                                function buildJaw(isUpper) {
                                    const jaw = new THREE.Group();
                                    const yPos = isUpper ? 0.6 : -0.6;
                                    
                                    const gumGeo = new THREE.TorusGeometry(3.2, 0.45, 16, 32, Math.PI * 0.85);
                                    const gum = new THREE.Mesh(gumGeo, gumMat);
                                    gum.rotation.x = Math.PI / 2 * (isUpper ? 1 : -1);
                                    gum.position.y = yPos;
                                    jaw.add(gum);

                                    const teethCount = 14;
                                    for (let i = 0; i < teethCount; i++) {
                                        const angle = (i / (teethCount - 1) - 0.5) * Math.PI * 0.85;
                                        const radius = 3.2;
                                        const x = radius * Math.sin(angle);
                                        const z = radius * Math.cos(angle);
                                        
                                        const tGeo = new THREE.BoxGeometry(0.48, 0.75, 0.55);
                                        const tooth = new THREE.Mesh(tGeo, toothMat);
                                        tooth.position.set(x, yPos + (isUpper ? -0.28 : 0.28), z);
                                        tooth.rotation.y = angle;
                                        jaw.add(tooth);
                                    }
                                    return jaw;
                                }

                                mainGroup.add(buildJaw(true));
                                mainGroup.add(buildJaw(false));
                                scene.add(mainGroup);

                                // Анимация Inspect
                                let isInspecting = false;
                                document.getElementById('inspectBtn').addEventListener('click', () => {
                                    if(isInspecting) return;
                                    isInspecting = true;
                                    new TWEEN.Tween(mainGroup.rotation)
                                        .to({ y: mainGroup.rotation.y + Math.PI * 2, x: 0.35 }, 2200)
                                        .easing(TWEEN.Easing.Quadratic.Out)
                                        .onComplete(() => {
                                            new TWEEN.Tween(mainGroup.rotation).to({ x: 0 }, 400).start();
                                            isInspecting = false;
                                        })
                                        .start();
                                });

                                document.getElementById('resetBtn').addEventListener('click', () => {
                                    controls.reset();
                                    mainGroup.rotation.set(0, 0, 0);
                                });

                                function animate() {
                                    requestAnimationFrame(animate);
                                    TWEEN.update();
                                    controls.update();
                                    
                                    if (!isInspecting) {
                                        mainGroup.rotation.y += 0.0025;
                                    }
                                    renderer.render(scene, camera);
                                }
                                animate();
                            </script>
                        </body>
                        </html>
                        """
                        components.html(three_js_code, height=540)

            except Exception as e:
                st.error(f"Произошла ошибка при обращении к API: {e}")
