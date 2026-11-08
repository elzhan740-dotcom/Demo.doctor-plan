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
import streamlit.components.v1 as components

# Настройка страницы
st.set_page_config(
    page_title="OrthoAI CS-Inspect 3D", 
    page_icon="🦷", 
    layout="wide",
    initial_sidebar_state="expanded"
)

# Кастомный CSS в тёмно-геймерском/клиническом стиле
st.markdown("""
    <style>
    .main { background-color: #0b0e14; }
    .stButton>button {
        width: 100%;
        background: linear-gradient(135deg, #FF4655, #0F1923);
        color: white;
        font-weight: bold;
        border-radius: 8px;
        border: 1px solid #FF4655;
        padding: 0.7rem 1rem;
        font-size: 16px;
        transition: 0.3s;
    }
    .stButton>button:hover {
        background: linear-gradient(135deg, #0F1923, #FF4655);
        box-shadow: 0 0 15px rgba(255, 70, 85, 0.6);
    }
    </style>
""", unsafe_allow_html=True)

# Боковая панель
with st.sidebar:
    st.image("https://img.icons8.com/color/96/tooth.png", width=70)
    st.title("OrthoAI Pro 3D")
    st.markdown("---")
    
    api_key = st.text_input("🔑 Google Gemini API Key:", type="password")
    
    st.markdown("---")
    st.subheader("⚙️ Параметры лечения")
    appliance_type = st.selectbox(
        "Тип аппаратуры:",
        ["Брекет-система (MBT / Прямая дуга)", "Элайнеры (Прозрачные каппы)", "Комбинированная биомеханика"]
    )
    
    cepha_school = st.selectbox(
        "Цефалометрия:",
        ["Школа Шварца / Стенаера", "Анализ Риккетса", "Джакобс / ТРГ"]
    )

st.title("🦷 3D-Осмотр окклюзии и биомеханики (CS Mode)")
st.markdown("Анализ снимков, расчёт ортодонтической биомеханики и 3D-осмотр результата в стиле CS:GO / CS2.")

col1, col2 = st.columns([1, 1], gap="large")

with col1:
    st.markdown("### 📁 Загрузка материалов пациента")
    uploaded_file = st.file_uploader(
        "Загрузите снимок, ТРГ-презентацию (PPTX) или PDF", 
        type=["png", "jpg", "jpeg", "pptx", "pdf"]
    )
    
    extracted_text_from_file = ""
    if uploaded_file is not None:
        file_ext = uploaded_file.name.split('.')[-1].lower()
        if file_ext in ["png", "jpg", "jpeg"]:
            st.image(uploaded_file, caption="Загруженный снимок", use_column_width=True)
        elif file_ext == "pptx":
            st.success("📊 Презентация прочитана!")
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
                with st.expander("📄 Посмотреть извлеченный текст из PPTX"):
                    st.text_area("", extracted_text_from_file, height=150)
            except Exception as e:
                st.error(f"Ошибка чтения: {e}")

with col2:
    st.markdown("### 📝 Данные осмотра и задачи")
    doctor_notes = st.text_area(
        "Описание клинической картины:", 
        placeholder="Дистальный прикус, скученность. Укажите биомеханику, пружины и векторы сил...",
        height=200
    )
    
    generate_btn = st.button("🚀 Сгенерировать план и отсмотреть 3D-модель")

SYSTEM_PROMPT = f"""
Роль: Ты — ведущий профессор ортодонтии, доктор медицинских наук.
Аппаратура: {appliance_type}.
Цефалометрия: {cepha_school}.

Правила:
1. Подробная биомеханика: укажи exact точное место для пружин (Open Coil Spring между какими зубами по Виола), вектор сил, межчелюстные тяги и накладки.
2. Формат вывода:
   - Анализ цефалометрии и диагноз.
   - Этап 1. Санация.
   - Этап 2. Нивелирование.
   - Этап 3. Биомеханика и прикус (детально по пружинам/тягам).
   - Этап 4. Ретенция.
   - 3D-Финал и прогноз: детали изменений окклюзии.
Дисклеймер в конце обязательно.
"""

if generate_btn:
    if not api_key:
        st.error("⚠️ Введите API Key в боковой панели слева!")
    elif not doctor_notes and not uploaded_file:
        st.warning("⚠️ Загрузите файл или введите описание!")
    else:
        with st.spinner("🔬 Расчет биомеханики и построение 3D-модели в стиле CS Inspect..."):
            try:
                client = genai.Client(api_key=api_key)
                
                full_query = SYSTEM_PROMPT + "\n\nДанные ординатора:\n" + doctor_notes
                if extracted_text_from_file:
                    full_query += f"\n\nДанные PPTX:\n{extracted_text_from_file}"

                contents = [full_query]
                if uploaded_file is not None and uploaded_file.name.split('.')[-1].lower() in ["png", "jpg", "jpeg"]:
                    contents.append(Image.open(uploaded_file))

                response = client.models.generate_content(
                    model='gemini-3.8-flash',
                    contents=contents
                )

                st.markdown("---")
                st.success("✅ Анализ выполнен!")
                
                tab1, tab2 = st.tabs(["📋 Клинический протокол", "🎮 3D CS-Inspect Модель"])
                
                with tab1:
                    st.markdown(response.text)
                    st.download_button("📥 Скачать план (.txt)", data=response.text, file_name="plan_3d.txt")

                with tab2:
                    st.markdown("### 🔍 3D Инспект челюсти (Управление как в CS)")
                    st.write("🖱️ **ЛКМ + Перетаскивание:** Вращение | 📜 **Колесико мыши:** Приближение/Зум | 🖱️ **ПКМ:** Смещение")
                    
                    # 3D Движок на Three.js с анимацией Inspect и контролем зума
                    three_js_code = """
                    <div style="position: relative;">
                        <div id="canvas-container" style="width: 100%; height: 500px; background: #0f1319; border-radius: 12px; border: 1px solid #2a3240;"></div>
                        <div style="position: absolute; top: 15px; left: 15px; display: flex; gap: 10px;">
                            <button id="inspectBtn" style="background: #FF4655; color: white; border: none; padding: 8px 15px; border-radius: 6px; cursor: pointer; font-weight: bold;">🔍 Inspect (F-анимация)</button>
                            <button id="resetBtn" style="background: #2a3240; color: white; border: none; padding: 8px 15px; border-radius: 6px; cursor: pointer;">🔄 Сброс камеры</button>
                        </div>
                    </div>

                    <script src="https://cdnjs.cloudflare.com/ajax/libs/three.js/r128/three.min.js"></script>
                    <script src="https://cdn.jsdelivr.net/npm/three@0.128.0/examples/js/controls/OrbitControls.js"></script>
                    <script src="https://cdnjs.cloudflare.com/ajax/libs/tween.js/18.6.4/tween.umd.js"></script>

                    <script>
                        const container = document.getElementById('canvas-container');
                        const scene = new THREE.Scene();
                        scene.background = new THREE.Color(0x0f1319);
                        
                        const camera = new THREE.PerspectiveCamera(45, container.clientWidth / container.clientHeight, 0.1, 1000);
                        camera.position.set(0, 3, 12);
                        
                        const renderer = new THREE.WebGLRenderer({ antialias: true });
                        renderer.setSize(container.clientWidth, container.clientHeight);
                        container.appendChild(renderer.domElement);
                        
                        const controls = new THREE.OrbitControls(camera, renderer.domElement);
                        controls.enableDamping = true;
                        controls.dampingFactor = 0.05;
                        controls.minDistance = 2;   // Максимальный зум вблизи
                        controls.maxDistance = 25;  // Ограничение отдаления

                        // Освещение (студийный свет)
                        scene.add(new THREE.AmbientLight(0xffffff, 0.7));
                        const light1 = new THREE.DirectionalLight(0xffffff, 0.8);
                        light1.position.set(5, 10, 7);
                        scene.add(light1);
                        const light2 = new THREE.DirectionalLight(0x4455ff, 0.4);
                        light2.position.set(-5, -5, -5);
                        scene.add(light2);

                        // Создание 3D модели челюсти с зубами
                        const mainGroup = new THREE.Group();
                        const toothMat = new THREE.MeshStandardMaterial({ color: 0xf0ece1, roughness: 0.15, metalness: 0.1 });
                        const gumMat = new THREE.MeshStandardMaterial({ color: 0xd96b74, roughness: 0.4 });

                        function buildJaw(isUpper) {
                            const jaw = new THREE.Group();
                            const yPos = isUpper ? 0.7 : -0.7;
                            
                            // Десна
                            const gumGeo = new THREE.TorusGeometry(3.2, 0.5, 16, 32, Math.PI * 0.85);
                            const gum = new THREE.Mesh(gumGeo, gumMat);
                            gum.rotation.x = Math.PI / 2 * (isUpper ? 1 : -1);
                            gum.position.y = yPos;
                            jaw.add(gum);

                            // Зубы по дуге
                            const teethCount = 14;
                            for (let i = 0; i < teethCount; i++) {
                                const angle = (i / (teethCount - 1) - 0.5) * Math.PI * 0.85;
                                const radius = 3.2;
                                const x = radius * Math.sin(angle);
                                const z = radius * Math.cos(angle);
                                
                                // Форма зуба
                                const tGeo = new THREE.BoxGeometry(0.5, 0.8, 0.6);
                                const tooth = new THREE.Mesh(tGeo, toothMat);
                                tooth.position.set(x, yPos + (isUpper ? -0.3 : 0.3), z);
                                tooth.rotation.y = angle;
                                jaw.add(tooth);
                            }
                            return jaw;
                        }

                        const upper = buildJaw(true);
                        const lower = buildJaw(false);
                        mainGroup.add(upper);
                        mainGroup.add(lower);
                        scene.add(mainGroup);

                        // CS Inspect Animation
                        let isInspecting = false;
                        document.getElementById('inspectBtn').addEventListener('click', () => {
                            isInspecting = true;
                            new TWEEN.Tween(mainGroup.rotation)
                                .to({ y: mainGroup.rotation.y + Math.PI * 2, x: 0.3 }, 2500)
                                .easing(TWEEN.Easing.Quadratic.Out)
                                .onComplete(() => {
                                    new TWEEN.Tween(mainGroup.rotation)
                                        .to({ x: 0 }, 500)
                                        .start();
                                    isInspecting = false;
                                })
                                .start();
                        });

                        document.getElementById('resetBtn').addEventListener('click', () => {
                            controls.reset();
                            mainGroup.rotation.set(0, 0, 0);
                        });

                        function animate(time) {
                            requestAnimationFrame(animate);
                            TWEEN.update();
                            controls.update();
                            
                            // Легкое покачивание (Idle Animation)
                            if (!isInspecting) {
                                mainGroup.rotation.y += 0.003;
                            }
                            
                            renderer.render(scene, camera);
                        }
                        animate();
                    </script>
                    """
                    components.html(three_js_code, height=530)
                    
                    st.info("🎮 Нажмите кнопку **«Inspect»**, чтобы запустить анимацию полного осмотра челюсти со всех сторон!")

            except Exception as e:
                st.error(f"Ошибка: {e}")
