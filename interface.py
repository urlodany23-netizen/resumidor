import streamlit as st
import json
import os
import time
from concurrent.futures import ThreadPoolExecutor
from google import genai
import pypdf
from docx import Document
import io

# 1. Configuración de página ancha 
st.set_page_config(
    page_title="Resumidor AI - Extractor y Resumen",
    page_icon="📖",
    layout="wide"
)

if "dark_mode" not in st.session_state:
    st.session_state["dark_mode"] = False
if "sound_enabled" not in st.session_state:
    st.session_state["sound_enabled"] = False

@st.fragment
def render_theme_selector():
    columnas_tema = st.columns([3, 1.5, 1])
    with columnas_tema[1]:
        st.toggle(
            "🔔 Aviso sonoro",
            key="sound_enabled",
            help="Actívalo una vez para permitir el sonido cuando termine el resumen.",
        )
    with columnas_tema[2]:
        st.toggle("🌙 Modo Oscuro", key="dark_mode")

    if st.session_state["dark_mode"]:
        theme_variables = """
            --bg-app: #0B0F19;
            --bg-card: #151D2E;
            --border-color: #1E293B;
            --text-primary: #F8FAFC;
            --text-muted: #94A3B8;
            --input-border: #475569;
            --selected-bg: #1E3A8A;
            --hover-bg: #1E293B;
            --radio-bg: #263247;
            --radio-border: #475569;
            --shadow: 0 4px 12px rgba(0, 0, 0, 0.18);
            color-scheme: dark;
        """
    else:
        theme_variables = """
            --bg-app: #F8FAFC;
            --bg-card: #FFFFFF;
            --border-color: #E2E8F0;
            --text-primary: #0F172A;
            --text-muted: #64748B;
            --input-border: #CBD5E1;
            --selected-bg: #EFF6FF;
            --hover-bg: #F1F5F9;
            --radio-bg: #F1F5F9;
            --radio-border: #CBD5E1;
            --shadow: 0 4px 12px rgba(15, 23, 42, 0.05);
            color-scheme: light;
        """

    st.markdown(
        f"<style>:root {{ {theme_variables} --accent: #2563EB; }}</style>",
        unsafe_allow_html=True,
    )


render_theme_selector()

st.markdown(
    f"""
    <style>
    .stApp,
    [data-testid="stAppViewContainer"],
    [data-testid="stMain"],
    [data-testid="stMainBlockContainer"] {{
        background-color: var(--bg-app) !important;
        color: var(--text-primary) !important;
        transition: background-color 0.2s ease, color 0.2s ease;
    }}

    header,
    header[data-testid="stHeader"],
    [data-testid="stHeader"],
    [data-testid="stToolbar"],
    [data-testid="stStatusWidget"],
    [data-testid="stDeployButton"],
    #MainMenu,
    footer {{
        display: none !important;
        visibility: hidden !important;
        height: 0 !important;
    }}

    [data-testid="stMainBlockContainer"] {{
        max-width: 1360px;
        padding: 0.5rem 2rem 3rem;
    }}

    [data-testid="stMarkdownContainer"],
    [data-testid="stMarkdownContainer"] p,
    [data-testid="stMarkdownContainer"] li,
    [data-testid="stWidgetLabel"],
    [data-testid="stWidgetLabel"] *,
    [data-testid="stRadio"] label,
    [data-testid="stToggle"] label,
    [data-testid="stFileUploader"] label,
    [data-testid="stFileUploader"] section,
    [data-testid="stMetric"],
    [data-testid="stMetric"] *,
    [data-testid="stTabs"] button[role="tab"],
    [data-testid="stTabs"] [role="tabpanel"] {{
        color: var(--text-primary) !important;
    }}

    .app-hero {{
        margin: 0 0 1.5rem;
        padding: 1.5rem 1.75rem;
        border: 1px solid var(--border-color);
        border-radius: 12px;
        background: var(--bg-card);
        box-shadow: var(--shadow);
    }}

    .hero-eyebrow {{
        margin: 0 0 0.4rem;
        color: var(--accent) !important;
        font-size: 0.78rem;
        font-weight: 650;
        letter-spacing: 0.06em;
        text-transform: uppercase;
    }}

    .app-hero p:last-child,
    [data-testid="stCaptionContainer"],
    [data-testid="stCaptionContainer"] *,
    [data-testid="stFileUploader"] section small {{
        color: var(--text-muted) !important;
    }}

    .app-hero h1,
    [data-testid="stMarkdownContainer"] h1,
    [data-testid="stMarkdownContainer"] h2,
    [data-testid="stMarkdownContainer"] h3 {{
        color: var(--text-primary) !important;
        font-weight: 700;
        letter-spacing: -0.025em;
    }}

    .app-hero h1 {{
        margin: 0;
        font-size: clamp(1.65rem, 3.5vw, 2.1rem);
        letter-spacing: -0.035em;
    }}

    .app-hero p:last-child {{
        margin: 0.5rem 0 0;
        font-size: 0.98rem;
        line-height: 1.6;
    }}

    [data-testid="column"]:has([data-testid="stTextArea"]),
    [data-testid="column"]:has([data-testid="stTabs"]) {{
        padding: 1.1rem 1.2rem 1rem;
        border: 1px solid var(--border-color);
        border-radius: 12px;
        background: var(--bg-card);
        box-shadow: var(--shadow);
        transition: background-color 0.2s ease, border-color 0.2s ease;
    }}

    [data-testid="stTextArea"] textarea {{
        min-height: 380px;
        background: var(--bg-card) !important;
        color: var(--text-primary) !important;
        border: 1px solid var(--input-border);
        border-radius: 10px;
        box-shadow: none;
        line-height: 1.6;
    }}

    [data-testid="stTextArea"] textarea::placeholder {{
        color: var(--text-muted) !important;
        opacity: 1;
    }}

    [data-testid="stTextArea"] textarea:focus {{
        border-color: var(--accent);
        box-shadow: 0 0 0 3px rgba(37, 99, 235, 0.16);
    }}

    [data-testid="stFileUploader"] section {{
        background: var(--bg-card) !important;
        color: var(--text-primary) !important;
        border: 1px dashed var(--input-border);
        border-radius: 10px;
    }}

    [data-testid="stFileUploader"] section *,
    [data-testid="stFileUploader"] small {{
        color: var(--text-primary) !important;
    }}

    [data-testid="stFileUploader"] button,
    [data-testid="stDownloadButton"] button {{
        color: var(--text-primary) !important;
        background-color: var(--hover-bg) !important;
        border: 1px solid var(--input-border) !important;
        border-radius: 9px;
        min-height: 2.4rem;
        padding: 0.4rem 0.9rem;
    }}

    [data-testid="stFileUploader"] button:hover,
    [data-testid="stDownloadButton"] button:hover {{
        background-color: var(--selected-bg) !important;
        border-color: var(--accent) !important;
    }}

    [data-testid="stFileUploader"] section:hover {{
        border-color: var(--accent);
    }}

    [data-testid="stFileUploader"] button[aria-label="Add files"] {{
        display: none !important;
        visibility: hidden !important;
    }}

    [data-testid="stRadio"] div[role="radiogroup"] {{
        display: flex;
        flex-wrap: wrap;
        gap: 0.4rem;
    }}

    [data-testid="stRadio"] div[role="radiogroup"] label {{
        padding: 0.4rem 0.7rem;
        border: 1px solid var(--border-color);
        border-radius: 999px;
        background: var(--bg-card);
        cursor: pointer;
    }}

    [data-testid="stRadio"] input[type="radio"] {{
        appearance: none !important;
        width: 1rem;
        height: 1rem;
        margin: 0 0.45rem 0 0;
        border: 1.5px solid var(--input-border) !important;
        border-radius: 50%;
        background-color: var(--bg-card) !important;
        box-shadow: none !important;
        vertical-align: -0.15rem;
        cursor: pointer;
    }}

    [data-testid="stRadio"] input[type="radio"]:checked {{
        border-color: var(--accent) !important;
        background-color: var(--accent) !important;
        background-image: radial-gradient(circle, #FFFFFF 0 30%, transparent 34%) !important;
    }}

    [data-testid="stRadio"] input[type="radio"]:focus-visible {{
        outline: 2px solid #60A5FA;
        outline-offset: 2px;
    }}

    [data-testid="stRadioOption"] > div > div:first-child {{
        background-color: var(--radio-bg) !important;
        border: 1px solid var(--radio-border) !important;
        box-shadow: none !important;
    }}

    [data-testid="stRadioOption"][data-selected="true"] > div > div:first-child {{
        background-color: var(--accent) !important;
        border-color: var(--accent) !important;
    }}

    [data-testid="stRadioOption"][data-selected="true"] > div > div:first-child > div {{
        background-color: #FFFFFF !important;
    }}

    [data-testid="stButton"] button[kind="primary"] {{
        color: #FFFFFF !important;
        min-height: 2.8rem;
        background: var(--accent) !important;
        border: 0 !important;
        border-radius: 10px;
        font-weight: 600;
        box-shadow: 0 2px 5px rgba(37, 99, 235, 0.18);
        transition: background 0.15s ease, box-shadow 0.15s ease;
    }}

    [data-testid="stButton"] button[kind="primary"]:hover {{
        color: #FFFFFF !important;
        background: #1D4ED8 !important;
        box-shadow: 0 4px 10px rgba(37, 99, 235, 0.22);
    }}

    [data-testid="stTabs"] button[role="tab"][aria-selected="true"] {{
        color: #60A5FA !important;
    }}

    [data-testid="stMetric"] {{
        color: var(--text-primary);
        background: var(--bg-card);
        border: 1px solid var(--border-color);
        border-radius: 10px;
        padding: 12px 14px;
        box-shadow: var(--shadow);
    }}

    [data-testid="stJson"],
    [data-testid="stCode"] {{
        background: var(--bg-card);
        color: var(--text-primary);
        border-radius: 10px;
    }}

    [data-testid="stAlert"] {{
        border-radius: 10px;
        color: var(--text-primary) !important;
        background: var(--bg-card) !important;
        border: 1px solid var(--border-color);
    }}

    [data-testid="stAlert"] p,
    [data-testid="stAlert"] div {{
        color: var(--text-primary) !important;
    }}

    [data-testid="stDownloadButton"] button {{
        color: #60A5FA !important;
        font-weight: 600;
    }}

    [data-testid="stToggle"] {{
        margin-top: 0.25rem;
    }}

    @media (max-width: 760px) {{
        [data-testid="stMainBlockContainer"] {{
            padding: 0.75rem clamp(0.75rem, 4vw, 1.5rem) 2rem;
        }}

        [data-testid="stHorizontalBlock"] {{
            flex-wrap: wrap;
            gap: 0.75rem;
        }}

        [data-testid="stHorizontalBlock"] > [data-testid="column"] {{
            flex: 1 1 100%;
            min-width: 0;
        }}

        [data-testid="column"]:has([data-testid="stTextArea"]),
        [data-testid="column"]:has([data-testid="stTabs"]) {{
            padding: 0.9rem;
        }}

        [data-testid="stFileUploader"] section {{
            min-width: 0;
        }}
    }}

    @media (max-width: 640px) {{
        [data-testid="stMainBlockContainer"] {{
            padding-right: 0.75rem;
            padding-left: 0.75rem;
        }}

        [data-testid="stTextArea"] textarea {{
            min-height: 300px;
        }}

        .app-hero {{
            padding: 1.25rem;
        }}

        .app-hero h1 {{
            font-size: 1.65rem;
        }}
    }}
    </style>
    """,
    unsafe_allow_html=True,
)

# 2. Cliente de la API
API_KEY = os.getenv("GEMINI_API_KEY")
if not API_KEY:
    st.error("Configura la variable de entorno GEMINI_API_KEY antes de ejecutar la aplicación.")
    st.stop()

client = genai.Client(api_key=API_KEY)

# 3. Funciones de extracción de archivos
def extraer_texto(uploaded_file):
    nombre = uploaded_file.name.lower()
    if nombre.endswith(".txt"):
        return uploaded_file.read().decode("utf-8", errors="ignore")
    elif nombre.endswith(".pdf"):
        lector = pypdf.PdfReader(uploaded_file)
        return "\n".join([page.extract_text() or "" for page in lector.pages])
    elif nombre.endswith(".docx"):
        doc = Document(io.BytesIO(uploaded_file.read()))
        return "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
    return ""

def procesar_ia(texto, modo_resumen):
    import httpx
    from google.genai._gaos.lib.compat_errors import APITimeoutError
    from google.genai._gaos.utils import RetryConfig

    prompt = f"""
Eres un motor de procesamiento de texto estilo Resoomer y extracción de información.
Modo seleccionado: {modo_resumen}

Analiza el siguiente texto y devuelve EXCLUSIVAMENTE un objeto JSON válido con la siguiente estructura:
{{
  "porcentaje_reduccion": "Porcentaje estimado de reducción (ej. 70%)",
  "conteo_palabras_resumen": 0,
  "resumen": "Resumen claro redactado según el modo seleccionado.",
  "ideas_principales": ["Idea clave 1", "Idea clave 2", "Idea clave 3"],
  "entidades_clave": {{
    "personas_o_empresas": ["lista"],
    "fechas": ["lista"],
    "datos_numericos_o_dinero": ["lista"]
  }}
}}

Texto a procesar:
{texto}
"""
    interactions = client.interactions
    interactions.sdk_configuration.retry_config = RetryConfig(
        strategy="none",
        backoff=None,
        retry_connection_errors=False,
    )

    modelos = (
        "gemini-3.8-flash",
        "gemini-3.5-flash-lite",
        "gemini-2.5-flash-lite",
    )
    ultimo_error = None
    for modelo in modelos:
        try:
            interaction = interactions.create(
                model=modelo,
                input=prompt,
                timeout=30,
            )
            break
        except Exception as error:
            if not isinstance(
                error, (httpx.TimeoutException, APITimeoutError)
            ) and getattr(error, "status_code", None) not in (404, 503):
                raise
            ultimo_error = error
    else:
        raise RuntimeError(
            "Todos los modelos de respaldo fallaron por timeout, error 404 o error 503."
        ) from ultimo_error

    salida = getattr(interaction, "output_text", getattr(interaction, "output", str(interaction)))
    salida = salida.strip()
    if "```json" in salida:
        salida = salida.split("```json")[1].split("```")[0].strip()
    elif "```" in salida:
        salida = salida.split("```")[1].split("```")[0].strip()

    return json.loads(salida)

# --- ENCABEZADO  ---
st.markdown(
    """
    <section class="app-hero">
        <p class="hero-eyebrow">Analizador inteligente de documentos</p>
        <h1>📖 Analiza. Resume. Comprende.</h1>
        <p>Extrae las ideas principales y los datos clave de tus documentos en un solo lugar.</p>
    </section>
    """,
    unsafe_allow_html=True,
)

# Barra de herramientas superior (Controles rápidos)
if "modo_resumen" not in st.session_state:
    st.session_state["modo_resumen"] = "Automático"


@st.fragment
def render_mode_selector():
    st.radio(
        "Tipo de resumen:",
        ["Automático", "Ejecutivo / Puntos Clave", "Análisis Detallado"],
        horizontal=True,
        key="modo_resumen",
    )


col_barra1, col_barra2 = st.columns([2, 1])
with col_barra1:
    render_mode_selector()
with col_barra2:
    archivo = st.file_uploader("O importar archivo (.pdf, .docx, .txt)", type=["pdf", "docx", "txt"], label_visibility="collapsed")
modo = st.session_state["modo_resumen"]

# Si el usuario sube un archivo, cargarlo como texto base
texto_inicial = ""
if archivo is not None:
    texto_inicial = extraer_texto(archivo)

# --- LAYOUT DE DOS COLUMNAS (TEXTO ORIGINAL vs RESULTADO) ---
col_izq, col_der = st.columns(2, gap="large")

with col_izq:
    st.markdown("### 📄 Documento Original")
    texto_entrada = st.text_area(
        "Texto de entrada:",
        value=texto_inicial,
        height=380,
        placeholder="Pega aquí tu texto, artículo, contrato o informe para analizar...",
        label_visibility="collapsed"
    )

    palabras_orig = len(texto_entrada.split()) if texto_entrada.strip() else 0
    caracteres_orig = len(texto_entrada)
    st.markdown(
        f'📊 **Conteo:** <span id="live-word-count">{palabras_orig} palabras | '
        f'{caracteres_orig} caracteres</span>',
        unsafe_allow_html=True,
    )
    st.html(
        """
        <script>
        (() => {
            const updateCounter = () => {
                const textarea = document.querySelector(
                    '[data-testid="stTextArea"] textarea'
                );
                const counter = document.getElementById("live-word-count");
                if (!textarea || !counter) return;

                if (textarea.dataset.liveCounterBound !== "true") {
                    textarea.addEventListener("input", () => {
                        const value = textarea.value;
                        const trimmed = value.trim();
                        const words = trimmed ? trimmed.split(/\\s+/u).length : 0;
                        const characters = Array.from(value).length;
                        const display = `${words} palabras | ${characters} caracteres`;
                        if (counter.textContent !== display) {
                            counter.textContent = display;
                        }
                    });
                    textarea.dataset.liveCounterBound = "true";
                }
            };

            if (!window.__liveCounterObserver) {
                window.__liveCounterObserver = new MutationObserver(updateCounter);
                window.__liveCounterObserver.observe(document.body, {
                    childList: true,
                    subtree: true
                });
            }

            if (!window.__summarySoundHooksInstalled) {
                const unlockSummaryAudio = () => {
                    const AudioContextClass =
                        window.AudioContext || window.webkitAudioContext;
                    if (!AudioContextClass) return;

                    if (!window.__summaryAudioContext) {
                        window.__summaryAudioContext = new AudioContextClass();
                    }
                    if (window.__summaryAudioContext.state === "suspended") {
                        window.__summaryAudioContext.resume().catch((error) => {
                            console.warn("No se pudo activar el sonido de finalización.", error);
                        });
                    }
                };
                document.addEventListener("pointerdown", unlockSummaryAudio, true);
                document.addEventListener("touchstart", unlockSummaryAudio, {
                    capture: true,
                    passive: true,
                });
                document.addEventListener("keydown", (event) => {
                    if (event.key === " " || event.key === "Enter") {
                        unlockSummaryAudio(event);
                    }
                }, true);
                window.__summarySoundHooksInstalled = true;
            }

            window.__playSummaryCompletionSound = () => {
                const audioContext = window.__summaryAudioContext;
                if (!audioContext || audioContext.state !== "running") return;

                const notes = [659, 784, 988, 1175];
                const start = audioContext.currentTime + 0.03;
                notes.forEach((frequency, index) => {
                    const noteStart = start + index * 0.2;
                    const oscillator = audioContext.createOscillator();
                    const gain = audioContext.createGain();
                    oscillator.type = "sine";
                    oscillator.frequency.setValueAtTime(frequency, noteStart);
                    gain.gain.setValueAtTime(0.0001, noteStart);
                    gain.gain.exponentialRampToValueAtTime(0.65, noteStart + 0.025);
                    gain.gain.setValueAtTime(0.65, noteStart + 0.12);
                    gain.gain.exponentialRampToValueAtTime(0.0001, noteStart + 0.17);
                    oscillator.connect(gain);
                    gain.connect(audioContext.destination);
                    oscillator.start(noteStart);
                    oscillator.stop(noteStart + 0.18);
                });
            };

            const checkSummaryCompletion = () => {
                const soundToggle = document.querySelector(
                    '[aria-label="🔔 Aviso sonoro"]'
                );
                const completion = [...document.querySelectorAll(
                    '[data-testid="stCaptionContainer"]'
                )].find((element) =>
                    element.textContent.includes("✅ Resumen generado en")
                );

                if (!completion) {
                    window.__lastSummaryCompletion = null;
                    return;
                }

                const completionText = completion.textContent.trim();
                if (
                    soundToggle?.checked &&
                    completionText !== window.__lastSummaryCompletion
                ) {
                    window.__lastSummaryCompletion = completionText;
                    window.__playSummaryCompletionSound();
                }
            };

            if (!window.__summaryCompletionObserver) {
                window.__summaryCompletionObserver = new MutationObserver(
                    checkSummaryCompletion
                );
                window.__summaryCompletionObserver.observe(document.body, {
                    childList: true,
                    subtree: true,
                    characterData: true,
                });
            }

            updateCounter();
            checkSummaryCompletion();
        })();
        </script>
        """,
        unsafe_allow_javascript=True,
    )

    boton_procesar = st.button("⚡ Resumir y Extraer", type="primary", use_container_width=True)

with col_der:
    st.markdown("### 🎯 Síntesis & Extracción")

    if boton_procesar:
        if not texto_entrada.strip():
            st.warning("Escribe o pega texto en la columna izquierda para iniciar.")
        else:
            with st.spinner("Sintetizando y organizando información..."):
                try:
                    with ThreadPoolExecutor(max_workers=1) as executor:
                        tarea = executor.submit(procesar_ia, texto_entrada, modo)
                        indicador_tiempo = st.empty()
                        inicio = time.monotonic()

                        while not tarea.done():
                            segundos = int(time.monotonic() - inicio)
                            tiempo_formateado = (
                                f"{segundos // 60} min {segundos % 60} s"
                            )
                            indicador_tiempo.caption(
                                f"⏱️ Tiempo transcurrido: {tiempo_formateado}"
                            )
                            time.sleep(0.5)

                        resultado = tarea.result()
                        segundos = int(time.monotonic() - inicio)
                        tiempo_formateado = (
                            f"{segundos // 60} min {segundos % 60} s"
                        )
                        indicador_tiempo.caption(
                            f"✅ Resumen generado en {tiempo_formateado}"
                        )
                    # Métricas de reducción arriba del resumen
                    palabras_resumen = len(resultado.get("resumen", "").split())
                    m1, m2 = st.columns(2)
                    m1.metric("Palabras síntesis", f"{palabras_resumen} palabras")
                    m2.metric("Reducción", resultado.get("porcentaje_reduccion", "-"))

                    # Pestañas  para alternar vistas
                    tab_resumen, tab_ideas, tab_json = st.tabs(["📝 Texto Resumido", "📌 Puntos Clave", "🗂️ Datos Estructurados (JSON)"])

                    with tab_resumen:
                        st.write(resultado.get("resumen", ""))

                    with tab_ideas:
                        st.markdown("**Ideas principales destacadas:**")
                        for idea in resultado.get("ideas_principales", []):
                            st.markdown(f"- {idea}")

                    with tab_json:
                        st.markdown("**Entidades detectadas:**")
                        st.json(resultado.get("entidades_clave", {}))
                        
                        # Opción para descargar el JSON completo
                        json_descarga = json.dumps(resultado, indent=4, ensure_ascii=False)
                        st.download_button(
                            label="⬇️ Descargar informe JSON",
                            data=json_descarga,
                            file_name="resumen_estructurado.json",
                            mime="application/json"
                        )

                except Exception as e:
                    st.error(f"Error en el análisis: {e}")
    else:
        st.info("Presiona **'⚡ Resumir y Extraer'** para ver el resultado aquí.")