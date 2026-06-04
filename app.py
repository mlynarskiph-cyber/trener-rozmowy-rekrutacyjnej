import os
import re
import base64
import requests
import streamlit as st
from bs4 import BeautifulSoup
from dotenv import load_dotenv
from openai import OpenAI
from pypdf import PdfReader
from docx import Document

# =============================
# Konfiguracja
# =============================

load_dotenv()

client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))

# Możesz zmienić model np. na "gpt-5.5", jeśli chcesz
MODEL = "gpt-4.1-mini"

st.set_page_config(
    page_title="Trener Rozmowy Rekrutacyjnej",
    page_icon="🎙️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# =============================
# Styl graficzny
# =============================

def inject_css():
    st.markdown("""
    <style>
    .stApp {
        background:
            radial-gradient(circle at top left, rgba(43,108,176,0.18), transparent 28%),
            radial-gradient(circle at top right, rgba(56,189,248,0.15), transparent 22%),
            linear-gradient(180deg, #07111f 0%, #0b1220 45%, #0f172a 100%);
        color: #e5eefb;
    }

    .block-container {
        max-width: 1220px;
        padding-top: 2rem;
        padding-bottom: 2.5rem;
    }

    h1, h2, h3 {
        color: #f8fbff !important;
        letter-spacing: -0.02em;
    }

    .hero-box {
        background: linear-gradient(135deg, rgba(18, 31, 54, 0.95), rgba(16, 24, 39, 0.92));
        border: 1px solid rgba(148, 163, 184, 0.16);
        border-radius: 22px;
        padding: 28px 30px 24px 30px;
        box-shadow: 0 20px 60px rgba(0,0,0,0.28);
        margin-bottom: 1.25rem;
    }

    .hero-kicker {
        color: #93c5fd;
        font-size: 0.92rem;
        font-weight: 700;
        letter-spacing: 0.08em;
        text-transform: uppercase;
        margin-bottom: 0.6rem;
    }

    .hero-title {
        font-size: 2.55rem;
        font-weight: 800;
        line-height: 1.08;
        margin-bottom: 0.45rem;
        color: #f8fbff;
    }

    .hero-subtitle {
        color: #cbd5e1;
        font-size: 1.05rem;
        line-height: 1.7;
        margin-bottom: 0.85rem;
        max-width: 900px;
    }

    .hero-author {
        color: #94a3b8;
        font-size: 0.96rem;
        margin-top: 0.4rem;
    }

    .quote-box {
        background: linear-gradient(135deg, rgba(14,165,233,0.13), rgba(59,130,246,0.11));
        border: 1px solid rgba(125, 211, 252, 0.22);
        border-radius: 18px;
        padding: 16px 18px;
        margin-top: 1rem;
    }

    .quote-text {
        color: #e0f2fe;
        font-size: 1rem;
        line-height: 1.65;
        margin: 0;
    }

    .stepper {
        display: flex;
        gap: 10px;
        flex-wrap: wrap;
        margin: 1rem 0 1.25rem 0;
    }

    .step-pill {
        padding: 10px 14px;
        border-radius: 999px;
        font-size: 0.92rem;
        font-weight: 700;
        border: 1px solid rgba(148, 163, 184, 0.16);
        background: rgba(15, 23, 42, 0.62);
        color: #94a3b8;
    }

    .step-pill.active {
        background: linear-gradient(135deg, rgba(37,99,235,0.95), rgba(14,165,233,0.95));
        color: white;
        border-color: transparent;
        box-shadow: 0 10px 30px rgba(14, 165, 233, 0.18);
    }

    .section-box {
        background: rgba(15, 23, 42, 0.72);
        border: 1px solid rgba(148, 163, 184, 0.14);
        border-radius: 20px;
        padding: 18px 18px 8px 18px;
        box-shadow: 0 10px 40px rgba(0,0,0,0.18);
        margin-bottom: 1rem;
    }

    .section-title {
        font-size: 1.18rem;
        font-weight: 800;
        color: #f8fbff;
        margin-bottom: 0.2rem;
    }

    .section-subtitle {
        color: #94a3b8;
        font-size: 0.95rem;
        margin-bottom: 0.9rem;
        line-height: 1.6;
    }

    .soft-note {
        color: #94a3b8;
        font-size: 0.92rem;
        line-height: 1.65;
    }

    div[data-testid="stSidebar"] {
        background: linear-gradient(180deg, rgba(10, 18, 32, 0.98), rgba(11, 18, 32, 0.98));
        border-right: 1px solid rgba(148, 163, 184, 0.14);
    }

    div[data-testid="stSidebar"] * {
        color: #dbeafe;
    }

    div.stButton > button,
    div[data-testid="stDownloadButton"] > button {
        width: 100%;
        border-radius: 14px;
        border: 0;
        padding: 0.82rem 1rem;
        font-weight: 700;
        font-size: 0.98rem;
        background: linear-gradient(135deg, #2563eb, #0ea5e9);
        color: white;
        box-shadow: 0 10px 30px rgba(37, 99, 235, 0.22);
    }

    div.stButton > button:hover,
    div[data-testid="stDownloadButton"] > button:hover {
        transform: translateY(-1px);
        filter: brightness(1.03);
    }

    .stTextInput input,
    .stTextArea textarea,
    .stSelectbox div[data-baseweb="select"] > div,
    .stFileUploader section,
    .stRadio > div {
        background: rgba(15, 23, 42, 0.78) !important;
        color: #f8fbff !important;
        border-radius: 14px !important;
    }

    .stTextInput label, .stTextArea label, .stRadio label {
        color: #dbeafe !important;
        font-weight: 600;
    }

    div[data-baseweb="radio"] > div label {
        color: #dbeafe !important;
    }

    .stAlert {
        border-radius: 16px;
    }

    .result-box {
        background: rgba(15, 23, 42, 0.78);
        border: 1px solid rgba(148, 163, 184, 0.14);
        border-radius: 18px;
        padding: 18px;
    }

    .question-box {
        background: linear-gradient(135deg, rgba(14,165,233,0.12), rgba(37,99,235,0.12));
        border: 1px solid rgba(125, 211, 252, 0.22);
        border-radius: 18px;
        padding: 16px 18px;
        margin-bottom: 0.8rem;
    }

    .question-label {
        color: #93c5fd;
        font-size: 0.88rem;
        font-weight: 800;
        text-transform: uppercase;
        letter-spacing: 0.08em;
        margin-bottom: 0.4rem;
    }

    .question-text {
        color: #f8fbff;
        font-size: 1.06rem;
        line-height: 1.65;
        margin: 0;
    }

    .progress-caption {
        color: #94a3b8;
        font-size: 0.94rem;
        margin-bottom: 0.35rem;
    }

    hr {
        border-color: rgba(148, 163, 184, 0.18);
    }
    </style>
    """, unsafe_allow_html=True)


# =============================
# UI helpers
# =============================

def render_sidebar():
    st.sidebar.markdown("## 🎙️ Trener Rozmowy Rekrutacyjnej")
    st.sidebar.caption("Autor: Robert Młynarski")

    st.sidebar.markdown("---")
    st.sidebar.markdown("### Jak działa aplikacja?")
    st.sidebar.markdown("""
1. **Wgraj CV**
2. **Dodaj ogłoszenie**
3. **Przeczytaj analizę**
4. **Przejdź symulację rozmowy**
5. **Odbierz feedback**
""")

    st.sidebar.markdown("---")
    st.sidebar.markdown("### Co możesz dodać?")
    st.sidebar.markdown("""
- tekst ogłoszenia,
- link do ogłoszenia,
- zrzut ekranu ogłoszenia,
- kontekst dotyczący studiów i dyspozycyjności.
""")

    st.sidebar.markdown("---")
    st.sidebar.info(
        "Aplikacja ma pomagać w przygotowaniu do rozmowy — nie zastępuje rekrutera, "
        "ale porządkuje ryzyka, pytania i obszary do poprawy."
    )


def render_hero():
    st.markdown("""
    <div class="hero-box">
        <div class="hero-kicker">AI Career Support</div>
        <div class="hero-title">🎙️ Trener Rozmowy Rekrutacyjnej</div>
        <div class="hero-subtitle">
            Przygotuj się do rozmowy kwalifikacyjnej bez zgadywania.
            Aplikacja analizuje CV i ofertę pracy, generuje realistyczne pytania,
            a następnie daje konkretny feedback po symulacji rozmowy.
        </div>
        <div class="hero-author">Autor: Robert Młynarski</div>
        <div class="quote-box">
            <p class="quote-text">
                💡 Pierwsza rozmowa nie wymaga perfekcji. Wymaga przygotowania,
                szczerości i odwagi, żeby zrobić pierwszy krok.
            </p>
        </div>
    </div>
    """, unsafe_allow_html=True)


def render_stepper(current_step):
    labels = [
        ("input", "1. Dane wejściowe"),
        ("analysis", "2. Analiza"),
        ("interview", "3. Symulacja"),
        ("feedback", "4. Feedback"),
    ]

    html = '<div class="stepper">'
    current_index = [k for k, _ in labels].index(current_step)

    for idx, (key, label) in enumerate(labels):
        active = "active" if idx <= current_index else ""
        html += f'<div class="step-pill {active}">{label}</div>'
    html += "</div>"
    st.markdown(html, unsafe_allow_html=True)


def section_header(icon, title, subtitle=""):
    st.markdown(f"""
    <div class="section-title">{icon} {title}</div>
    <div class="section-subtitle">{subtitle}</div>
    """, unsafe_allow_html=True)


# =============================
# Funkcje do odczytu CV
# =============================

def extract_text_from_pdf(uploaded_file):
    reader = PdfReader(uploaded_file)
    text = ""

    for page in reader.pages:
        page_text = page.extract_text()
        if page_text:
            text += page_text + "\n"

    return text.strip()


def extract_text_from_docx(uploaded_file):
    doc = Document(uploaded_file)
    paragraphs = [p.text for p in doc.paragraphs]
    return "\n".join(paragraphs).strip()


def extract_text_from_file(uploaded_file):
    if uploaded_file is None:
        return ""

    filename = uploaded_file.name.lower()

    if filename.endswith(".pdf"):
        return extract_text_from_pdf(uploaded_file)

    if filename.endswith(".docx"):
        return extract_text_from_docx(uploaded_file)

    if filename.endswith(".txt"):
        return uploaded_file.read().decode("utf-8")

    return ""


# =============================
# Funkcje do ogłoszeń
# =============================

def clean_web_text(text):
    lines = text.splitlines()
    cleaned_lines = []

    for line in lines:
        line = line.strip()
        line = re.sub(r"\s+", " ", line)

        if len(line) > 2:
            cleaned_lines.append(line)

    cleaned_text = "\n".join(cleaned_lines)
    cleaned_text = re.sub(r"\n{3,}", "\n\n", cleaned_text)

    return cleaned_text.strip()


def fetch_text_from_url(url):
    if not url.startswith("http://") and not url.startswith("https://"):
        url = "https://" + url

    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/120.0 Safari/537.36"
        )
    }

    response = requests.get(url, headers=headers, timeout=15)
    response.raise_for_status()

    soup = BeautifulSoup(response.text, "html.parser")

    for tag in soup(["script", "style", "noscript", "header", "footer", "svg"]):
        tag.decompose()

    text = soup.get_text(separator="\n")
    return clean_web_text(text)


def image_to_data_url(uploaded_image):
    image_bytes = uploaded_image.getvalue()
    mime_type = uploaded_image.type

    if not mime_type:
        filename = uploaded_image.name.lower()

        if filename.endswith(".png"):
            mime_type = "image/png"
        elif filename.endswith(".jpg") or filename.endswith(".jpeg"):
            mime_type = "image/jpeg"
        elif filename.endswith(".webp"):
            mime_type = "image/webp"
        else:
            mime_type = "image/png"

    encoded = base64.b64encode(image_bytes).decode("utf-8")
    return f"data:{mime_type};base64,{encoded}"


def extract_job_text_from_screenshot(uploaded_image):
    image_data_url = image_to_data_url(uploaded_image)

    prompt = """
Odczytaj treść ogłoszenia o pracę ze zrzutu ekranu.

Zasady:
- Przepisz tylko informacje związane z ogłoszeniem.
- Usuń elementy interfejsu strony, reklamy, menu, stopki i przypadkowe teksty.
- Zachowaj: nazwę stanowiska, firmę, lokalizację, obowiązki, wymagania, oferowane warunki, tryb pracy, wynagrodzenie, jeśli jest widoczne.
- Jeżeli jakiś fragment jest niewyraźny, napisz: "[fragment niewyraźny]".
- Nie dopowiadaj niczego, czego nie widać na obrazie.
- Odpowiedz po polsku.
"""

    response = client.responses.create(
        model=MODEL,
        input=[
            {
                "role": "user",
                "content": [
                    {
                        "type": "input_text",
                        "text": prompt
                    },
                    {
                        "type": "input_image",
                        "image_url": image_data_url
                    }
                ]
            }
        ]
    )

    return response.output_text.strip()


# =============================
# Funkcje pomocnicze
# =============================

def clear_answer_widgets():
    keys_to_delete = []

    for key in list(st.session_state.keys()):
        if str(key).startswith("answer_"):
            keys_to_delete.append(key)

    for key in keys_to_delete:
        del st.session_state[key]


def build_markdown_report(title, content):
    return f"""# {title}

{content}
"""


# =============================
# Funkcje AI
# =============================

def ask_ai(prompt):
    response = client.responses.create(
        model=MODEL,
        input=prompt
    )
    return response.output_text


def analyze_cv_and_job(cv_text, job_text, extra_context):
    prompt = f"""
Jesteś trenerem rozmów kwalifikacyjnych dla młodej osoby szukającej pierwszej pracy.

Twoim zadaniem jest przeanalizować CV, ogłoszenie o pracę i dodatkowy kontekst.

Bardzo ważne zasady:
- Nie wymyślaj faktów o kandydacie.
- Jeżeli czegoś nie wiesz, napisz: "nie wiadomo".
- Nie pisz gotowych odpowiedzi do nauczenia się na pamięć.
- Skup się na przygotowaniu do realnej rozmowy kwalifikacyjnej.
- Uwzględnij, że kandydat może nie mieć doświadczenia zawodowego.
- Feedback ma być konkretny, praktyczny i wspierający.
- Jeżeli kandydat ma braki względem ogłoszenia, nazwij je uczciwie, ale nie zniechęcaj.
- Oddziel fakty z CV od wniosków i hipotez.
- Jeśli ogłoszenie zostało odczytane z linku albo screenshotu i wygląda na niepełne, zaznacz to.

CV KANDYDATA:
{cv_text}

OGŁOSZENIE O PRACĘ:
{job_text}

DODATKOWY KONTEKST OD UŻYTKOWNIKA:
{extra_context}

Przygotuj analizę w języku polskim według tej struktury:

## 1. Jak rozumiem stanowisko
Napisz krótko, o jaką pracę chodzi i jaki typ rozmowy może się odbyć.

## 2. Najważniejsze wymagania z ogłoszenia
Wypunktuj najważniejsze wymagania.

## 3. Mocne strony kandydata pod tę ofertę
Wskaż, co z CV może pomóc kandydatowi.

## 4. Potencjalne ryzyka na rozmowie
Wskaż trudne tematy, które rekruter może poruszyć.

## 5. Pytania, które prawdopodobnie padną
Podaj 8 realistycznych pytań rekrutacyjnych.

## 6. Tematy, które kandydat powinien przemyśleć przed rozmową
Podaj konkretne rzeczy do przygotowania.

## 7. Czy CV wymaga dopasowania do tej oferty?
Oceń krótko, czy CV wygląda na dobrze dopasowane do tej konkretnej oferty.
"""
    return ask_ai(prompt)


def generate_interview_questions(cv_text, job_text, extra_context, analysis):
    prompt = f"""
Jesteś rekruterem prowadzącym realistyczną rozmowę kwalifikacyjną.

Przygotuj 10 pytań do symulacji rozmowy na podstawie CV, ogłoszenia, kontekstu i analizy.

Zasady:
- Kandydat szuka jednej z pierwszych prac.
- Pytania mają być realistyczne, nie przesadnie korporacyjne.
- Uwzględnij trudne miejsca z analizy.
- Nie podawaj odpowiedzi, tylko pytania.
- Każde pytanie ma być osobne.
- Zadaj pytania o motywację, brak doświadczenia, dyspozycyjność, dopasowanie do stanowiska, obsługę klienta i trudne sytuacje.
- Jeżeli ogłoszenie dotyczy sprzedaży, dodaj pytanie sytuacyjne o klienta.
- Jeżeli ogłoszenie dotyczy banku/finansów, dodaj pytanie o dokładność, procedury i odpowiedzialność.
- Jeżeli widać ryzyko konfliktu studiów z pracą, dodaj pytanie o dyspozycyjność.
- Nie wymyślaj faktów o kandydacie.

CV:
{cv_text}

OGŁOSZENIE:
{job_text}

DODATKOWY KONTEKST:
{extra_context}

ANALIZA:
{analysis}

Zwróć dokładnie 10 pytań w takim formacie:

Pytanie 1: ...
Pytanie 2: ...
Pytanie 3: ...
Pytanie 4: ...
Pytanie 5: ...
Pytanie 6: ...
Pytanie 7: ...
Pytanie 8: ...
Pytanie 9: ...
Pytanie 10: ...
"""

    raw_questions = ask_ai(prompt)
    questions = []

    for line in raw_questions.splitlines():
        line = line.strip()
        match = re.match(r"^Pytanie\s+\d+:\s*(.*)", line)
        if match:
            questions.append(match.group(1).strip())

    if not questions:
        questions = [
            "Proszę opowiedzieć coś o sobie.",
            "Dlaczego zainteresowała Pana ta oferta?",
            "Co wie Pan o tej firmie i tym stanowisku?",
            "Nie ma Pan jeszcze dużego doświadczenia w tej branży. Dlaczego uważa Pan, że poradzi sobie Pan na tym stanowisku?",
            "Jakie swoje mocne strony może Pan wykorzystać w tej pracy?",
            "Proszę opowiedzieć o sytuacji, w której musiał Pan obsłużyć wymagającego klienta lub gościa.",
            "Jak radzi sobie Pan z dokładnością i pracą według procedur?",
            "Jak wygląda Pana dyspozycyjność w kontekście planowanych studiów?",
            "Co zrobi Pan, jeśli klient będzie niezadowolony albo zdenerwowany?",
            "Czy ma Pan pytania do nas?"
        ]

    return questions[:10]


def generate_final_feedback(cv_text, job_text, extra_context, analysis, answers):
    transcript = ""

    for i, item in enumerate(answers, start=1):
        transcript += f"""
PYTANIE {i}:
{item["question"]}

ODPOWIEDŹ KANDYDATA:
{item["answer"]}
"""

    prompt = f"""
Jesteś trenerem rozmów kwalifikacyjnych dla młodych osób szukających pierwszej pracy.

Oceń całą rozmowę kandydata na podstawie CV, ogłoszenia, analizy i transkryptu.

Zasady:
- Nie oceniaj osoby, oceniaj odpowiedzi.
- Nie sugeruj kłamstwa.
- Nie wymyślaj faktów o kandydacie.
- Feedback ma być szczery, konkretny i wspierający.
- Wskaż maksymalnie 3 najważniejsze rzeczy do poprawy.
- Nie pisz ogólników typu "bądź bardziej konkretny" bez wyjaśnienia.
- Jeżeli odpowiedź była zbyt ogólna, pokaż, czego zabrakło.
- Jeżeli odpowiedź była dobra, wyjaśnij dlaczego.
- Nie dawaj gotowych formułek do nauczenia się na pamięć.
- Dawaj przykłady kierunku poprawy, ale zaznaczaj, że kandydat ma mówić własnymi słowami.
- Uwzględnij, że to może być jedna z pierwszych rozmów kwalifikacyjnych w życiu kandydata.
- Feedback ma budować pewność siebie, ale nie ukrywać realnych ryzyk.

CV:
{cv_text}

OGŁOSZENIE:
{job_text}

DODATKOWY KONTEKST:
{extra_context}

ANALIZA PRZED ROZMOWĄ:
{analysis}

TRANSKRYPT ROZMOWY:
{transcript}

Przygotuj feedback w języku polskim według struktury:

## 1. Ogólna ocena rozmowy

## 2. Co kandydat zrobił dobrze

## 3. Największe ryzyka na prawdziwej rozmowie

## 4. Trzy najważniejsze rzeczy do poprawy

## 5. Najlepsza odpowiedź i dlaczego

## 6. Najsłabsza odpowiedź i dlaczego

## 7. Jak poprawić odpowiedzi bez uczenia się ich na pamięć

## 8. Krótki plan ćwiczeń przed prawdziwą rozmową

## 9. Gotowość do rozmowy — ocena od 1 do 10

Podaj ocenę i krótko ją uzasadnij.
"""
    return ask_ai(prompt)


# =============================
# Stan aplikacji
# =============================

if "step" not in st.session_state:
    st.session_state.step = "input"

if "cv_text" not in st.session_state:
    st.session_state.cv_text = ""

if "job_text" not in st.session_state:
    st.session_state.job_text = ""

if "extra_context" not in st.session_state:
    st.session_state.extra_context = ""

if "analysis" not in st.session_state:
    st.session_state.analysis = ""

if "questions" not in st.session_state:
    st.session_state.questions = []

if "current_question" not in st.session_state:
    st.session_state.current_question = 0

if "answers" not in st.session_state:
    st.session_state.answers = []

if "feedback" not in st.session_state:
    st.session_state.feedback = ""


# =============================
# Start UI
# =============================

inject_css()
render_sidebar()
render_hero()
render_stepper(st.session_state.step)

# =============================
# Ekran 1: Dane wejściowe
# =============================

if st.session_state.step == "input":
    left_col, right_col = st.columns([1.35, 0.95], gap="large")

    with left_col:
        st.markdown('<div class="section-box">', unsafe_allow_html=True)
        section_header(
            "📄",
            "CV kandydata",
            "Wgraj CV w formacie PDF, DOCX albo TXT."
        )
        uploaded_cv = st.file_uploader(
            "Wgraj CV",
            type=["pdf", "docx", "txt"],
            label_visibility="collapsed"
        )
        st.markdown("</div>", unsafe_allow_html=True)

        st.markdown('<div class="section-box">', unsafe_allow_html=True)
        section_header(
            "💼",
            "Ogłoszenie o pracę",
            "Możesz wkleić treść, dodać link albo wgrać zrzut ekranu."
        )

        job_input_method = st.radio(
            "Wybierz sposób dodania ogłoszenia:",
            [
                "Wkleję treść ogłoszenia",
                "Wkleję link do ogłoszenia",
                "Wgram zrzut ekranu ogłoszenia"
            ]
        )

        job_text_manual = ""
        job_url = ""
        uploaded_job_image = None

        if job_input_method == "Wkleję treść ogłoszenia":
            job_text_manual = st.text_area(
                "Treść ogłoszenia",
                height=260,
                placeholder="Wklej tutaj treść ogłoszenia..."
            )

        elif job_input_method == "Wkleję link do ogłoszenia":
            job_url = st.text_input(
                "Link do ogłoszenia",
                placeholder="https://..."
            )
            st.info(
                "Aplikacja spróbuje pobrać treść z linku. "
                "Jeśli portal blokuje pobieranie, użyj tekstu albo zrzutu ekranu."
            )

        elif job_input_method == "Wgram zrzut ekranu ogłoszenia":
            uploaded_job_image = st.file_uploader(
                "Wgraj zrzut ekranu ogłoszenia",
                type=["png", "jpg", "jpeg", "webp"]
            )

            if uploaded_job_image is not None:
                st.image(
                    uploaded_job_image,
                    caption="Podgląd zrzutu ekranu",
                    use_container_width=True
                )

        st.markdown("</div>", unsafe_allow_html=True)

    with right_col:
        st.markdown('<div class="section-box">', unsafe_allow_html=True)
        section_header(
            "🧠",
            "Kontekst rozmowy",
            "Dodaj informacje, które pomogą AI lepiej ocenić sytuację kandydata."
        )

        planned_studies = st.text_input(
            "Czy kandydat planuje studia? Jeśli tak, jakie i w jakim trybie?"
        )

        availability = st.text_input(
            "Jaka jest realna dyspozycyjność? Dni, godziny, weekendy?"
        )

        concerns = st.text_area(
            "Czego kandydat najbardziej obawia się na rozmowie?",
            height=140
        )

        st.markdown(
            '<p class="soft-note">Im bardziej konkretny kontekst podasz, tym lepsza będzie analiza i pytania rekrutacyjne.</p>',
            unsafe_allow_html=True
        )

        analyze_clicked = st.button("Przeanalizuj CV i ogłoszenie")
        st.markdown("</div>", unsafe_allow_html=True)

    if analyze_clicked:
        cv_text = extract_text_from_file(uploaded_cv)

        if not cv_text:
            st.error("Nie udało się odczytać CV. Wgraj plik PDF, DOCX albo TXT.")
            st.stop()

        job_text = ""

        if job_input_method == "Wkleję treść ogłoszenia":
            job_text = job_text_manual.strip()

            if not job_text:
                st.error("Wklej treść ogłoszenia.")
                st.stop()

        elif job_input_method == "Wkleję link do ogłoszenia":
            if not job_url.strip():
                st.error("Wklej link do ogłoszenia.")
                st.stop()

            with st.spinner("Pobieram treść ogłoszenia z linku..."):
                try:
                    job_text = fetch_text_from_url(job_url.strip())
                except Exception as e:
                    st.error(
                        "Nie udało się pobrać treści z linku. "
                        "Portal może blokować automatyczne pobieranie. "
                        "Spróbuj wkleić treść ręcznie albo użyć zrzutu ekranu."
                    )
                    st.caption(f"Szczegóły techniczne: {e}")
                    st.stop()

            if len(job_text) < 300:
                st.warning(
                    "Pobrana treść jest dość krótka. Możliwe, że portal nie udostępnił pełnego ogłoszenia."
                )

        elif job_input_method == "Wgram zrzut ekranu ogłoszenia":
            if uploaded_job_image is None:
                st.error("Wgraj zrzut ekranu ogłoszenia.")
                st.stop()

            with st.spinner("Odczytuję treść ogłoszenia ze zrzutu ekranu..."):
                job_text = extract_job_text_from_screenshot(uploaded_job_image)

            if len(job_text) < 100:
                st.warning(
                    "Odczytana treść jest krótka. Możliwe, że zrzut ekranu nie zawiera całego ogłoszenia."
                )

        extra_context = f"""
Planowane studia: {planned_studies}
Dyspozycyjność: {availability}
Obawy przed rozmową: {concerns}
Sposób dodania ogłoszenia: {job_input_method}
"""

        st.session_state.cv_text = cv_text
        st.session_state.job_text = job_text
        st.session_state.extra_context = extra_context

        with st.spinner("Analizuję CV i ogłoszenie..."):
            st.session_state.analysis = analyze_cv_and_job(
                cv_text,
                job_text,
                extra_context
            )

        st.session_state.step = "analysis"
        st.rerun()


# =============================
# Ekran 2: Analiza
# =============================

elif st.session_state.step == "analysis":
    col1, col2 = st.columns([1.15, 0.85], gap="large")

    with col1:
        st.markdown('<div class="section-box">', unsafe_allow_html=True)
        section_header(
            "📑",
            "Odczytana treść ogłoszenia",
            "To właśnie tę treść wykorzystuje AI do przygotowania analizy."
        )

        with st.expander("Pokaż / ukryj treść ogłoszenia"):
            st.text_area(
                "Treść ogłoszenia",
                value=st.session_state.job_text,
                height=320
            )
        st.markdown("</div>", unsafe_allow_html=True)

    with col2:
        st.markdown('<div class="section-box">', unsafe_allow_html=True)
        section_header(
            "⚡ Akcje",
            "Możesz przejść do symulacji albo wrócić i poprawić dane wejściowe."
        )

        analysis_file = build_markdown_report(
            "Analiza przed rozmową kwalifikacyjną",
            st.session_state.analysis
        )

        st.download_button(
            label="Pobierz analizę jako Markdown",
            data=analysis_file,
            file_name="analiza_przed_rozmowa.md",
            mime="text/markdown"
        )

        start_simulation = st.button("Rozpocznij symulację rozmowy")
        go_back = st.button("Wróć i popraw dane")
        st.markdown("</div>", unsafe_allow_html=True)

        if start_simulation:
            with st.spinner("Przygotowuję pytania rekrutera..."):
                st.session_state.questions = generate_interview_questions(
                    st.session_state.cv_text,
                    st.session_state.job_text,
                    st.session_state.extra_context,
                    st.session_state.analysis
                )

            clear_answer_widgets()
            st.session_state.step = "interview"
            st.session_state.current_question = 0
            st.session_state.answers = []
            st.rerun()

        if go_back:
            st.session_state.step = "input"
            st.rerun()

    st.markdown('<div class="section-box">', unsafe_allow_html=True)
    section_header(
        "🧾",
        "Analiza przed rozmową",
        "Poniżej znajdziesz ocenę dopasowania, ryzyk i najważniejszych tematów do przygotowania."
    )

    st.markdown('<div class="result-box">', unsafe_allow_html=True)
    st.markdown(st.session_state.analysis)
    st.markdown("</div>", unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)


# =============================
# Ekran 3: Rozmowa
# =============================

elif st.session_state.step == "interview":
    questions = st.session_state.questions
    current = st.session_state.current_question

    st.markdown('<div class="section-box">', unsafe_allow_html=True)
    section_header(
        "🎤",
        "Symulacja rozmowy",
        "Odpowiadaj własnymi słowami. Nie chodzi o perfekcję, tylko o ćwiczenie rozmowy."
    )

    if current < len(questions):
        question = questions[current]
        progress_value = current / len(questions)

        st.markdown(
            f'<div class="progress-caption">Postęp: pytanie {current + 1} z {len(questions)}</div>',
            unsafe_allow_html=True
        )
        st.progress(progress_value)

        st.markdown(f"""
        <div class="question-box">
            <div class="question-label">Pytanie rekrutera</div>
            <p class="question-text">{question}</p>
        </div>
        """, unsafe_allow_html=True)

        answer = st.text_area(
            "Twoja odpowiedź:",
            height=220,
            key=f"answer_{current}",
            placeholder="Wpisz tutaj swoją odpowiedź..."
        )

        col_a, col_b = st.columns(2, gap="large")
        with col_a:
            next_clicked = st.button("Zapisz odpowiedź i przejdź dalej")
        with col_b:
            finish_early = st.button("Zakończ wcześniej i przejdź do feedbacku")

        if next_clicked:
            if not answer.strip():
                st.warning("Wpisz odpowiedź przed przejściem dalej.")
                st.stop()

            st.session_state.answers.append({
                "question": question,
                "answer": answer
            })

            st.session_state.current_question += 1
            st.rerun()

        if finish_early:
            if len(st.session_state.answers) == 0 and not answer.strip():
                st.warning("Najpierw odpowiedz przynajmniej na jedno pytanie.")
                st.stop()

            if answer.strip():
                st.session_state.answers.append({
                    "question": question,
                    "answer": answer
                })

            with st.spinner("Analizuję dotychczasową rozmowę..."):
                st.session_state.feedback = generate_final_feedback(
                    st.session_state.cv_text,
                    st.session_state.job_text,
                    st.session_state.extra_context,
                    st.session_state.analysis,
                    st.session_state.answers
                )

            st.session_state.step = "feedback"
            st.rerun()

    else:
        st.success("Rozmowa została zakończona. Możesz teraz wygenerować feedback końcowy.")
        if st.button("Wygeneruj feedback końcowy"):
            with st.spinner("Analizuję całą rozmowę..."):
                st.session_state.feedback = generate_final_feedback(
                    st.session_state.cv_text,
                    st.session_state.job_text,
                    st.session_state.extra_context,
                    st.session_state.analysis,
                    st.session_state.answers
                )

            st.session_state.step = "feedback"
            st.rerun()

    st.markdown("</div>", unsafe_allow_html=True)


# =============================
# Ekran 4: Feedback
# =============================

elif st.session_state.step == "feedback":
    top_left, top_right = st.columns([1.15, 0.85], gap="large")

    with top_left:
        st.markdown('<div class="section-box">', unsafe_allow_html=True)
        section_header(
            "✅",
            "Feedback końcowy",
            "Ocena dotyczy odpowiedzi, a nie osoby. Potraktuj ją jako plan dalszego treningu."
        )

        st.markdown('<div class="result-box">', unsafe_allow_html=True)
        st.markdown(st.session_state.feedback)
        st.markdown("</div>", unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)

    with top_right:
        st.markdown('<div class="section-box">', unsafe_allow_html=True)
        section_header(
            "⚙️",
            "Co dalej?",
            "Możesz pobrać feedback, powtórzyć rozmowę albo zacząć nową analizę."
        )

        feedback_file = build_markdown_report(
            "Feedback po symulacji rozmowy kwalifikacyjnej",
            st.session_state.feedback
        )

        st.download_button(
            label="Pobierz feedback jako Markdown",
            data=feedback_file,
            file_name="feedback_rozmowa.md",
            mime="text/markdown"
        )

        repeat_interview = st.button("Powtórz rozmowę z tym samym ogłoszeniem")
        start_over = st.button("Zacznij od nowa")
        st.markdown("</div>", unsafe_allow_html=True)

        if repeat_interview:
            clear_answer_widgets()
            st.session_state.step = "interview"
            st.session_state.current_question = 0
            st.session_state.answers = []
            st.rerun()

        if start_over:
            for key in list(st.session_state.keys()):
                del st.session_state[key]
            st.rerun()