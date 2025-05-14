import re
import os
from io import BytesIO

import streamlit as st
import openai
from openpyxl import load_workbook

# ================== Налаштування ==================
# Встановіть свій ключ OpenAI через змінну середовища (OPENAI_API_KEY)
openai.api_key = os.getenv(
    "OPENAI_API_KEY"
)

# ================== Налаштування сторінки ==================
st.set_page_config(page_title="Генератор тестів для Moodle (XML)", layout="wide")
st.markdown("""
<style>
textarea, input[type=text] {
    border-radius: 8px;
    padding: 8px;
    font-size: 1rem;
}
div[data-testid="column"] > div {
    background-color: #f9f9f9;
    padding: 1rem;
    margin-bottom: 1rem;
    border-radius: 8px;
    box-shadow: 0 0 10px rgba(0,0,0,0.1);
}
</style>
""", unsafe_allow_html=True)
st.title("Генератор тестів для Moodle (XML)")

with st.expander("📖 Інструкція до роботи з програмою", expanded=False):
    st.markdown("""
**Excel-режим**
- Підготуйте Excel-файл (.xlsx).
- Колонка A: текст питання.
- Колонка A: варіанти відповідей.
- Позначте правильні відповіді жовтим фоном (FFFF00).

**GPT-режим**
- Вставте текст українською мовою.
- Система створить 10 питань:
  - 4–5 Single-choice,
  - 2–3 True/False,
  - 2–3 Multiple-choice.
- Для кожного, крім True/False, 4 варіанти A–D.

**Ручне створення**
- Оберіть тип питання: Single-choice, Multiple-choice, True/False.
- Введіть текст питання та варіанти.
- Позначте правильні відповіді.

**Готовий тест**
- Формат:

1. Питання...
A. Відповідь A  
B. Відповідь B  
C. Відповідь C  
D. Відповідь D  
Правильний відповідь: B

- Для True/False:

7. Питання...  
Варіанти: True / False  
Правильний відповідь: True

- Обов'язкові: нумерація, варіанти, рядок з `Правильний відповідь:`.
""", unsafe_allow_html=False)

# ================== Утиліти для XML ==================

def wrap_cdata(text: str) -> str:
    """Обгортає текст у CDATA з HTML-тегом <p>."""
    return f"<![CDATA[<p>{text}</p>]]>"

def detect_question_type(answers):
    """Визначаємо тип питання за списком відповідей."""
    if all('-' in ans for ans, _ in answers):
        return "matching"
    correct = sum(1 for _, c in answers if c)
    if len(answers) == 2 and correct <= 1:
        return "truefalse"
    if correct == 1:
        return "single"
    if correct > 1:
        return "multiple"
    return "unknown"

def generate_moodle_xml_string(questions) -> str:
    """Генерує Moodle XML зі списку питань."""
    lines = ['<?xml version="1.0" encoding="UTF-8"?>', '<quiz>']
    for q in questions:
        q_type = detect_question_type(q["answers"])
        if q_type in ("single", "multiple"):
            lines.append('  <question type="multichoice">')
        elif q_type == "truefalse":
            lines.append('  <question type="truefalse">')
        elif q_type == "matching":
            lines.append('  <question type="matching">')
        else:
            continue

        preview = q["text"][:30] + ('...' if len(q["text"]) > 30 else '')
        lines.append('    <name>')
        lines.append(f'      <text>{wrap_cdata(preview)}</text>')
        lines.append('    </name>')

        lines.append('    <questiontext format="html">')
        lines.append(f'      <text>{wrap_cdata(q["text"])}</text>')
        lines.append('    </questiontext>')

        if q_type in ("single", "multiple"):
            total_correct = sum(1 for _, c in q["answers"] if c)
            penalty = 1.0 / total_correct if total_correct else 0
            lines.append('    <shuffleanswers>true</shuffleanswers>')
            lines.append(f'    <single>{"true" if q_type=="single" else "false"}</single>')
            lines.append('    <answernumbering>abc</answernumbering>')
            lines.append(f'    <penalty>{penalty:.6f}</penalty>')
            for text, corr in q["answers"]:
                frac = 100 if corr else 0
                lines.append(f'    <answer fraction="{frac}" format="html">')
                lines.append(f'      <text><![CDATA[{text}]]></text>')
                lines.append('    </answer>')
        elif q_type == "truefalse":
            correct_true = q["answers"][0][1]
            for val in ("true", "false"):
                frac = 100 if (val=="true" and correct_true) or (val=="false" and not correct_true) else 0
                lines.append(f'    <answer fraction="{frac}" format="html">')
                lines.append(f'      <text><![CDATA[{val}]]></text>')
                lines.append('    </answer>')
        elif q_type == "matching":
            lines.append('    <shuffleanswers>true</shuffleanswers>')
            for pair, _ in q["answers"]:
                left, right = map(str.strip, pair.split('-', 1))
                lines.append('    <subquestion format="html">')
                lines.append(f'      <text><![CDATA[{left}]]></text>')
                lines.append(f'      <answer><![CDATA[{right}]]></answer>')
                lines.append('    </subquestion>')

        lines.append('  </question>')
    lines.append('</quiz>')
    return "\n".join(lines)

def download_xml(data, filename: str):
    """Створює кнопку для завантаження XML-файлу."""
    xml_bytes = data.encode('utf-8') if isinstance(data, str) else data
    st.download_button(
        label="📥 Завантажити XML",
        data=xml_bytes,
        file_name=filename,
        mime="application/xml"
    )

# ================== Парсери ==================

def parse_text_format(text: str):
    """Парсер тестів з тексту з нумерацією та правильними відповідями."""
    mapping = {'а':'A','А':'A','б':'B','Б':'B','в':'C','В':'C','г':'D','Г':'D'}
    blocks = re.split(r"(?m)(?=^\d+\.)", text.strip())
    blocks = [blk for blk in blocks if blk.strip()]
    corr_pattern = re.compile(r"(?i)^(?:правильн\w*\s+(?:ответ|відповід))[:]?", re.IGNORECASE)
    questions, errors = [], []

    for idx, block in enumerate(blocks, 1):
        lines = [ln.strip() for ln in block.splitlines() if ln.strip()]
        m_q = re.match(r"^\d+\.\s*(.+)", lines[0])
        q_text = m_q.group(1).strip() if m_q else lines[0]

        corr_idx = next((i for i, ln in enumerate(lines) if corr_pattern.match(ln)), None)
        if corr_idx is None:
            errors.append((idx, "Не знайдено рядок із правильними відповідями"))
            continue

        ans_lines = lines[1:corr_idx]
        corr_line = lines[corr_idx]
        letters = re.findall(r"[A-ГA-D]", corr_line)
        correct_set = {mapping.get(l, l.upper()) for l in letters}

        answers = []
        if len(ans_lines) == 1 and re.search(r"(?i)true", ans_lines[0]):
            is_true = re.search(r"(?i)true", corr_line) is not None
            answers = [("true", is_true), ("false", not is_true)]
        else:
            for ln in ans_lines:
                m_a = re.match(r"^([A-Г])[\)\.]*\s*(.+)", ln)
                if not m_a:
                    errors.append((idx, f"Невірний формат відповіді: '{ln}'"))
                    break
                key, txt = m_a.group(1), m_a.group(2).strip()
                letter = mapping.get(key, key.upper())
                answers.append((txt, letter in correct_set))
            if len(answers) < 2:
                errors.append((idx, "Менше двох варіантів відповіді"))
                continue

        questions.append({"text": q_text, "answers": answers})

    return questions, errors

def parse_from_excel(uploaded_file) -> tuple[list, list]:
    """Парсер Excel: питання в колонці A, жовтий фон = правильна відповідь."""
    wb = load_workbook(uploaded_file, data_only=True)
    ws = wb.active
    items = []
    for cell in ws['A']:
        txt = str(cell.value).strip() if cell.value else ""
        is_corr = False
        if cell.fill and getattr(cell.fill, 'fill_type', None) == 'solid':
            if getattr(cell.fill.start_color, 'rgb', '').endswith('FFFF00'):
                is_corr = True
        items.append((txt, is_corr))

    blocks, curr = [], []
    for txt, corr in items:
        if not txt and curr:
            blocks.append(curr)
            curr = []
        elif txt:
            curr.append((txt, corr))
    if curr:
        blocks.append(curr)

    questions, errors = [], []
    for idx, blk in enumerate(blocks, 1):
        if len(blk) < 3:
            errors.append((idx, "Потрібно питання + ≥2 відповіді"))
            continue
        q_text = blk[0][0]
        answers = blk[1:]
        questions.append({"text": q_text, "answers": answers})

    return questions, errors

# ================== Інтерфейс режимів ==================

mode = st.sidebar.selectbox("Виберіть режим створення тесту", [
    "1. Excel", "2. По тексту (GPT)", "3. Вручну", "4. Готовий тест"
])

if mode == "1. Excel":
    st.header("1️⃣ Режим Excel")
    file = st.file_uploader("Завантажте .xlsx", type=["xlsx"])
    if file:
        qs, errs = parse_from_excel(file)
        if errs:
            st.error("Помилки при парсингу Excel:")
            for i, msg in errs:
                st.write(f"- Блок {i}: {msg}")
        else:
            st.success(f"Знайдено {len(qs)} питань")
        if qs and st.button("Генерувати XML"):
            xml_str = generate_moodle_xml_string(qs)
            st.subheader("📄 Попередній перегляд XML")
            st.text_area("", xml_str, height=300)
            download_xml(xml_str, "excel_test.xml")

elif mode == "2. По тексту (GPT)":
    st.header("2️⃣ Режим GPT-генерації")
    user_text = st.text_area("Вставте текст для генерації тесту українською", height=200)
    if st.button("Створити тест"):
        prog = st.progress(0)
        status = st.empty()
        status.text("1/3: Підготовка запиту…"); prog.progress(10)
        system_prompt = (
            """Ви — асистент із жорстким обмеженням на 10 питань у форматі Moodle XML українською мовою.
1) Використайте тільки наданий текст.
2) Створіть **саме 10** логічних питань українською:
   – 4–5 питань Single-choice,
   – 2–3 питання True/False,
   – 2–3 питання Multiple-choice.
3) Кожне питання (окрім True/False) має мати 4 варіанти (A, B, C, D).
Для Multiple-choice відзначте кілька варіантів як правильні.
4) Перед поверненням перевірте, що загальна кількість питань = 10.
5) Поверніть **тільки** XML-код (без будь-яких коментарів) і одразу припиніть після 10-го питання.
<END>
"""
        )
        status.text("2/3: Відправка запиту…"); prog.progress(40)
        resp = openai.ChatCompletion.create(
            model="gpt-4",
            messages=[{"role": "system", "content": system_prompt},
                      {"role": "user", "content": user_text}],
            temperature=0,
        )
        status.text("3/3: Отримання відповіді…"); prog.progress(70)
        text = resp.choices[0].message.content.strip()
        if text.startswith("<?xml") or text.lstrip().startswith("<quiz"):
            count = len(re.findall(r'<question', text))
            if count != 10:
                st.error(f"GPT повернув {count} питань замість 10.")
            xml_str = text if text.startswith("<?xml") else '<?xml version="1.0" encoding="UTF-8"?>\n' + text
            if count > 10:
                parts = re.findall(r'(<question[\s\S]*?</question>)', xml_str)
                xml_str = '<?xml version="1.0" encoding="UTF-8"?>\n<quiz>' + ''.join(parts[:10]) + '</quiz>'
                st.warning("Обрізано до перших 10 питань.")
            elif count < 10:
                st.error(f"GPT згенерував тільки {count} питань, а потрібно 10.")
            with open("gpt_test.xml", "w", encoding="utf-8") as f:
                f.write(xml_str)
            st.subheader("📄 Попередній перегляд XML")
            st.code(xml_str)
            download_xml(xml_str, "gpt_test.xml")
            status.text("Готово!")
        else:
            status.text("3/3: Парсинг відповіді…"); prog.progress(80)
            qs, errs = parse_text_format(text)
            if errs:
                st.error("Помилки при парсингу GPT-відповіді:")
                for i, msg in errs:
                    st.write(f"- Блок {i}: {msg}")
            else:
                if len(qs) != 10:
                    st.error(f"GPT згенерував {len(qs)} питань замість 10.")
                    qs = qs[:10]
                status.text("4/4: Генерація XML…"); prog.progress(100)
                xml_str = generate_moodle_xml_string(qs)
                with open("gpt_test.xml", "w", encoding="utf-8") as f:
                    f.write(xml_str)
                st.subheader("📄 Попередній перегляд XML")
                st.code(xml_str)
                download_xml(xml_str, "gpt_test.xml")
                status.text("Готово!")

elif mode == "3. Вручну":
    st.header("3️⃣ Ручне створення питань")
    if "manual_qs" not in st.session_state:
        st.session_state.manual_qs = []

    qtype = st.selectbox("Тип питання", ["Single-choice", "Multiple-choice", "True/False"])

    with st.form("manual_form"):
        q_txt = st.text_input("Текст питання")
        answers = []
        if qtype == "True/False":
            corr = st.radio("Правильна відповідь", ["true", "false"])
            answers = [("true", corr == "true"), ("false", corr == "false")]
        else:
            cols = st.columns([8, 1])
            with cols[0]:
                a1 = st.text_input("A."); a2 = st.text_input("B.")
                a3 = st.text_input("C."); a4 = st.text_input("D.")
            with cols[1]:
                c1 = st.checkbox("", key="c1")
                c2 = st.checkbox("", key="c2")
                c3 = st.checkbox("", key="c3")
                c4 = st.checkbox("", key="c4")
            answers = [(a1, c1), (a2, c2), (a3, c3), (a4, c4)]

        submitted = st.form_submit_button("Додати питання")

    if submitted:
        st.session_state.manual_qs.append({"text": q_txt, "answers": answers})

    if st.session_state.manual_qs:
        st.subheader("Список питань")
        for idx, q in enumerate(st.session_state.manual_qs, 1):
            st.write(f"{idx}. {q['text']}")
        if st.button("Генерувати XML для ручних питань"):
            xml_str = generate_moodle_xml_string(st.session_state.manual_qs)
            st.text_area("", xml_str, height=300)
            download_xml(xml_str, "manual_test.xml")

elif mode == "4. Готовий тест":
    st.header("4️⃣ Режим готового тесту")

    ready = st.text_area(
        "Вставте готовий тест у зазначеному форматі (1. … A. … Правильний відповідь: …)",
        height=200
    )

    if st.button("Генерувати XML з готового тесту", key="gen_ready"):
        qs, errs = parse_text_format(ready)

        if errs:
            st.error("Помилки при розборі тесту:")
            for idx, msg in errs:
                st.write(f"- Блок {idx}: {msg}")
        else:
            st.success(f"Знайдено {len(qs)} питань — генеруємо XML…")
            xml_str = generate_moodle_xml_string(qs)

            st.text_area(
                "XML-код",
                xml_str,
                height=300,
                label_visibility="collapsed"
            )
            download_xml(xml_str, "ready_test.xml")
