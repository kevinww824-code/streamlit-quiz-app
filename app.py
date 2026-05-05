import re
import random
import pandas as pd
import streamlit as st

EXCEL_PATH = "CDMP-Associate_1.xlsx"
SHEET_NAME = "sheet1"

# ---------- Helpers ----------
def extract_answer_letter(v):
    # Example in Excel: "正确答案：D" -> returns "D"
    m = re.search(r"([A-E])", str(v).upper())
    return m.group(1) if m else ""

def clean_option_text(text):
    """
    Removes leading 'A. ', 'B. ', etc. if your Excel already contains it.
    'A. Attributes' -> 'Attributes'
    """
    return re.sub(r"^[A-E]\s*\.\s*", "", str(text)).strip()

@st.cache_data(show_spinner=False)
def load_quiz():
    """
    Your Excel has 2 header rows and data begins from row index 2. [1](https://newworlddevelopment-my.sharepoint.com/personal/kevinwong_nwd_com_hk/_layouts/15/Doc.aspx?sourcedoc=%7B43BF5816-F67A-4355-A61C-A0D2516DAFD0%7D&file=CDMP-Associate_1%20-%20Copy.xlsx&action=default&mobileredirect=true)
    Columns used (based on your file structure):
    r[1]  = Question EN
    r[3]  = Option A EN
    r[5]  = Option B EN
    r[7]  = Option C EN
    r[9]  = Option D EN
    r[11] = Option E EN
    r[13] = Answer (e.g., 正确答案：D)
    r[14] = Interpretation/Explanation (your sample shows Chinese text here) [1](https://newworlddevelopment-my.sharepoint.com/personal/kevinwong_nwd_com_hk/_layouts/15/Doc.aspx?sourcedoc=%7B43BF5816-F67A-4355-A61C-A0D2516DAFD0%7D&file=CDMP-Associate_1%20-%20Copy.xlsx&action=default&mobileredirect=true)
    """
    raw = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_NAME, header=None, engine="openpyxl")
    data = raw.iloc[2:].copy()  # data starts from row index 2 [1](https://newworlddevelopment-my.sharepoint.com/personal/kevinwong_nwd_com_hk/_layouts/15/Doc.aspx?sourcedoc=%7B43BF5816-F67A-4355-A61C-A0D2516DAFD0%7D&file=CDMP-Associate_1%20-%20Copy.xlsx&action=default&mobileredirect=true)

    quiz = []
    for _, r in data.iterrows():
        q = str(r[1]).strip() if pd.notna(r[1]) else ""
        if not q:
            continue

        options = {
            "A": clean_option_text(r[3]) if pd.notna(r[3]) else "",
            "B": clean_option_text(r[5]) if pd.notna(r[5]) else "",
            "C": clean_option_text(r[7]) if pd.notna(r[7]) else "",
            "D": clean_option_text(r[9]) if pd.notna(r[9]) else "",
            "E": clean_option_text(r[11]) if pd.notna(r[11]) else "",
        }
        options = {k: v for k, v in options.items() if v}

        quiz.append({
            "question": q,
            "options": options,
            "answer": extract_answer_letter(r[13]),
            "interpretation": str(r[14]).strip() if pd.notna(r[14]) else ""
        })

    return quiz

def reset_session_for_mode(mode: str, n_questions: int):
    """Reset progress, order, score when switching modes."""
    st.session_state.mode = mode
    st.session_state.idx = 0
    st.session_state.score = 0
    st.session_state.submitted = False
    st.session_state.feedback = None
    st.session_state.last_choice = None
    st.session_state.responses = []  # store (q_index, chosen, correct)

    order = list(range(n_questions))
    if mode == "Exam Mode":
        random.shuffle(order)  # random order for exam
    st.session_state.order = order

# ---------- UI ----------
st.set_page_config(page_title="Quiz App", layout="centered")
st.title("Quiz App (English)")

quiz = load_quiz()
if not quiz:
    st.error("No questions loaded. Check Excel format / file path.")
    st.stop()

n = len(quiz)

# ---------- Session state init ----------
if "mode" not in st.session_state:
    st.session_state.mode = "Learning Mode"
if "idx" not in st.session_state:
    st.session_state.idx = 0
if "score" not in st.session_state:
    st.session_state.score = 0
if "submitted" not in st.session_state:
    st.session_state.submitted = False
if "feedback" not in st.session_state:
    st.session_state.feedback = None
if "last_choice" not in st.session_state:
    st.session_state.last_choice = None
if "responses" not in st.session_state:
    st.session_state.responses = []
if "order" not in st.session_state:
    st.session_state.order = list(range(n))  # default learning mode order

# ---------- Mode switch control ----------
st.sidebar.header("Mode")
selected_mode = st.sidebar.radio(
    "Choose a mode:",
    ["Learning Mode", "Exam Mode"],
    index=0 if st.session_state.mode == "Learning Mode" else 1
)

# If user switched mode, reset session for that mode
if selected_mode != st.session_state.mode:
    reset_session_for_mode(selected_mode, n)
    st.rerun()

# Optional controls
with st.sidebar.expander("Exam Options", expanded=False):
    hide_explain_in_exam = st.checkbox(
        "Hide explanations during Exam Mode (show at end)",
        value=True,
        disabled=(st.session_state.mode != "Exam Mode")
    )

# ---------- Completion screen ----------
if st.session_state.idx >= n:
    st.success(f"✅ Completed {st.session_state.mode}! Score: {st.session_state.score}/{n} ({st.session_state.score/n*100:.1f}%)")

    # Review section (especially useful for Exam Mode)
    st.subheader("Review")
    if st.session_state.responses:
        for (q_idx, chosen, correct) in st.session_state.responses:
            item = quiz[q_idx]
            st.markdown(f"**Q:** {item['question']}")
            st.write(f"Your answer: **{chosen}** | Correct: **{correct}**")
            if item["interpretation"]:
                st.info("Explanation / Interpretation:")
                st.write(item["interpretation"])
            st.divider()

    if st.button("Restart"):
        reset_session_for_mode(st.session_state.mode, n)
        st.rerun()

    st.stop()

# ---------- Progress bar (both modes) ----------
current_num = st.session_state.idx + 1
progress = st.session_state.idx / n  # 0 to <1
st.progress(progress)
st.caption(f"Progress: {current_num} / {n} ({progress*100:.1f}%)")

# ---------- Get the current question using order ----------
q_index = st.session_state.order[st.session_state.idx]
item = quiz[q_index]

st.write(f"### Q{current_num}: {item['question']}")

choice = st.radio(
    "Select an answer:",
    list(item["options"].keys()),
    format_func=lambda k: f"{k}. {item['options'][k]}",
    disabled=st.session_state.submitted
)

col1, col2, col3 = st.columns([1, 1, 1])

with col1:
    if st.button("Submit", disabled=st.session_state.submitted):
        st.session_state.submitted = True
        st.session_state.last_choice = choice

        correct = item["answer"]
        if choice == correct:
            st.session_state.score += 1
            st.session_state.feedback = ("success", "Correct ✅")
        else:
            st.session_state.feedback = ("error", f"Incorrect ❌  Correct answer: {correct}")

        # store response for end review
        st.session_state.responses.append((q_index, choice, correct))
        st.rerun()

with col2:
    if st.button("Next", disabled=not st.session_state.submitted):
        st.session_state.idx += 1
        st.session_state.submitted = False
        st.session_state.feedback = None
        st.session_state.last_choice = None
        st.rerun()

with col3:
    if st.button("Restart Quiz"):
        reset_session_for_mode(st.session_state.mode, n)
        st.rerun()

# ---------- Feedback + explanation ----------
if st.session_state.submitted and st.session_state.feedback:
    level, msg = st.session_state.feedback
    getattr(st, level)(msg)

    # Learning Mode: always show explanation immediately
    # Exam Mode: show explanation depending on hide_explain_in_exam option
    if st.session_state.mode == "Learning Mode":
        if item["interpretation"]:
            st.info("Explanation / Interpretation:")
            st.write(item["interpretation"])
    else:
        # Exam Mode
        if not hide_explain_in_exam:
            if item["interpretation"]:
                st.info("Explanation / Interpretation:")
                st.write(item["interpretation"])
        else:
            st.caption("Exam Mode: Explanation will be shown in the review section after completion.")
