import re
import pandas as pd
import streamlit as st

EXCEL_PATH = "CDMP-Associate_1.xlsx"
SHEET_NAME = "sheet1"

def extract_answer_letter(v):
    # Example in Excel: "正确答案：D" -> returns "D"
    m = re.search(r"([A-E])", str(v).upper())
    return m.group(1) if m else ""

def clean_option_text(text):
    """
    Removes leading 'A. ', 'B. ' ... if your Excel already contains it.
    So 'A. Attributes' -> 'Attributes'
    """
    return re.sub(r"^[A-E]\s*\.\s*", "", str(text)).strip()

def load_quiz():
    # Your Excel has 2 header rows and real data starts from row index 2. [1](https://newworlddevelopment-my.sharepoint.com/personal/kevinwong_nwd_com_hk/_layouts/15/Doc.aspx?sourcedoc=%7B43BF5816-F67A-4355-A61C-A0D2516DAFD0%7D&file=CDMP-Associate_1%20-%20Copy.xlsx&action=default&mobileredirect=true)
    raw = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_NAME, header=None, engine="openpyxl")
    data = raw.iloc[2:].copy()

    quiz = []
    for _, r in data.iterrows():
        q = str(r[1]).strip() if pd.notna(r[1]) else ""
        if not q:
            continue

        # Apply clean_option_text here to remove duplicated "A." etc.
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

# ---------------- Streamlit UI ----------------
st.title("Quiz App (English)")
quiz = load_quiz()

if not quiz:
    st.error("No questions loaded. Check Excel format.")
    st.stop()

# Initialize session state
if "idx" not in st.session_state:
    st.session_state.idx = 0
if "score" not in st.session_state:
    st.session_state.score = 0
if "submitted" not in st.session_state:
    st.session_state.submitted = False
if "feedback" not in st.session_state:
    st.session_state.feedback = None  # ("success"/"error", message)

# Completed
if st.session_state.idx >= len(quiz):
    st.success(f"Quiz completed! Score: {st.session_state.score}/{len(quiz)}")
    st.stop()

item = quiz[st.session_state.idx]
st.write(f"### Q{st.session_state.idx + 1}: {item['question']}")

# Disable choice after submit to avoid changing after grading
choice = st.radio(
    "Select an answer:",
    list(item["options"].keys()),
    format_func=lambda k: f"{k}. {item['options'][k]}",
    disabled=st.session_state.submitted
)

col1, col2 = st.columns(2)

with col1:
    if st.button("Submit", disabled=st.session_state.submitted):
        if choice == item["answer"]:
            st.session_state.score += 1
            st.session_state.feedback = ("success", "Correct ✅")
        else:
            st.session_state.feedback = ("error", f"Incorrect ❌  Correct answer: {item['answer']}")
        st.session_state.submitted = True
        st.rerun()

with col2:
    if st.button("Next Question", disabled=not st.session_state.submitted):
        st.session_state.idx += 1
        st.session_state.submitted = False
        st.session_state.feedback = None
        st.rerun()

# Show feedback + interpretation after submit
if st.session_state.submitted and st.session_state.feedback:
    level, msg = st.session_state.feedback
    getattr(st, level)(msg)

    if item["interpretation"]:
        st.info("Explanation / Interpretation:")
        st.write(item["interpretation"])
