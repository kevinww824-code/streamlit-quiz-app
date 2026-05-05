import re
import time
import random
import pandas as pd
import streamlit as st

# -----------------------------
# CONFIG
# -----------------------------
EXCEL_PATH = "CDMP-Associate_1.xlsx"
SHEET_NAME = "sheet1"

st.set_page_config(page_title="Quiz App", layout="centered")


# -----------------------------
# HELPERS (Excel parsing)
# -----------------------------
def extract_answer_letter(v):
    """
    Excel answer cell example: '正确答案：D' -> returns 'D'
    """
    m = re.search(r"([A-E])", str(v).upper())
    return m.group(1) if m else ""


def clean_option_text(text):
    """
    Removes leading 'A. ', 'B. ', etc. if Excel already contains it.
    e.g., 'A. Attributes' -> 'Attributes'
    """
    return re.sub(r"^[A-E]\s*\.\s*", "", str(text)).strip()


@st.cache_data(show_spinner=False)
def load_quiz():
    """
    Matches your Excel layout:
    - 2 header rows; real data starts at row index 2
    - Question EN: column 1
    - Options EN: A=3, B=5, C=7, D=9, E=11
    - Answer: column 13 (e.g., 正确答案：D)
    - Interpretation/Explanation: column 14
    """
    raw = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_NAME, header=None, engine="openpyxl")
    data = raw.iloc[2:].copy()  # data starts from row index 2 [1](https://newworlddevelopment-my.sharepoint.com/personal/kevinwong_nwd_com_hk/_layouts/15/Doc.aspx?sourcedoc=%7B43BF5816-F67A-4355-A61C-A0D2516DAFD0%7D&file=CDMP-Associate_1%20-%20Copy.xlsx&action=default&mobileredirect=true)

    quiz = []
    for _, r in data.iterrows():
        q = str(r[1]).strip() if pd.notna(r[1]) else ""  # English question in col 1 [1](https://newworlddevelopment-my.sharepoint.com/personal/kevinwong_nwd_com_hk/_layouts/15/Doc.aspx?sourcedoc=%7B43BF5816-F67A-4355-A61C-A0D2516DAFD0%7D&file=CDMP-Associate_1%20-%20Copy.xlsx&action=default&mobileredirect=true)
        if not q:
            continue

        options = {
            "A": clean_option_text(r[3]) if pd.notna(r[3]) else "",   # Option A EN in col 3 [1](https://newworlddevelopment-my.sharepoint.com/personal/kevinwong_nwd_com_hk/_layouts/15/Doc.aspx?sourcedoc=%7B43BF5816-F67A-4355-A61C-A0D2516DAFD0%7D&file=CDMP-Associate_1%20-%20Copy.xlsx&action=default&mobileredirect=true)
            "B": clean_option_text(r[5]) if pd.notna(r[5]) else "",   # Option B EN in col 5 [1](https://newworlddevelopment-my.sharepoint.com/personal/kevinwong_nwd_com_hk/_layouts/15/Doc.aspx?sourcedoc=%7B43BF5816-F67A-4355-A61C-A0D2516DAFD0%7D&file=CDMP-Associate_1%20-%20Copy.xlsx&action=default&mobileredirect=true)
            "C": clean_option_text(r[7]) if pd.notna(r[7]) else "",   # Option C EN in col 7 [1](https://newworlddevelopment-my.sharepoint.com/personal/kevinwong_nwd_com_hk/_layouts/15/Doc.aspx?sourcedoc=%7B43BF5816-F67A-4355-A61C-A0D2516DAFD0%7D&file=CDMP-Associate_1%20-%20Copy.xlsx&action=default&mobileredirect=true)
            "D": clean_option_text(r[9]) if pd.notna(r[9]) else "",   # Option D EN in col 9 [1](https://newworlddevelopment-my.sharepoint.com/personal/kevinwong_nwd_com_hk/_layouts/15/Doc.aspx?sourcedoc=%7B43BF5816-F67A-4355-A61C-A0D2516DAFD0%7D&file=CDMP-Associate_1%20-%20Copy.xlsx&action=default&mobileredirect=true)
            "E": clean_option_text(r[11]) if pd.notna(r[11]) else "", # Option E EN in col 11 [1](https://newworlddevelopment-my.sharepoint.com/personal/kevinwong_nwd_com_hk/_layouts/15/Doc.aspx?sourcedoc=%7B43BF5816-F67A-4355-A61C-A0D2516DAFD0%7D&file=CDMP-Associate_1%20-%20Copy.xlsx&action=default&mobileredirect=true)
        }
        options = {k: v for k, v in options.items() if v}

        answer = extract_answer_letter(r[13])  # Answer in col 13 [1](https://newworlddevelopment-my.sharepoint.com/personal/kevinwong_nwd_com_hk/_layouts/15/Doc.aspx?sourcedoc=%7B43BF5816-F67A-4355-A61C-A0D2516DAFD0%7D&file=CDMP-Associate_1%20-%20Copy.xlsx&action=default&mobileredirect=true)
        interpretation = str(r[14]).strip() if pd.notna(r[14]) else ""  # Explanation in col 14 [1](https://newworlddevelopment-my.sharepoint.com/personal/kevinwong_nwd_com_hk/_layouts/15/Doc.aspx?sourcedoc=%7B43BF5816-F67A-4355-A61C-A0D2516DAFD0%7D&file=CDMP-Associate_1%20-%20Copy.xlsx&action=default&mobileredirect=true)

        quiz.append({
            "question": q,
            "options": options,
            "answer": answer,
            "interpretation": interpretation
        })

    return quiz


# -----------------------------
# SESSION STATE MANAGEMENT
# -----------------------------
def reset_quiz(mode: str, n_questions: int, shuffle_questions: bool):
    st.session_state.mode = mode
    st.session_state.idx = 0
    st.session_state.score = 0
    st.session_state.submitted = False
    st.session_state.feedback = None
    st.session_state.last_choice = None
    st.session_state.responses = []  # list of dicts
    st.session_state.started_at = None  # for timer

    order = list(range(n_questions))
    if shuffle_questions:
        random.shuffle(order)
    st.session_state.order = order


def ensure_state(quiz_len: int):
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
        st.session_state.order = list(range(quiz_len))
    if "started_at" not in st.session_state:
        st.session_state.started_at = None


# -----------------------------
# APP UI
# -----------------------------
st.title("Quiz App (English)")

quiz = load_quiz()
if not quiz:
    st.error("No questions loaded. Please check the Excel file path/name and format.")
    st.stop()

n = len(quiz)
ensure_state(n)

# -----------------------------
# SIDEBAR CONTROLS
# -----------------------------
st.sidebar.header("Settings")

selected_mode = st.sidebar.radio(
    "Mode",
    ["Learning Mode", "Exam Mode"],
    index=0 if st.session_state.mode == "Learning Mode" else 1
)

# Mode behavior
is_exam = (selected_mode == "Exam Mode")

shuffle_questions = st.sidebar.checkbox(
    "Shuffle question order (Exam Mode)",
    value=True,
    disabled=not is_exam
)

shuffle_options = st.sidebar.checkbox(
    "Shuffle option order (A–E) (Exam Mode)",
    value=True,
    disabled=not is_exam
)

hide_feedback_until_end = st.sidebar.checkbox(
    "Hide correctness until end (Exam Mode)",
    value=True,
    disabled=not is_exam
)

hide_explanations_until_end = st.sidebar.checkbox(
    "Hide explanations until end (Exam Mode)",
    value=True,
    disabled=not is_exam
)

enable_timer = st.sidebar.checkbox(
    "Enable timer (Exam Mode)",
    value=False,
    disabled=not is_exam
)

timer_minutes = st.sidebar.number_input(
    "Timer minutes",
    min_value=1,
    max_value=300,
    value=60,
    step=5,
    disabled=not (is_exam and enable_timer)
)

# Reset / Apply mode switch
if selected_mode != st.session_state.mode:
    # Learning Mode: keep original order
    if selected_mode == "Learning Mode":
        reset_quiz("Learning Mode", n, shuffle_questions=False)
    else:
        reset_quiz("Exam Mode", n, shuffle_questions=shuffle_questions)
    st.rerun()

# Manual restart
if st.sidebar.button("Restart / New Attempt"):
    if st.session_state.mode == "Learning Mode":
        reset_quiz("Learning Mode", n, shuffle_questions=False)
    else:
        reset_quiz("Exam Mode", n, shuffle_questions=shuffle_questions)
    st.rerun()


# -----------------------------
# TIMER HANDLING (Exam Mode)
# -----------------------------
def get_time_left_seconds():
    if not enable_timer or not is_exam:
        return None
    if st.session_state.started_at is None:
        st.session_state.started_at = time.time()
    elapsed = time.time() - st.session_state.started_at
    total = timer_minutes * 60
    return max(0, int(total - elapsed))


time_left = get_time_left_seconds()
if is_exam and enable_timer:
    mins = time_left // 60
    secs = time_left % 60
    st.sidebar.metric("Time left", f"{mins:02d}:{secs:02d}")

    # Auto-finish when time is up
    if time_left <= 0:
        st.warning("⏰ Time is up! Submitting your exam...")
        st.session_state.idx = n  # jump to completion page
        st.rerun()


# -----------------------------
# COMPLETION PAGE
# -----------------------------
if st.session_state.idx >= n:
    st.success(f"✅ Completed {st.session_state.mode}!")
    st.write(f"**Score:** {st.session_state.score}/{n} ({st.session_state.score/n*100:.1f}%)")

    # Review controls
    st.subheader("Review")
    wrong_only = st.checkbox("Show wrong answers only", value=is_exam)

    # Build review dataframe + display
    review_rows = []
    for resp in st.session_state.responses:
        correct = (resp["chosen"] == resp["correct"])
        if wrong_only and correct:
            continue

        q = quiz[resp["q_index"]]
        st.markdown(f"**Q:** {q['question']}")
        st.write(f"Your answer: **{resp['chosen']}** | Correct: **{resp['correct']}**")

        if q["interpretation"]:
            st.info("Explanation / Interpretation:")
            st.write(q["interpretation"])

        st.divider()

        review_rows.append({
            "question": q["question"],
            "chosen": resp["chosen"],
            "correct": resp["correct"],
            "is_correct": correct
        })

    # Download results
    if review_rows:
        df_out = pd.DataFrame(review_rows)
        csv = df_out.to_csv(index=False).encode("utf-8")
        st.download_button(
            "Download results (CSV)",
            data=csv,
            file_name="quiz_results.csv",
            mime="text/csv"
        )

    st.stop()


# -----------------------------
# PROGRESS BAR (Both Modes)
# -----------------------------
current_num = st.session_state.idx + 1
progress = st.session_state.idx / n
st.progress(progress)
st.caption(f"Progress: {current_num} / {n} ({progress*100:.1f}%)")


# -----------------------------
# CURRENT QUESTION
# -----------------------------
q_index = st.session_state.order[st.session_state.idx]
item = quiz[q_index]

st.write(f"### Q{current_num}: {item['question']}")

# Option order (shuffle only in Exam Mode if enabled)
option_keys = list(item["options"].keys())
if is_exam and shuffle_options:
    random.shuffle(option_keys)

choice = st.radio(
    "Select an answer:",
    option_keys,
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

        # Save response for end review
        st.session_state.responses.append({
            "q_index": q_index,
            "chosen": choice,
            "correct": correct
        })

        st.rerun()

with col2:
    if st.button("Next", disabled=not st.session_state.submitted):
        st.session_state.idx += 1
        st.session_state.submitted = False
        st.session_state.feedback = None
        st.session_state.last_choice = None
        st.rerun()

with col3:
    if st.button("Quit / Finish Now"):
        st.session_state.idx = n
        st.rerun()


# -----------------------------
# FEEDBACK + EXPLANATION
# -----------------------------
if st.session_state.submitted and st.session_state.feedback:
    show_feedback = True
    show_explain = True

    # Exam Mode rules
    if is_exam and hide_feedback_until_end:
        show_feedback = False
    if is_exam and hide_explanations_until_end:
        show_explain = False

    # Learning Mode: always show feedback + explanation
    if not is_exam:
        show_feedback = True
        show_explain = True

    if show_feedback:
        level, msg = st.session_state.feedback
        getattr(st, level)(msg)
    else:
        st.caption("Exam Mode: correctness will be shown in the final review.")

    if show_explain:
        if item["interpretation"]:
            st.info("Explanation / Interpretation:")
            st.write(item["interpretation"])
    else:
        st.caption("Exam Mode: explanations will be shown in the final review.")