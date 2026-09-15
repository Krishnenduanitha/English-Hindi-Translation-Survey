import streamlit as st
import pandas as pd
import os
import random
from datetime import datetime


# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="English-Hindi Speech Study",
    page_icon="🎧",
    layout="centered"
)


# ============================================================
# FILE PATHS
# ============================================================

BASE_DIR = os.path.dirname(
    os.path.abspath(__file__)
)

EXCEL_FILE = os.path.join(
    BASE_DIR,
    "top10_gold14_nllb_FAILURES.xlsx"
)

AUDIO_FOLDER = os.path.join(
    BASE_DIR,
    "audio"
)

RESPONSES_FILE = os.path.join(
    BASE_DIR,
    "responses.csv"
)


# ============================================================
# LOAD EXCEL
# ============================================================

@st.cache_data
def load_questions():

    if not os.path.exists(EXCEL_FILE):
        st.error(
            "Excel file not found: "
            + EXCEL_FILE
        )
        st.stop()

    df = pd.read_excel(EXCEL_FILE)

    df = df.dropna(
        how="all"
    ).reset_index(drop=True)

    required_columns = [
        "sample id",
        "english sentence",
        "hindi translation",
        "machine translation",
        "prosody translation",
        "emphasized word",
        "audiofile"
    ]

    missing = [
        col
        for col in required_columns
        if col not in df.columns
    ]

    if missing:
        st.error(
            "Missing columns in Excel:\n\n"
            + "\n".join(missing)
        )
        st.stop()

    return df


df = load_questions()


# ============================================================
# GENERATE PARTICIPANT ID
# ============================================================

def generate_participant_id():

    if not os.path.exists(RESPONSES_FILE):
        return "P001"

    try:

        old_data = pd.read_csv(
            RESPONSES_FILE
        )

        if (
            "participant_id"
            not in old_data.columns
        ):
            return "P001"

        ids = (
            old_data["participant_id"]
            .dropna()
            .astype(str)
            .unique()
        )

        numbers = []

        for pid in ids:

            if pid.startswith("P"):

                try:
                    numbers.append(
                        int(pid[1:])
                    )
                except ValueError:
                    pass

        if not numbers:
            return "P001"

        return f"P{max(numbers) + 1:03d}"

    except Exception:

        return "P001"


# ============================================================
# SESSION STATE
# ============================================================

defaults = {

    "page": "participant",

    "participant_id": None,

    "participant_name": "",

    "place": "",

    "age": None,

    "english_proficiency": "",

    "hindi_proficiency": "",

    "participant_start_time": None,

    "current_question": 0,

    "answers": {},

    "randomized_options": {},

    "question_start_times": {}

}

for key, value in defaults.items():

    if key not in st.session_state:

        st.session_state[key] = value


# ============================================================
# SAVE RESPONSES
# ============================================================

def save_responses():

    rows = []

    submission_time = datetime.now().strftime(
        "%Y-%m-%d %H:%M:%S"
    )

    for _, row in df.iterrows():

        sample_id = row["sample id"]

        if sample_id not in st.session_state.answers:
            continue

        answer = st.session_state.answers[
            sample_id
        ]

        rows.append({

            "participant_id":
                st.session_state.participant_id,

            "participant_name":
                st.session_state.participant_name,

            "place":
                st.session_state.place,

            "age":
                st.session_state.age,

            "english_proficiency":
                st.session_state.english_proficiency,

            "hindi_proficiency":
                st.session_state.hindi_proficiency,

            "participant_start_time":
                st.session_state.participant_start_time,

            "submission_time":
                submission_time,

            "sample_id":
                sample_id,

            "english_sentence":
                row["english sentence"],

            "emphasized_word":
                row["emphasized word"],

            "audiofile":
                row["audiofile"],

            "selected_translation":
                answer["translation"],

            "selected_option_number":
                answer["option_number"],

            "selected_translation_type":
                answer["translation_type"],

            "response_time_seconds":
                answer["response_time"]
        })


    if not rows:
        return


    new_data = pd.DataFrame(rows)


    if os.path.exists(RESPONSES_FILE):

        try:

            old_data = pd.read_csv(
                RESPONSES_FILE
            )

            final_data = pd.concat(
                [
                    old_data,
                    new_data
                ],
                ignore_index=True
            )

        except Exception:

            final_data = new_data

    else:

        final_data = new_data


    final_data.to_csv(
        RESPONSES_FILE,
        index=False
    )


# ============================================================
# PAGE 1 — PARTICIPANT DETAILS
# ============================================================

if st.session_state.page == "participant":

    st.title(
        "English–Hindi Speech Study"
    )

    st.write(
        """
        Welcome to the study.

        This experiment examines how emphasis
        and prosody in English speech influence
        English-to-Hindi translation.

        Please enter your details below.
        """
    )

    st.divider()


    name = st.text_input(
        "Name"
    )


    place = st.text_input(
        "Place"
    )


    age = st.number_input(
        "Age",
        min_value=10,
        max_value=100,
        value=20,
        step=1
    )


    english_proficiency = st.selectbox(
        "English Proficiency",
        [
            "Beginner",
            "Intermediate",
            "Advanced",
            "Native"
        ]
    )


    hindi_proficiency = st.selectbox(
        "Hindi Proficiency",
        [
            "Beginner",
            "Intermediate",
            "Advanced",
            "Native"
        ]
    )


    st.write("")


    if st.button(
        "Start Experiment",
        type="primary",
        use_container_width=True
    ):

        if not name.strip():

            st.warning(
                "Please enter your name."
            )

        elif not place.strip():

            st.warning(
                "Please enter your place."
            )

        else:

            st.session_state.participant_id = (
                generate_participant_id()
            )

            st.session_state.participant_name = (
                name.strip()
            )

            st.session_state.place = (
                place.strip()
            )

            st.session_state.age = age

            st.session_state.english_proficiency = (
                english_proficiency
            )

            st.session_state.hindi_proficiency = (
                hindi_proficiency
            )

            st.session_state.participant_start_time = (
                datetime.now().strftime(
                    "%Y-%m-%d %H:%M:%S"
                )
            )

            st.session_state.current_question = 0

            st.session_state.answers = {}

            st.session_state.randomized_options = {}

            st.session_state.question_start_times = {}

            st.session_state.page = "experiment"

            st.rerun()


# ============================================================
# PAGE 2 — EXPERIMENT
# ============================================================

elif st.session_state.page == "experiment":

    question_index = (
        st.session_state.current_question
    )

    total_questions = len(df)

    row = df.iloc[
        question_index
    ]

    sample_id = row["sample id"]


    # --------------------------------------------------------
    # HEADER
    # --------------------------------------------------------

    st.title(
        "Translation Experiment"
    )

    st.caption(
        f"Participant ID: "
        f"{st.session_state.participant_id}"
    )

    st.progress(
        (question_index + 1)
        / total_questions
    )

    st.write(
        f"Question {question_index + 1} "
        f"of {total_questions}"
    )

    st.divider()


    # --------------------------------------------------------
    # ENGLISH SENTENCE
    # --------------------------------------------------------

    st.subheader(
        "English Sentence"
    )

    st.write(
        str(row["english sentence"])
    )


    # --------------------------------------------------------
    # EMPHASIZED WORD
    # --------------------------------------------------------

    st.info(
        f"Emphasized word: "
        f"**{row['emphasized word']}**"
    )


    # --------------------------------------------------------
    # AUDIO
    # --------------------------------------------------------

    st.subheader(
        "Listen to the sentence"
    )

    audio_filename = str(
        row["audiofile"]
    )

    audio_path = os.path.join(
        AUDIO_FOLDER,
        audio_filename
    )

    if os.path.exists(audio_path):

        st.audio(
            audio_path,
            format="audio/mp4"
        )

    else:

        st.error(
            f"Audio file not found: "
            f"{audio_filename}"
        )


    st.divider()


    # --------------------------------------------------------
    # TRANSLATION OPTIONS
    # --------------------------------------------------------

    options = [

        (
            "Hindi Translation",
            row["hindi translation"]
        ),

        (
            "Machine Translation",
            row["machine translation"]
        ),

        (
            "Prosody Translation",
            row["prosody translation"]
        )

    ]


    # Remove empty options
    valid_options = []

    for source, translation in options:

        if pd.notna(translation):

            translation = str(
                translation
            ).strip()

            if translation:

                valid_options.append(
                    (
                        source,
                        translation
                    )
                )


    # --------------------------------------------------------
    # RANDOMIZE ORDER
    # --------------------------------------------------------

    if sample_id not in st.session_state.randomized_options:

        shuffled_options = (
            valid_options.copy()
        )

        random.shuffle(
            shuffled_options
        )

        st.session_state.randomized_options[
            sample_id
        ] = shuffled_options

    else:

        shuffled_options = (
            st.session_state.randomized_options[
                sample_id
            ]
        )


    # --------------------------------------------------------
    # QUESTION
    # --------------------------------------------------------

    st.subheader(
        "Which Hindi translation best matches the sentence you heard?"
    )

    display_options = [

        translation

        for source, translation
        in shuffled_options

    ]


    # --------------------------------------------------------
    # TIMER
    # --------------------------------------------------------

    if sample_id not in st.session_state.question_start_times:

        st.session_state.question_start_times[
            sample_id
        ] = datetime.now()


    # --------------------------------------------------------
    # RADIO BUTTONS
    # --------------------------------------------------------

    selected_translation = st.radio(

        "Hindi options",

        display_options,

        key=f"radio_{sample_id}",

        label_visibility="collapsed"
    )


    # --------------------------------------------------------
    # STORE ANSWER
    # --------------------------------------------------------

    if selected_translation:

        selected_index = (
            display_options.index(
                selected_translation
            )
        )

        selected_source = (
            shuffled_options[
                selected_index
            ][0]
        )

        start_time = (
            st.session_state.question_start_times[
                sample_id
            ]
        )

        response_time = (
            datetime.now()
            - start_time
        ).total_seconds()


        st.session_state.answers[
            sample_id
        ] = {

            "translation":
                selected_translation,

            "option_number":
                selected_index + 1,

            "translation_type":
                selected_source,

            "response_time":
                round(
                    response_time,
                    2
                )
        }


    st.divider()


    # --------------------------------------------------------
    # NAVIGATION
    # --------------------------------------------------------

    col1, col2 = st.columns(2)


    with col1:

        if question_index > 0:

            if st.button(
                "← Previous",
                use_container_width=True
            ):

                st.session_state.current_question -= 1

                st.rerun()


    with col2:

        if question_index < total_questions - 1:

            if st.button(
                "Next →",
                type="primary",
                use_container_width=True
            ):

                if sample_id not in st.session_state.answers:

                    st.warning(
                        "Please select an answer."
                    )

                else:

                    st.session_state.current_question += 1

                    st.rerun()

        else:

            if st.button(
                "Submit Experiment",
                type="primary",
                use_container_width=True
            ):

                if sample_id not in st.session_state.answers:

                    st.warning(
                        "Please select an answer."
                    )

                else:

                    save_responses()

                    st.session_state.page = "completed"

                    st.rerun()


# ============================================================
# PAGE 3 — COMPLETED
# ============================================================

elif st.session_state.page == "completed":

    st.title(
        "Thank You! 🎉"
    )

    st.success(
        "Your responses have been recorded successfully."
    )

    st.write(
        "Your Participant ID is:"
    )

    st.code(
        st.session_state.participant_id
    )

    st.write(
        "You may now close this page."
    )