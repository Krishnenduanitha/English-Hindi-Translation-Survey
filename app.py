import streamlit as st
import pandas as pd
import os
import random
from datetime import datetime

from streamlit_gsheets import GSheetsConnection


# ============================================================
# PAGE CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="English–Hindi Speech Study",
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


# ============================================================
# GOOGLE SHEETS CONNECTION
# ============================================================

try:

    conn = st.connection(
        "gsheets",
        type=GSheetsConnection
    )

except Exception as e:

    st.error(
        "Unable to connect to Google Sheets."
    )

    st.write(
        "Please check your Streamlit secrets configuration."
    )

    st.stop()


# ============================================================
# CUSTOM CSS
# ============================================================

st.markdown(
    """
    <style>

    /* -------------------------------------------------------
       MAIN PAGE
       ------------------------------------------------------- */

    .block-container {
        max-width: 1000px;
        padding-top: 2rem;
        padding-bottom: 4rem;
    }


    /* -------------------------------------------------------
       ENGLISH SENTENCE
       ------------------------------------------------------- */

    .sentence-box {
        font-size: 28px;
        line-height: 1.6;
        text-align: center;

        padding: 28px;
        margin: 25px 0;

        border-radius: 12px;

        background-color:
            rgba(128, 128, 128, 0.08);
    }


    /* -------------------------------------------------------
       EMPHASIZED WORD
       ------------------------------------------------------- */

    .emphasis-word {
        color: #ff4b4b;
        font-weight: 700;
        text-decoration: underline;
    }


    /* -------------------------------------------------------
       TRANSLATION OPTIONS
       ------------------------------------------------------- */

    div[data-testid="stRadio"] > div {
        gap: 12px;
    }


    div[data-testid="stRadio"] label {
        border: 1px solid
            rgba(128, 128, 128, 0.35);

        border-radius: 12px;

        padding: 18px 20px;

        margin-bottom: 12px;

        width: 100%;

        cursor: pointer;

        transition: 0.2s;
    }


    div[data-testid="stRadio"] label:hover {
        border-color: #ff4b4b;

        background-color:
            rgba(255, 75, 75, 0.05);
    }


    /* -------------------------------------------------------
       INSTRUCTION BOX
       ------------------------------------------------------- */

    .instruction-box {
        padding: 25px;

        border-radius: 12px;

        background-color:
            rgba(70, 130, 180, 0.10);

        border-left: 5px solid #3498db;

        margin-bottom: 25px;

        line-height: 1.7;
    }


    /* -------------------------------------------------------
       THANK YOU
       ------------------------------------------------------- */

    .thank-you {
        text-align: center;

        padding: 50px 20px;
    }


    </style>
    """,
    unsafe_allow_html=True
)


# ============================================================
# LOAD EXCEL FILE
# ============================================================

@st.cache_data
def load_questions():

    if not os.path.exists(EXCEL_FILE):

        st.error(
            "Excel file not found:\n\n"
            + EXCEL_FILE
        )

        st.stop()


    df = pd.read_excel(
        EXCEL_FILE
    )


    df = df.dropna(
        how="all"
    ).reset_index(drop=True)


    required_columns = [

        "sample id",

        "english sentence",

        "hindi translation",

        "machine translation",

        "emphasized word",

        "audiofile"

    ]


    missing_columns = [

        column

        for column in required_columns

        if column not in df.columns

    ]


    if missing_columns:

        st.error(
            "The following columns are missing "
            "from your Excel file:\n\n"
            + "\n".join(missing_columns)
        )

        st.stop()


    return df


df = load_questions()


# ============================================================
# READ GOOGLE SHEET
# ============================================================

def read_responses():

    try:

        data = conn.read(
            worksheet="Responses",
            ttl=0
        )


        if data is None:

            return pd.DataFrame()


        return data


    except Exception:

        return pd.DataFrame()


# ============================================================
# CHECK WHETHER PARTICIPANT EXISTS
# ============================================================

def participant_exists(participant_id):

    data = read_responses()


    if data.empty:

        return False


    if "participant_id" not in data.columns:

        return False


    participant_ids = (
        data["participant_id"]
        .dropna()
        .astype(str)
        .str.strip()
    )


    return participant_id.strip() in (
        participant_ids.values
    )


# ============================================================
# LOAD EXISTING PARTICIPANT PROGRESS
# ============================================================

def load_participant_progress(
    participant_id
):

    data = read_responses()


    if data.empty:

        return False


    if "participant_id" not in data.columns:

        return False


    participant_rows = data[
        data["participant_id"]
        .astype(str)
        .str.strip()
        ==
        participant_id.strip()
    ]


    if participant_rows.empty:

        return False


    # --------------------------------------------------------
    # Recover answers
    # --------------------------------------------------------

    for _, saved_row in participant_rows.iterrows():

        sample_id = saved_row.get(
            "sample_id"
        )


        if pd.isna(sample_id):

            continue


        selected_translation = (
            saved_row.get(
                "selected_translation"
            )
        )


        if pd.isna(selected_translation):

            continue


        selected_source = (
            saved_row.get(
                "selected_translation_type"
            )
        )


        option_number = (
            saved_row.get(
                "selected_option_number"
            )
        )


        emphasis_rating = (
            saved_row.get(
                "emphasis_rating"
            )
        )


        response_time = (
            saved_row.get(
                "response_time_seconds"
            )
        )


        st.session_state.answers[
            sample_id
        ] = {

            "translation":
                str(selected_translation),

            "option_number":
                int(option_number)
                if pd.notna(option_number)
                else None,

            "translation_type":
                str(selected_source)
                if pd.notna(selected_source)
                else "",

            "emphasis_rating":
                int(emphasis_rating)
                if pd.notna(emphasis_rating)
                else None,

            "response_time":
                float(response_time)
                if pd.notna(response_time)
                else None
        }


    # --------------------------------------------------------
    # Find first unanswered question
    # --------------------------------------------------------

    next_question = 0


    for i, question_row in df.iterrows():

        sample_id = question_row[
            "sample id"
        ]


        if sample_id not in st.session_state.answers:

            next_question = i

            break


        next_question = i + 1


    # --------------------------------------------------------
    # Store next question
    # --------------------------------------------------------

    if next_question >= len(df):

        st.session_state.current_question = (
            len(df) - 1
        )

    else:

        st.session_state.current_question = (
            next_question
        )


    return True


# ============================================================
# SESSION STATE
# ============================================================

defaults = {

    "page":
        "welcome",

    "participant_id":
        "",

    "participant_start_time":
        None,

    "current_question":
        0,

    "answers":
        {},

    "randomized_options":
        {},

    "question_start_times":
        {},

    "age_range":
        "",

    "native_languages":
        "",

    "english_proficiency":
        "",

    "hindi_proficiency":
        "",

    "headphones":
        "",

    "hearing_difficulties":
        "",

    "speech_experience":
        "",

    "prosody_understanding":
        "",

    "listening_test_experience":
        ""

}


for key, value in defaults.items():

    if key not in st.session_state:

        st.session_state[key] = value


# ============================================================
# SAVE ONE QUESTION'S RESPONSE
# ============================================================

def save_progress(sample_id):

    if sample_id not in st.session_state.answers:

        return


    answer = st.session_state.answers[
        sample_id
    ]


    question_data = df[
        df["sample id"] == sample_id
    ]


    if question_data.empty:

        return


    row = question_data.iloc[0]


    current_time = datetime.now().strftime(
        "%Y-%m-%d %H:%M:%S"
    )


    new_row = {

        # ----------------------------------------------------
        # PARTICIPANT INFORMATION
        # ----------------------------------------------------

        "participant_id":
            st.session_state.participant_id,

        "age_range":
            st.session_state.age_range,

        "native_languages":
            st.session_state.native_languages,

        "english_proficiency":
            st.session_state.english_proficiency,

        "hindi_proficiency":
            st.session_state.hindi_proficiency,

        "headphones":
            st.session_state.headphones,

        "hearing_difficulties":
            st.session_state.hearing_difficulties,

        "speech_experience":
            st.session_state.speech_experience,

        "prosody_understanding":
            st.session_state.prosody_understanding,

        "listening_test_experience":
            st.session_state.listening_test_experience,

        "participant_start_time":
            st.session_state.participant_start_time,


        # ----------------------------------------------------
        # QUESTION INFORMATION
        # ----------------------------------------------------

        "sample_id":
            sample_id,

        "english_sentence":
            row["english sentence"],

        "emphasized_word":
            row["emphasized word"],

        "audiofile":
            row["audiofile"],


        # ----------------------------------------------------
        # RESPONSE
        # ----------------------------------------------------

        "selected_translation":
            answer["translation"],

        "selected_option_number":
            answer["option_number"],

        # Hidden information for research analysis
        "selected_translation_type":
            answer["translation_type"],

        # 1–5 rating
        "emphasis_rating":
            answer["emphasis_rating"],

        # Time spent on question
        "response_time_seconds":
            answer["response_time"],

        "last_updated":
            current_time
    }


    new_data = pd.DataFrame(
        [new_row]
    )


    # --------------------------------------------------------
    # READ EXISTING DATA
    # --------------------------------------------------------

    existing_data = read_responses()


    if existing_data.empty:

        final_data = new_data

    else:

        # ----------------------------------------------------
        # Remove previous saved copy of this question
        # for this participant.
        #
        # This is important when the participant uses
        # Previous and changes their answer.
        # ----------------------------------------------------

        if (
            "participant_id"
            in existing_data.columns
            and
            "sample_id"
            in existing_data.columns
        ):

            existing_data = existing_data[
                ~(
                    (
                        existing_data[
                            "participant_id"
                        ]
                        .astype(str)
                        .str.strip()
                        ==
                        str(
                            st.session_state.participant_id
                        )
                        .strip()
                    )
                    &
                    (
                        existing_data[
                            "sample_id"
                        ]
                        .astype(str)
                        .str.strip()
                        ==
                        str(sample_id).strip()
                    )
                )
            ]


        final_data = pd.concat(
            [
                existing_data,
                new_data
            ],
            ignore_index=True
        )


    # --------------------------------------------------------
    # UPDATE GOOGLE SHEET
    # --------------------------------------------------------

    conn.update(
        worksheet="Responses",
        data=final_data
    )


# ============================================================
# PAGE 1 — WELCOME
# ============================================================

if st.session_state.page == "welcome":

    st.title(
        "🎧 English–Hindi Speech Study"
    )


    st.markdown(
        """
        <div class="instruction-box">

        <strong>Welcome!</strong>

        <br><br>

        This study is designed to examine the transfer of
        <strong>prosodic emphasis</strong> from English speech
        into <strong>English-to-Hindi translation</strong>.

        <br><br>

        In each trial, you will listen to an English sentence
        containing one or more indicated emphasized words.
        You will then choose the Hindi translation that you
        think best matches the intended meaning, taking the
        emphasis into account.

        <br><br>

        You will also rate how strongly you perceive the
        indicated English word(s) to be emphasized in the
        audio.

        </div>
        """,
        unsafe_allow_html=True
    )


    st.subheader(
        "Participant ID"
    )


    st.write(
        """
        Enter a unique Participant ID that you can remember.
        If you leave the study and return later, enter the
        same ID to resume your progress.
        """
    )


    participant_id_input = st.text_input(
        "Enter your Participant ID",
        placeholder="Example: P001"
    )


    st.write("")


    if st.button(
        "Start / Resume",
        type="primary",
        use_container_width=False
    ):

        if not participant_id_input.strip():

            st.warning(
                "Please enter a Participant ID."
            )

        else:

            participant_id = (
                participant_id_input
                .strip()
            )


            # ------------------------------------------------
            # EXISTING PARTICIPANT
            # ------------------------------------------------

            if participant_exists(
                participant_id
            ):

                st.session_state.participant_id = (
                    participant_id
                )

                st.session_state.answers = {}

                st.session_state.randomized_options = {}

                st.session_state.question_start_times = {}

                found = load_participant_progress(
                    participant_id
                )


                if found:

                    # Check whether all questions
                    # have already been completed.

                    if len(
                        st.session_state.answers
                    ) >= len(df):

                        st.session_state.page = (
                            "already_completed"
                        )

                    else:

                        st.session_state.page = (
                            "experiment"
                        )


                    st.rerun()


            # ------------------------------------------------
            # NEW PARTICIPANT
            # ------------------------------------------------

            else:

                st.session_state.participant_id = (
                    participant_id
                )


                st.session_state.answers = {}

                st.session_state.randomized_options = {}

                st.session_state.question_start_times = {}

                st.session_state.current_question = 0


                st.session_state.participant_start_time = (
                    datetime.now().strftime(
                        "%Y-%m-%d %H:%M:%S"
                    )
                )


                st.session_state.page = "about"


                st.rerun()


# ============================================================
# PAGE 2 — PARTICIPANT INFORMATION
# ============================================================

elif st.session_state.page == "about":

    st.title(
        "📝 About You"
    )


    st.markdown(
        """
        <div class="instruction-box">

        Please provide a few details before beginning
        the listening experiment. These details will help
        us interpret the study results.

        </div>
        """,
        unsafe_allow_html=True
    )


    # --------------------------------------------------------
    # AGE
    # --------------------------------------------------------

    age_range = st.selectbox(
        "Age range",
        [
            "18–24",
            "25–34",
            "35–44",
            "45–54",
            "55+"
        ]
    )


    # --------------------------------------------------------
    # NATIVE LANGUAGE
    # --------------------------------------------------------

    native_languages = st.text_input(
        "Native language(s)",
        placeholder="Example: Malayalam, English"
    )


    # --------------------------------------------------------
    # ENGLISH PROFICIENCY
    # --------------------------------------------------------

    english_proficiency = st.selectbox(
        "English proficiency",
        [
            "Beginner",
            "Intermediate",
            "Advanced",
            "Native / Near-native"
        ]
    )


    # --------------------------------------------------------
    # HINDI PROFICIENCY
    # --------------------------------------------------------

    hindi_proficiency = st.selectbox(
        "Hindi proficiency",
        [
            "Beginner",
            "Intermediate",
            "Advanced",
            "Native / Near-native"
        ]
    )


    # --------------------------------------------------------
    # HEADPHONES
    # --------------------------------------------------------

    headphones = st.radio(
        "Are you using headphones?",
        [
            "Yes",
            "No"
        ],
        horizontal=True
    )


    # --------------------------------------------------------
    # HEARING
    # --------------------------------------------------------

    hearing_difficulties = st.radio(
        "Do you have any known hearing difficulties?",
        [
            "No",
            "Yes"
        ],
        horizontal=True
    )


    st.divider()


    st.subheader(
        "Speech & Listening Experience"
    )


    # --------------------------------------------------------
    # SPEECH EXPERIENCE
    # --------------------------------------------------------

    speech_experience = st.radio(
        "Do you have any experience in speech, linguistics, or audio research?",
        [
            "No",
            "Yes"
        ],
        horizontal=True
    )


    # --------------------------------------------------------
    # PROSODY KNOWLEDGE
    # --------------------------------------------------------

    prosody_understanding = st.radio(
        "Do you understand what prosody (stress, rhythm, and intonation) means?",
        [
            "No",
            "Somewhat",
            "Yes"
        ],
        horizontal=True
    )


    # --------------------------------------------------------
    # PREVIOUS LISTENING TEST EXPERIENCE
    # --------------------------------------------------------

    listening_test_experience = st.radio(
        "Have you participated in perceptual or listening tests before?",
        [
            "No",
            "Yes"
        ],
        horizontal=True
    )


    st.write("")


    if st.button(
        "Continue →",
        type="primary"
    ):

        # Store participant information

        st.session_state.age_range = (
            age_range
        )

        st.session_state.native_languages = (
            native_languages
        )

        st.session_state.english_proficiency = (
            english_proficiency
        )

        st.session_state.hindi_proficiency = (
            hindi_proficiency
        )

        st.session_state.headphones = (
            headphones
        )

        st.session_state.hearing_difficulties = (
            hearing_difficulties
        )

        st.session_state.speech_experience = (
            speech_experience
        )

        st.session_state.prosody_understanding = (
            prosody_understanding
        )

        st.session_state.listening_test_experience = (
            listening_test_experience
        )


        st.session_state.page = (
            "instructions"
        )


        st.rerun()


# ============================================================
# PAGE 3 — INSTRUCTIONS
# ============================================================

elif st.session_state.page == "instructions":

    st.title(
        "📋 Instructions"
    )


    st.markdown(
        """
        <div class="instruction-box">

        <h3>Task</h3>

        <ol>

        <li>
        Listen carefully to the English sentence.
        </li>

        <li>
        The English word(s) intended to carry emphasis
        will be highlighted.
        </li>

        <li>
        Pay attention to how the highlighted word(s)
        are emphasized in the audio.
        </li>

        <li>
        Choose the Hindi translation that best conveys
        the intended meaning of the English sentence,
        considering the indicated emphasis.
        </li>

        <li>
        After selecting a translation, rate how strongly
        you perceive the indicated word(s) to be emphasized
        in the audio.
        </li>

        </ol>


        <h3>Emphasis Rating</h3>

        <p>
        <strong>1</strong> — Not emphasized at all
        <br>

        <strong>2</strong> — Slightly emphasized
        <br>

        <strong>3</strong> — Moderately emphasized
        <br>

        <strong>4</strong> — Strongly emphasized
        <br>

        <strong>5</strong> — Very strongly emphasized
        </p>


        <h3>Important</h3>

        <ul>

        <li>
        Focus specifically on the indicated word(s).
        </li>

        <li>
        Do not judge the speaker based on voice,
        gender, accent, or overall loudness.
        </li>

        <li>
        Listen to the complete sentence before making
        your decision.
        </li>

        <li>
        Choose the translation based on meaning and
        the role of emphasis, rather than trying to
        identify which system produced it.
        </li>

        </ul>

        </div>
        """,
        unsafe_allow_html=True
    )


    if st.button(
        "Begin Experiment →",
        type="primary",
        use_container_width=True
    ):

        st.session_state.page = (
            "experiment"
        )

        st.rerun()


# ============================================================
# PAGE 4 — EXPERIMENT
# ============================================================

elif st.session_state.page == "experiment":

    question_index = (
        st.session_state.current_question
    )


    total_questions = len(df)


    # --------------------------------------------------------
    # SAFETY CHECK
    # --------------------------------------------------------

    if question_index >= total_questions:

        st.session_state.page = (
            "completed"
        )

        st.rerun()


    row = df.iloc[
        question_index
    ]


    sample_id = row[
        "sample id"
    ]


    # ========================================================
    # HEADER
    # ========================================================

    st.title(
        "🎧 Translation & Emphasis Test"
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


    # ========================================================
    # ENGLISH SENTENCE
    # ========================================================

    english_sentence = str(
        row["english sentence"]
    )


    emphasized_word = str(
        row["emphasized word"]
    )


    # --------------------------------------------------------
    # Highlight emphasized words
    # --------------------------------------------------------

    highlighted_sentence = (
        english_sentence
    )


    emphasized_words = [

        word.strip()

        for word
        in emphasized_word.split(",")

        if word.strip()

    ]


    for word in emphasized_words:

        highlighted_sentence = (
            highlighted_sentence.replace(
                word,
                (
                    '<span class="emphasis-word">'
                    + word
                    + '</span>'
                )
            )
        )


    st.markdown(
        f"""
        <div class="sentence-box">
        "{highlighted_sentence}"
        </div>
        """,
        unsafe_allow_html=True
    )


    st.markdown(
        f"""
        **Emphasized English word(s):**
        <span class="emphasis-word">
        {emphasized_word}
        </span>
        """,
        unsafe_allow_html=True
    )


    # ========================================================
    # AUDIO
    # ========================================================

    st.subheader(
        "🔊 Listen to the audio"
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


    # ========================================================
    # TRANSLATION OPTIONS
    # ========================================================

    st.subheader(
        "Which Hindi translation best matches the intended meaning?"
    )


    st.write(
        """
        Consider the indicated emphasis when making
        your choice.
        """
    )


    # --------------------------------------------------------
    # ONLY TWO OPTIONS
    # --------------------------------------------------------

    options = [

        (
            "Hindi Translation",
            row["hindi translation"]
        ),

        (
            "Machine Translation",
            row["machine translation"]
        )

    ]


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
    # Check that both options exist
    # --------------------------------------------------------

    if len(valid_options) < 2:

        st.error(
            "This question does not contain both "
            "Hindi translation options."
        )

        st.write(
            "Please make sure the Excel file contains "
            "both 'hindi translation' and "
            "'machine translation'."
        )

        st.stop()


    # ========================================================
    # RANDOMIZE THE TWO OPTIONS
    # ========================================================

    if sample_id not in (
        st.session_state.randomized_options
    ):

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


    display_options = [

        translation

        for source, translation
        in shuffled_options

    ]


    # ========================================================
    # QUESTION TIMER
    # ========================================================

    if sample_id not in (
        st.session_state.question_start_times
    ):

        st.session_state.question_start_times[
            sample_id
        ] = datetime.now()


    # ========================================================
    # PREVIOUS ANSWER
    # ========================================================

    previous_answer = (
        st.session_state.answers.get(
            sample_id
        )
    )


    default_index = None


    if previous_answer:

        previous_translation = (
            previous_answer["translation"]
        )


        if previous_translation in (
            display_options
        ):

            default_index = (
                display_options.index(
                    previous_translation
                )
            )


    # ========================================================
    # TRANSLATION SELECTION
    # ========================================================

    selected_translation = st.radio(

        "Hindi translation options",

        display_options,

        index=default_index,

        key=f"translation_{sample_id}",

        label_visibility="collapsed"
    )


    # ========================================================
    # EMPHASIS RATING
    # ========================================================

    st.write("")


    st.subheader(
        "How strongly are the indicated English word(s) emphasized in the audio?"
    )


    st.write(
        """
        Rate only the perceived emphasis of the
        indicated word(s), not the overall loudness
        of the speaker.
        """
    )


    rating_labels = [

        "1 — Not emphasized at all",

        "2 — Slightly emphasized",

        "3 — Moderately emphasized",

        "4 — Strongly emphasized",

        "5 — Very strongly emphasized"

    ]


    previous_rating = None


    if previous_answer:

        previous_rating = (
            previous_answer.get(
                "emphasis_rating"
            )
        )


    rating_index = None


    if previous_rating in [
        1,
        2,
        3,
        4,
        5
    ]:

        rating_index = (
            previous_rating - 1
        )


    emphasis_rating_label = st.radio(

        "Emphasis rating",

        rating_labels,

        index=rating_index,

        key=f"rating_{sample_id}",

        label_visibility="collapsed"
    )


    emphasis_rating = (
        rating_labels.index(
            emphasis_rating_label
        ) + 1
    )


    # ========================================================
    # STORE ANSWER IN SESSION
    # ========================================================

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

            "emphasis_rating":
                emphasis_rating,

            "response_time":
                round(
                    response_time,
                    2
                )
        }


    st.divider()


    # ========================================================
    # NAVIGATION
    # ========================================================

    col1, col2 = st.columns(2)


    # ========================================================
    # PREVIOUS BUTTON
    # ========================================================

    with col1:

        if question_index > 0:

            if st.button(
                "← Previous",
                use_container_width=True
            ):

                st.session_state.current_question -= 1

                st.rerun()


    # ========================================================
    # NEXT BUTTON
    # ========================================================

    with col2:

        if question_index < total_questions - 1:

            if st.button(
                "Next →",
                type="primary",
                use_container_width=True
            ):

                if not selected_translation:

                    st.warning(
                        "Please select a Hindi translation."
                    )

                elif not emphasis_rating:

                    st.warning(
                        "Please provide an emphasis rating."
                    )

                else:

                    # Save immediately to Google Sheets
                    save_progress(
                        sample_id
                    )


                    st.session_state.current_question += 1


                    st.rerun()


        # ====================================================
        # FINAL SUBMIT
        # ====================================================

        else:

            if st.button(
                "Submit Experiment",
                type="primary",
                use_container_width=True
            ):

                if not selected_translation:

                    st.warning(
                        "Please select a Hindi translation."
                    )

                elif not emphasis_rating:

                    st.warning(
                        "Please provide an emphasis rating."
                    )

                else:

                    try:

                        # Save final question
                        save_progress(
                            sample_id
                        )


                        st.session_state.page = (
                            "completed"
                        )


                        st.rerun()


                    except Exception as e:

                        st.error(
                            "Unable to save your response."
                        )

                        st.exception(e)


# ============================================================
# PAGE 5 — ALREADY COMPLETED
# ============================================================

elif st.session_state.page == "already_completed":

    st.markdown(
        """
        <div class="thank-you">

        <h1>Study Already Completed</h1>

        <p>
        This Participant ID has already completed
        the questionnaire.
        </p>

        <p>
        Thank you for your participation.
        </p>

        </div>
        """,
        unsafe_allow_html=True
    )


# ============================================================
# PAGE 6 — COMPLETED
# ============================================================

elif st.session_state.page == "completed":

    st.markdown(
        """
        <div class="thank-you">

        <h1>Thank You! 🎉</h1>

        <h3>
        Your responses have been recorded successfully.
        </h3>

        <p>
        Thank you for taking part in this study.
        </p>

        </div>
        """,
        unsafe_allow_html=True
    )


    st.info(
        f"Participant ID: "
        f"{st.session_state.participant_id}"
    )


    st.write(
        "You may now close this page."
    )