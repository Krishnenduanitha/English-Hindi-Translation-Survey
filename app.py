import streamlit as st
import pandas as pd
import os
import random
import hashlib
import html
import gspread

from datetime import datetime
from google.oauth2.service_account import Credentials


# ============================================================
# PAGE CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="English-Hindi Translation Study",
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
# GOOGLE SHEETS CONFIGURATION
# ============================================================

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]


@st.cache_resource
def get_google_sheet():

    try:

        service_account_info = dict(
            st.secrets["connections"]["gsheets"]
        )

        credentials = Credentials.from_service_account_info(
            service_account_info,
            scopes=SCOPES
        )

        client = gspread.authorize(credentials)

        spreadsheet_url = service_account_info["spreadsheet"]

        spreadsheet = client.open_by_url(
            spreadsheet_url
        )

        worksheet_name = service_account_info.get(
            "worksheet",
            "Responses"
        )

        worksheet = spreadsheet.worksheet(
            worksheet_name
        )

        return worksheet

    except Exception as e:

        st.error(
            "Unable to connect to Google Sheets.\n\n"
            f"{e}"
        )

        st.stop()


worksheet = get_google_sheet()


# ============================================================
# GOOGLE SHEET HEADERS
# ============================================================

HEADERS = [

    "participant_name",

    "age_range",

    "native_language",

    "english_proficiency",

    "hindi_proficiency",

    "headphones",

    "hearing_difficulties",

    "speech_experience",

    "prosody_understanding",

    "listening_test_experience",

    "sample_id",

    "english_sentence",

    "emphasized_word",

    "audiofile",

    "prosodic_feature",

    "selected_translation",

    "selected_translation_type",

    "prosodic_rating",

    "response_time_seconds",

    "remarks",

    "last_updated"
]


# ============================================================
# INITIALIZE GOOGLE SHEET
# ============================================================

def initialize_sheet():

    try:

        values = worksheet.get_all_values()

        # ----------------------------------------------------
        # Completely empty sheet
        # ----------------------------------------------------

        if not values:

            worksheet.append_row(
                HEADERS,
                value_input_option="USER_ENTERED"
            )

            return

        # ----------------------------------------------------
        # Existing sheet
        # ----------------------------------------------------

        existing_headers = values[0]

        missing_headers = [
            header
            for header in HEADERS
            if header not in existing_headers
        ]

        if missing_headers:

            start_column = len(existing_headers) + 1

            for offset, header in enumerate(
                missing_headers
            ):

                worksheet.update_cell(
                    1,
                    start_column + offset,
                    header
                )

    except Exception as e:

        st.error(
            "Could not initialize Google Sheet.\n\n"
            f"{e}"
        )

        st.stop()


initialize_sheet()


# ============================================================
# CUSTOM CSS
# ============================================================

st.markdown(
    """
    <style>

    /* ======================================================
       GENERAL
       ====================================================== */

    .main-title {
        text-align: center;
        font-size: 32px;
        font-weight: 700;
        margin-bottom: 10px;
        line-height: 1.3;
    }

    .subtitle {
        text-align: center;
        font-size: 18px;
        font-weight: 500;
        margin-bottom: 30px;
        line-height: 1.5;
    }

    .section-title {
        font-size: 24px;
        font-weight: 650;
        margin-top: 20px;
        margin-bottom: 15px;
        line-height: 1.4;
    }


    /* ======================================================
       SENTENCE
       ====================================================== */

    .sentence-box {
        padding: 20px;
        border-radius: 12px;
        border: 1px solid rgba(128, 128, 128, 0.35);
        font-size: 20px;
        line-height: 1.6;
        margin-top: 15px;
        margin-bottom: 20px;
    }


    /* ======================================================
       EMPHASIZED WORD
       ====================================================== */

    .emphasis-word {
        color: #ff3333;
        font-weight: 700;
        text-decoration: underline;
    }


    /* ======================================================
       TRANSLATION SECTION
       ====================================================== */

    .translation-note {
        font-size: 15px;
        margin-bottom: 14px;
    }


    /* ======================================================
       PROGRESS
       ====================================================== */

    .progress-text {
        text-align: center;
        font-size: 15px;
        margin-bottom: 10px;
    }


    /* ======================================================
       TRANSLATION RADIO CARDS
       
       IMPORTANT:
       This CSS ONLY affects the translation radio widget.
       Participant-information radios remain normal.
       ====================================================== */

    div.st-key-translation_choice [role="radiogroup"] {
        gap: 18px !important;
    }

    div.st-key-translation_choice
    [role="radiogroup"] > label {

        display: flex !important;

        align-items: flex-start !important;

        width: 100% !important;

        min-height: 150px !important;

        box-sizing: border-box !important;

        padding: 28px 30px !important;

        margin-bottom: 16px !important;

        border: 1px solid rgba(128, 128, 128, 0.45) !important;

        border-radius: 14px !important;

        background-color: transparent !important;

        cursor: pointer !important;

        transition:
            background-color 0.2s ease,
            border-color 0.2s ease,
            transform 0.15s ease,
            box-shadow 0.2s ease !important;
    }


    /* Hover */

    div.st-key-translation_choice
    [role="radiogroup"] > label:hover {

        background-color: rgba(80, 140, 255, 0.12) !important;

        border-color: rgba(80, 140, 255, 0.85) !important;

        transform: translateY(-2px) !important;

        box-shadow:
            0 4px 12px rgba(80, 140, 255, 0.15) !important;
    }


    /* Translation text */

    div.st-key-translation_choice
    [role="radiogroup"] > label p {

        white-space: normal !important;

        overflow: visible !important;

        text-overflow: clip !important;

        word-break: normal !important;

        overflow-wrap: anywhere !important;

        line-height: 1.7 !important;

        font-size: 19px !important;

        margin: 0 !important;

        width: 100% !important;
    }


    /* Radio circle */

    div.st-key-translation_choice
    [role="radiogroup"] > label
    [data-baseweb="radio"] {

        margin-top: 3px !important;

        flex-shrink: 0 !important;
    }


    /* ======================================================
       MOBILE
       ====================================================== */

    @media (max-width: 700px) {

        div.st-key-translation_choice
        [role="radiogroup"] > label {

            min-height: 120px !important;

            padding: 20px !important;
        }

        div.st-key-translation_choice
        [role="radiogroup"] > label p {

            font-size: 17px !important;

            line-height: 1.6 !important;
        }

    }

    </style>
    """,
    unsafe_allow_html=True
)


# ============================================================
# LOAD QUESTIONS FROM EXCEL
# ============================================================

@st.cache_data
def load_questions():

    if not os.path.exists(EXCEL_FILE):

        st.error(
            f"Excel file not found:\n{EXCEL_FILE}"
        )

        st.stop()

    df = pd.read_excel(
        EXCEL_FILE
    )

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
            "The following required columns are missing "
            "from the Excel file:\n\n"
            +
            "\n".join(missing_columns)
        )

        st.stop()

    df = df.fillna("")

    return df


questions = load_questions()


# ============================================================
# DETECT PROSODIC FEATURE COLUMN
# ============================================================

def get_prosodic_feature(row):

    possible_columns = [

        "prosodic feature",

        "prosodic_feature",

        "feature",

        "feature type",

        "feature_type",

        "prosody",

        "prosody type",

        "prosody_type"
    ]

    for column in possible_columns:

        if column in row.index:

            value = str(
                row[column]
            ).strip()

            if value:

                return value

    # If the Excel does not have a separate feature column,
    # do not invent a specific feature for the question.

    return ""


# ============================================================
# READ RESPONSES
# ============================================================

def read_responses():

    try:

        records = worksheet.get_all_records()

        if not records:

            return pd.DataFrame()

        return pd.DataFrame(
            records
        )

    except Exception as e:

        st.error(
            "Unable to read responses from Google Sheets.\n\n"
            f"{e}"
        )

        return pd.DataFrame()


# ============================================================
# CHECK WHETHER PARTICIPANT EXISTS
# ============================================================

def participant_exists(participant_name):

    responses = read_responses()

    if responses.empty:

        return False

    if "participant_name" not in responses.columns:

        return False

    names = (
        responses[
            "participant_name"
        ]
        .astype(str)
        .str.strip()
        .str.lower()
    )

    return (
        participant_name.strip().lower()
        in names.values
    )


# ============================================================
# LOAD PARTICIPANT PROGRESS
# ============================================================

def load_participant_progress(participant_name):

    responses = read_responses()

    if responses.empty:

        st.session_state.answers = {}

        st.session_state.current_question = 0

        return

    if "participant_name" not in responses.columns:

        st.session_state.answers = {}

        st.session_state.current_question = 0

        return

    participant_rows = responses[
        responses[
            "participant_name"
        ]
        .astype(str)
        .str.strip()
        .str.lower()
        ==
        participant_name.strip().lower()
    ]

    answers = {}

    # ========================================================
    # RESTORE ANSWERS
    # ========================================================

    for _, row in participant_rows.iterrows():

        sample_id = str(
            row.get(
                "sample_id",
                ""
            )
        ).strip()

        if not sample_id:

            continue

        answers[sample_id] = {

            "selected_translation":
                str(
                    row.get(
                        "selected_translation",
                        ""
                    )
                ),

            "selected_translation_type":
                str(
                    row.get(
                        "selected_translation_type",
                        ""
                    )
                ),

            "prosodic_rating":
                row.get(
                    "prosodic_rating",
                    ""
                ),

            "response_time_seconds":
                row.get(
                    "response_time_seconds",
                    ""
                )
        }

    st.session_state.answers = answers


    # ========================================================
    # RESTORE PARTICIPANT INFORMATION
    # ========================================================

    if not participant_rows.empty:

        latest_row = participant_rows.iloc[-1]

        st.session_state.age_range = str(
            latest_row.get(
                "age_range",
                ""
            )
        )

        st.session_state.native_language = str(
            latest_row.get(
                "native_language",
                ""
            )
        )

        st.session_state.english_proficiency = str(
            latest_row.get(
                "english_proficiency",
                ""
            )
        )

        st.session_state.hindi_proficiency = str(
            latest_row.get(
                "hindi_proficiency",
                ""
            )
        )

        st.session_state.headphones = str(
            latest_row.get(
                "headphones",
                ""
            )
        )

        st.session_state.hearing_difficulties = str(
            latest_row.get(
                "hearing_difficulties",
                ""
            )
        )

        st.session_state.speech_experience = str(
            latest_row.get(
                "speech_experience",
                ""
            )
        )

        st.session_state.prosody_understanding = str(
            latest_row.get(
                "prosody_understanding",
                ""
            )
        )

        st.session_state.listening_test_experience = str(
            latest_row.get(
                "listening_test_experience",
                ""
            )
        )

        # Restore remarks if available

        st.session_state.remarks = str(
            latest_row.get(
                "remarks",
                ""
            )
        )


    # ========================================================
    # FIND FIRST UNANSWERED QUESTION
    # ========================================================

    first_unanswered = len(
        questions
    )

    for index, row in questions.iterrows():

        sample_id = str(
            row["sample id"]
        ).strip()

        if sample_id not in answers:

            first_unanswered = index

            break

    st.session_state.current_question = (
        first_unanswered
    )


# ============================================================
# CREATE RESPONSE DATA
# ============================================================

def create_response_data(sample_id):

    question_index = (
        st.session_state.current_question
    )

    row = questions.iloc[
        question_index
    ]

    answer = (
        st.session_state.answers[
            str(sample_id)
        ]
    )

    response_data = {

        "participant_name":
            st.session_state.participant_name,

        "age_range":
            st.session_state.age_range,

        "native_language":
            st.session_state.native_language,

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

        "sample_id":
            str(sample_id),

        "english_sentence":
            str(
                row["english sentence"]
            ),

        "emphasized_word":
            str(
                row["emphasized word"]
            ),

        "audiofile":
            str(
                row["audiofile"]
            ),

        "prosodic_feature":
            get_prosodic_feature(row),

        "selected_translation":
            answer[
                "selected_translation"
            ],

        "selected_translation_type":
            answer[
                "selected_translation_type"
            ],

        "prosodic_rating":
            answer[
                "prosodic_rating"
            ],

        "response_time_seconds":
            answer[
                "response_time_seconds"
            ],

        "remarks":
            st.session_state.remarks,

        "last_updated":
            datetime.now().strftime(
                "%Y-%m-%d %H:%M:%S"
            )
    }

    return response_data


# ============================================================
# SAVE QUESTION RESPONSE
# ============================================================

def save_progress(sample_id):

    response_data = create_response_data(
        sample_id
    )

    try:

        all_values = worksheet.get_all_values()

        # ====================================================
        # GET CURRENT HEADERS
        # ====================================================

        if not all_values:

            worksheet.append_row(
                HEADERS,
                value_input_option="USER_ENTERED"
            )

            headers = HEADERS

            all_values = [
                HEADERS
            ]

        else:

            headers = all_values[0]


        # ====================================================
        # MAKE SURE REQUIRED HEADERS EXIST
        # ====================================================

        missing_headers = [

            header

            for header in HEADERS

            if header not in headers
        ]

        if missing_headers:

            start_column = len(headers) + 1

            for offset, header in enumerate(
                missing_headers
            ):

                worksheet.update_cell(
                    1,
                    start_column + offset,
                    header
                )

            headers = (
                headers
                +
                missing_headers
            )


        # ====================================================
        # BUILD ROW
        # ====================================================

        new_row = [

            response_data.get(
                header,
                ""
            )

            for header in headers
        ]


        # ====================================================
        # FIND EXISTING PARTICIPANT + SAMPLE
        # ====================================================

        participant_index = headers.index(
            "participant_name"
        )

        sample_index = headers.index(
            "sample_id"
        )

        existing_row_number = None

        for row_number, existing_row in enumerate(
            all_values[1:],
            start=2
        ):

            existing_participant = ""

            existing_sample = ""

            if participant_index < len(
                existing_row
            ):

                existing_participant = (
                    str(
                        existing_row[
                            participant_index
                        ]
                    )
                    .strip()
                    .lower()
                )

            if sample_index < len(
                existing_row
            ):

                existing_sample = (
                    str(
                        existing_row[
                            sample_index
                        ]
                    )
                    .strip()
                )

            if (
                existing_participant
                ==
                st.session_state.participant_name
                .strip()
                .lower()
                and
                existing_sample
                ==
                str(sample_id).strip()
            ):

                existing_row_number = (
                    row_number
                )

                break


        # ====================================================
        # UPDATE EXISTING RESPONSE
        # ====================================================

        if existing_row_number is not None:

            end_column_letter = (
                gspread.utils.rowcol_to_a1(
                    1,
                    len(headers)
                ).rstrip("1")
            )

            worksheet.update(
                f"A{existing_row_number}:"
                f"{end_column_letter}"
                f"{existing_row_number}",
                [new_row],
                value_input_option="USER_ENTERED"
            )

        # ====================================================
        # ADD NEW RESPONSE
        # ====================================================

        else:

            worksheet.append_row(
                new_row,
                value_input_option="USER_ENTERED"
            )

    except Exception as e:

        st.error(
            "Your response could not be saved.\n\n"
            f"{e}"
        )


# ============================================================
# SAVE PARTICIPANT REMARKS
# ============================================================

def save_remarks():

    participant_name = (
        st.session_state.participant_name
        .strip()
        .lower()
    )

    remarks = (
        st.session_state.remarks
    )

    try:

        all_values = worksheet.get_all_values()

        if not all_values:

            return False

        headers = all_values[0]

        # ----------------------------------------------------
        # Make sure remarks column exists
        # ----------------------------------------------------

        if "remarks" not in headers:

            worksheet.update_cell(
                1,
                len(headers) + 1,
                "remarks"
            )

            headers.append(
                "remarks"
            )

        remarks_index = headers.index(
            "remarks"
        )

        participant_index = headers.index(
            "participant_name"
        )

        last_updated_index = None

        if "last_updated" in headers:

            last_updated_index = headers.index(
                "last_updated"
            )


        # ----------------------------------------------------
        # UPDATE ALL ROWS FOR PARTICIPANT
        # ----------------------------------------------------

        for row_number, existing_row in enumerate(
            all_values[1:],
            start=2
        ):

            if participant_index >= len(
                existing_row
            ):

                continue

            existing_participant = (
                str(
                    existing_row[
                        participant_index
                    ]
                )
                .strip()
                .lower()
            )

            if (
                existing_participant
                ==
                participant_name
            ):

                worksheet.update_cell(
                    row_number,
                    remarks_index + 1,
                    remarks
                )

                if last_updated_index is not None:

                    worksheet.update_cell(
                        row_number,
                        last_updated_index + 1,
                        datetime.now().strftime(
                            "%Y-%m-%d %H:%M:%S"
                        )
                    )

        return True

    except Exception as e:

        st.error(
            "Unable to save your remarks.\n\n"
            f"{e}"
        )

        return False


# ============================================================
# HIGHLIGHT EMPHASIZED WORDS
# ============================================================

def highlight_emphasis(sentence, emphasized):

    sentence = str(sentence)

    emphasized = str(
        emphasized
    ).strip()

    if not emphasized:

        return html.escape(sentence)

    words = [
        emphasized
    ]

    # Support comma-separated and slash-separated entries

    for separator in [
        ",",
        "/"
    ]:

        new_words = []

        for word in words:

            new_words.extend(
                word.split(
                    separator
                )
            )

        words = new_words

    words = [

        word.strip()

        for word in words

        if word.strip()
    ]

    # Longest phrases first

    words = sorted(
        words,
        key=len,
        reverse=True
    )

    highlighted_sentence = html.escape(
        sentence
    )

    for word in words:

        escaped_word = html.escape(
            word
        )

        highlighted_sentence = (
            highlighted_sentence.replace(
                escaped_word,
                (
                    "<span class='emphasis-word'>"
                    +
                    escaped_word
                    +
                    "</span>"
                )
            )
        )

    return highlighted_sentence


# ============================================================
# STABLE RANDOMIZATION
# ============================================================

def get_randomized_options(
    participant_name,
    sample_id,
    options
):

    seed_string = (
        participant_name
        .strip()
        .lower()
        +
        "_"
        +
        str(sample_id)
    )

    seed = int(
        hashlib.sha256(
            seed_string.encode(
                "utf-8"
            )
        ).hexdigest(),
        16
    )

    rng = random.Random(
        seed
    )

    shuffled = options.copy()

    rng.shuffle(
        shuffled
    )

    return shuffled


# ============================================================
# SESSION STATE
# ============================================================

defaults = {

    "page":
        "welcome",

    "participant_name":
        "",

    "participant_start_time":
        None,

    "current_question":
        0,

    "answers":
        {},

    "randomized_options":
        {},

    "translation_selections":
        {},

    "rating_selections":
        {},

    "question_start_times":
        {},

    "remarks":
        "",

    "age_range":
        "",

    "native_language":
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
# WELCOME PAGE
# ============================================================

if st.session_state.page == "welcome":

    st.markdown(
        '<div class="main-title">'
        'English-to-Hindi Translation Study'
        '</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="subtitle">'
        'A study on the transfer of prosodic features '
        'from English speech into Hindi translation'
        '</div>',
        unsafe_allow_html=True
    )


    # ========================================================
    # ABOUT THE STUDY
    # ========================================================

    st.markdown(
        """
        ### About the Study

        This study is designed to examine how prosodic
        features in English speech are reflected in
        English-to-Hindi translation.

        In particular, the study focuses on two aspects of
        prosody: **emphasis** and **rising contour**. These
        features can contribute to how a speaker's intended
        meaning is perceived and interpreted.

        In each trial, you will listen to an English sentence
        in an audio recording. Certain words or parts of the
        sentence may carry prosodic prominence or a rising
        intonation pattern.

        After listening to the audio, you will first rate
        how strongly you perceive the relevant prosodic
        feature in the speech. You will then choose the
        Hindi translation that you think best matches the
        intended meaning of the English sentence, taking
        the prosodic information into account.

        Your responses will help us understand how prosodic
        information in English speech, including emphasis
        and rising contour, is reflected in English-to-Hindi
        translation.
        """
    )


    st.markdown("---")


    # ========================================================
    # NAME
    # ========================================================

    st.markdown(
        "### Enter Your Name"
    )

    name_input = st.text_input(
        "Your name",
        placeholder="Enter your name",
        key="name_input"
    )

    st.info(
        "Please use the same name if you return later "
        "and need to resume the study."
    )


    if st.button(
        "Start / Resume Study",
        use_container_width=True
    ):

        name = name_input.strip()

        if not name:

            st.warning(
                "Please enter your name before continuing."
            )

        else:

            st.session_state.participant_name = name

            # =================================================
            # EXISTING PARTICIPANT
            # =================================================

            if participant_exists(name):

                load_participant_progress(
                    name
                )

                if (
                    st.session_state.current_question
                    >=
                    len(questions)
                ):

                    st.session_state.page = (
                        "completed"
                    )

                else:

                    st.session_state.page = (
                        "experiment"
                    )

                    st.session_state.question_start_times[
                        st.session_state.current_question
                    ] = datetime.now()

            # =================================================
            # NEW PARTICIPANT
            # =================================================

            else:

                st.session_state.current_question = 0

                st.session_state.answers = {}

                st.session_state.randomized_options = {}

                st.session_state.translation_selections = {}

                st.session_state.rating_selections = {}

                st.session_state.remarks = ""

                st.session_state.page = (
                    "participant_info"
                )

            st.session_state.participant_start_time = (
                datetime.now()
            )

            st.rerun()


# ============================================================
# PARTICIPANT INFORMATION PAGE
# ============================================================

elif st.session_state.page == "participant_info":

    st.markdown(
        '<div class="section-title">'
        'Participant Information'
        '</div>',
        unsafe_allow_html=True
    )

    st.write(
        "Please provide the following information before "
        "beginning the listening experiment."
    )


    # ========================================================
    # AGE
    # ========================================================

    age_options = [

        "18–20",
        "21–25",
        "26–30",
        "31–40",
        "41–50",
        "51+"
    ]

    current_age = (

        st.session_state.age_range

        if st.session_state.age_range
        in age_options

        else age_options[0]
    )

    st.session_state.age_range = st.selectbox(
        "Age range",
        age_options,
        index=age_options.index(
            current_age
        )
    )


    # ========================================================
    # NATIVE LANGUAGE
    # ========================================================

    st.session_state.native_language = st.text_input(
        "Native language(s)",
        value=st.session_state.native_language
    )


    # ========================================================
    # ENGLISH PROFICIENCY
    # ========================================================

    english_options = [

        "Beginner",
        "Intermediate",
        "Advanced",
        "Native / Near-native"
    ]

    current_english = (

        st.session_state.english_proficiency

        if st.session_state.english_proficiency
        in english_options

        else english_options[0]
    )

    st.session_state.english_proficiency = st.selectbox(
        "English proficiency",
        english_options,
        index=english_options.index(
            current_english
        )
    )


    # ========================================================
    # HINDI PROFICIENCY
    # ========================================================

    hindi_options = [

        "Beginner",
        "Intermediate",
        "Advanced",
        "Native / Near-native"
    ]

    current_hindi = (

        st.session_state.hindi_proficiency

        if st.session_state.hindi_proficiency
        in hindi_options

        else hindi_options[0]
    )

    st.session_state.hindi_proficiency = st.selectbox(
        "Hindi proficiency",
        hindi_options,
        index=hindi_options.index(
            current_hindi
        )
    )


    # ========================================================
    # HEADPHONES
    # ========================================================

    headphones_options = [
        "Yes",
        "No"
    ]

    current_headphones = (

        st.session_state.headphones

        if st.session_state.headphones
        in headphones_options

        else headphones_options[0]
    )

    st.session_state.headphones = st.radio(
        "Are you using headphones or earphones?",
        headphones_options,
        index=headphones_options.index(
            current_headphones
        )
    )


    # ========================================================
    # HEARING DIFFICULTIES
    # ========================================================

    hearing_options = [
        "No",
        "Yes"
    ]

    current_hearing = (

        st.session_state.hearing_difficulties

        if st.session_state.hearing_difficulties
        in hearing_options

        else hearing_options[0]
    )

    st.session_state.hearing_difficulties = st.radio(
        "Do you have any difficulty hearing speech?",
        hearing_options,
        index=hearing_options.index(
            current_hearing
        )
    )


    # ========================================================
    # SPEECH / LINGUISTICS EXPERIENCE
    # ========================================================

    experience_options = [
        "No",
        "Yes"
    ]

    current_experience = (

        st.session_state.speech_experience

        if st.session_state.speech_experience
        in experience_options

        else experience_options[0]
    )

    st.session_state.speech_experience = st.radio(
        "Do you have experience in speech, "
        "linguistics, audio, or related research?",
        experience_options,
        index=experience_options.index(
            current_experience
        )
    )


    # ========================================================
    # PROSODY FAMILIARITY
    # ========================================================

    prosody_options = [

        "Not familiar",
        "Somewhat familiar",
        "Familiar",
        "Very familiar"
    ]

    current_prosody = (

        st.session_state.prosody_understanding

        if st.session_state.prosody_understanding
        in prosody_options

        else prosody_options[0]
    )

    st.session_state.prosody_understanding = st.radio(
        "How familiar are you with the concept of prosody in speech?",
        prosody_options,
        index=prosody_options.index(
            current_prosody
        )
    )


    # ========================================================
    # LISTENING TEST EXPERIENCE
    # ========================================================

    listening_options = [
        "No",
        "Yes"
    ]

    current_listening = (

        st.session_state.listening_test_experience

        if st.session_state.listening_test_experience
        in listening_options

        else listening_options[0]
    )

    st.session_state.listening_test_experience = st.radio(
        "Have you participated in a listening or "
        "speech perception experiment before?",
        listening_options,
        index=listening_options.index(
            current_listening
        )
    )


    st.markdown("---")


    if st.button(
        "Continue",
        use_container_width=True
    ):

        st.session_state.page = (
            "instructions"
        )

        st.rerun()


# ============================================================
# INSTRUCTIONS PAGE
# ============================================================

elif st.session_state.page == "instructions":

    st.markdown(
        '<div class="section-title">'
        'Instructions'
        '</div>',
        unsafe_allow_html=True
    )


    st.markdown(
        """
        Please read the following instructions carefully
        before starting the experiment.

        ### For each question

        **1. Listen carefully to the English audio.**

        Listen to the complete sentence before making your
        judgment. You may replay the audio if necessary.

        **2. Pay attention to the speaker's prosody.**

        The study focuses on two types of prosodic information:

        - **Emphasis:** a word or phrase may sound more
          prominent than the surrounding words.

        - **Rising contour:** the pitch may rise toward the
          end of a word, phrase, or sentence.

        **3. Pay attention to the indicated part of the sentence.**

        Where applicable, the relevant word or words will be
        indicated in the English sentence.

        **4. Rate the prosodic feature you perceive.**

        Based on what you hear in the audio, rate how strongly
        you perceive the relevant prosodic feature.

        **5. Choose the Hindi translation.**

        After giving your rating, select the Hindi translation
        that you think best matches the intended meaning of
        the English sentence, taking the prosodic information
        into account.

        ### Important points

        - Listen carefully before answering.
        - You may replay the audio if needed.
        - Focus on what you hear in the English speech.
        - Consider both the meaning of the sentence and its
          prosodic characteristics.
        - Do not judge the speaker based on their voice,
          gender, accent, or loudness.
        - Choose the translation based on its intended meaning
          and how well it reflects the prosodic information.
        - There are no right or wrong answers from the
          participant's perspective. We are interested in
          your perception and interpretation.
        """
    )


    st.markdown("---")


    st.markdown(
        """
        ### Prosodic Feature Rating Scale

        **1 — Not perceived at all**

        **2 — Slightly perceived**

        **3 — Moderately perceived**

        **4 — Strongly perceived**

        **5 — Very strongly perceived**
        """
    )


    st.markdown("---")


    if st.button(
        "Begin Experiment",
        use_container_width=True
    ):

        st.session_state.page = (
            "experiment"
        )

        st.session_state.question_start_times[
            st.session_state.current_question
        ] = datetime.now()

        st.rerun()


# ============================================================
# EXPERIMENT PAGE
# ============================================================

elif st.session_state.page == "experiment":

    question_index = (
        st.session_state.current_question
    )

    total_questions = len(
        questions
    )


    # ========================================================
    # SAFETY CHECK
    # ========================================================

    if question_index >= total_questions:

        st.session_state.page = (
            "remarks"
        )

        st.rerun()


    row = questions.iloc[
        question_index
    ]

    sample_id = str(
        row["sample id"]
    ).strip()

    question_number = (
        question_index + 1
    )


    # ========================================================
    # PROGRESS
    # ========================================================

    st.progress(
        question_number
        /
        total_questions
    )

    st.markdown(
        f'<div class="progress-text">'
        f'Question {question_number} of '
        f'{total_questions}'
        f'</div>',
        unsafe_allow_html=True
    )


    # ========================================================
    # ENGLISH SENTENCE
    # ========================================================

    st.markdown(
        '<div class="section-title">'
        'Listen to the sentence'
        '</div>',
        unsafe_allow_html=True
    )

    english_sentence = (
        row["english sentence"]
    )

    emphasized_word = (
        row["emphasized word"]
    )

    highlighted_sentence = (
        highlight_emphasis(
            english_sentence,
            emphasized_word
        )
    )

    st.markdown(
        f'<div class="sentence-box">'
        f'{highlighted_sentence}'
        f'</div>',
        unsafe_allow_html=True
    )


    # ========================================================
    # FEATURE INFORMATION
    # ========================================================

    prosodic_feature = get_prosodic_feature(
        row
    )

    if prosodic_feature:

        st.markdown(
            f"**Prosodic feature:** "
            f"<span class='emphasis-word'>"
            f"{html.escape(prosodic_feature)}"
            f"</span>",
            unsafe_allow_html=True
        )


    # ========================================================
    # EMPHASIZED WORD
    # ========================================================

    if str(
        emphasized_word
    ).strip():

        st.markdown(
            f"**Indicated word(s):** "
            f"<span class='emphasis-word'>"
            f"{html.escape(str(emphasized_word))}"
            f"</span>",
            unsafe_allow_html=True
        )


    # ========================================================
    # AUDIO
    # ========================================================

    audio_filename = str(
        row["audiofile"]
    ).strip()

    audio_path = os.path.join(
        AUDIO_FOLDER,
        audio_filename
    )

    if os.path.exists(
        audio_path
    ):

        st.audio(
            audio_path,
            format="audio/mp4"
        )

    else:

        st.error(
            f"Audio file not found: "
            f"{audio_filename}"
        )


    # ========================================================
    # EXISTING ANSWER
    # ========================================================

    existing_answer = (
        st.session_state.answers.get(
            sample_id,
            {}
        )
    )

    previous_selection = (
        existing_answer.get(
            "selected_translation",
            None
        )
    )

    previous_rating = (
        existing_answer.get(
            "prosodic_rating",
            None
        )
    )


    # ========================================================
    # START QUESTION TIMER
    # ========================================================

    if question_index not in (
        st.session_state.question_start_times
    ):

        st.session_state.question_start_times[
            question_index
        ] = datetime.now()


    # ========================================================
    # PROSODIC RATING — FIRST
    # ========================================================

    st.markdown("---")

    st.markdown(
        "### How strongly did you perceive the relevant "
        "prosodic feature in the audio?"
    )

    st.write(
        "Please base your rating on what you hear in the "
        "audio recording."
    )


    rating_labels = {

        1: "Not perceived at all",

        2: "Slightly perceived",

        3: "Moderately perceived",

        4: "Strongly perceived",

        5: "Very strongly perceived"
    }

    rating_options = [
        1,
        2,
        3,
        4,
        5
    ]


    # Convert saved value safely

    try:

        if previous_rating != "":

            previous_rating = int(
                float(previous_rating)
            )

        else:

            previous_rating = None

    except Exception:

        previous_rating = None


    emphasis_rating = st.radio(
        "Prosodic feature rating",
        rating_options,
        index=(
            rating_options.index(
                previous_rating
            )
            if previous_rating
            in rating_options
            else None
        ),
        format_func=lambda x:
            f"{x} — {rating_labels[x]}",
        key=f"rating_{sample_id}"
    )


    # ========================================================
    # TRANSLATION CHOICE — SECOND
    # ========================================================

    st.markdown("---")

    st.markdown(
        "### Choose the Hindi translation"
    )

    st.markdown(
        '<div class="translation-note">'
        'Select the translation that best matches the '
        'intended meaning and prosodic interpretation.'
        '</div>',
        unsafe_allow_html=True
    )


    # ========================================================
    # EXACTLY TWO TRANSLATION OPTIONS
    # ========================================================

    options = [

        (
            "Hindi Translation",
            str(
                row["hindi translation"]
            ).strip()
        ),

        (
            "Machine Translation",
            str(
                row["machine translation"]
            ).strip()
        )
    ]


    valid_options = [

        option

        for option in options

        if option[1]
    ]


    if len(valid_options) != 2:

        st.error(
            "This question must contain exactly two "
            "valid translation options."
        )

        st.stop()


    # ========================================================
    # STABLE RANDOMIZATION
    # ========================================================

    if sample_id not in (
        st.session_state.randomized_options
    ):

        shuffled_options = (
            get_randomized_options(
                st.session_state.participant_name,
                sample_id,
                valid_options
            )
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


    display_texts = [

        option[1]

        for option in shuffled_options
    ]


    # ========================================================
    # FIND PREVIOUS TRANSLATION SELECTION
    # ========================================================

    if sample_id not in (
        st.session_state.translation_selections
    ):

        st.session_state.translation_selections[
            sample_id
        ] = previous_selection


    previous_translation = (
        st.session_state.translation_selections[
            sample_id
        ]
    )


    # ========================================================
    # TRANSLATION CARDS
    # ========================================================

    selected_translation = st.radio(
        "Translation options",
        display_texts,
        index=(
            display_texts.index(
                previous_translation
            )
            if previous_translation
            in display_texts
            else None
        ),
        key=f"translation_choice_{sample_id}",
        label_visibility="collapsed"
    )


    # ========================================================
    # STORE CURRENT TRANSLATION SELECTION
    # ========================================================

    if selected_translation is not None:

        st.session_state.translation_selections[
            sample_id
        ] = selected_translation


    # ========================================================
    # NAVIGATION
    # ========================================================

    st.markdown("---")

    col1, col2 = st.columns(2)


    # ========================================================
    # PREVIOUS
    # ========================================================

    with col1:

        if question_index > 0:

            if st.button(
                "← Previous",
                use_container_width=True
            ):

                st.session_state.current_question -= 1

                st.session_state.question_start_times[
                    st.session_state.current_question
                ] = datetime.now()

                st.rerun()


    # ========================================================
    # NEXT / SUBMIT
    # ========================================================

    with col2:

        if (
            question_index
            ==
            total_questions - 1
        ):

            button_text = (
                "Submit Study"
            )

        else:

            button_text = (
                "Next →"
            )


        if st.button(
            button_text,
            use_container_width=True
        ):


            # =================================================
            # VALIDATION
            # =================================================

            if emphasis_rating is None:

                st.warning(
                    "Please provide a prosodic feature rating."
                )

                st.stop()


            if selected_translation is None:

                st.warning(
                    "Please select one of the two Hindi translations."
                )

                st.stop()


            # =================================================
            # RESPONSE TIME
            # =================================================

            start_time = (
                st.session_state.question_start_times.get(
                    question_index,
                    datetime.now()
                )
            )

            response_time = (
                datetime.now()
                -
                start_time
            ).total_seconds()


            # =================================================
            # IDENTIFY TRANSLATION SOURCE
            # =================================================

            selected_source = ""

            for source, text in shuffled_options:

                if (
                    text
                    ==
                    selected_translation
                ):

                    selected_source = source

                    break


            # =================================================
            # STORE ANSWER
            # =================================================

            st.session_state.answers[
                sample_id
            ] = {

                "selected_translation":
                    selected_translation,

                "selected_translation_type":
                    selected_source,

                "prosodic_rating":
                    emphasis_rating,

                "response_time_seconds":
                    round(
                        response_time,
                        2
                    )
            }


            # =================================================
            # SAVE RESPONSE
            # =================================================

            save_progress(
                sample_id
            )


            # =================================================
            # NEXT QUESTION
            # =================================================

            if (
                question_index
                ==
                total_questions - 1
            ):

                st.session_state.current_question = (
                    total_questions
                )

                st.session_state.page = (
                    "remarks"
                )

            else:

                st.session_state.current_question += 1

                st.session_state.question_start_times[
                    st.session_state.current_question
                ] = datetime.now()


            st.rerun()


# ============================================================
# REMARKS PAGE
# ============================================================

elif st.session_state.page == "remarks":

    st.markdown(
        '<div class="main-title">'
        'Thank You for Completing the Study'
        '</div>',
        unsafe_allow_html=True
    )


    st.markdown(
        """
        We would appreciate any comments or remarks you
        may have about your experience with the study.

        You may comment on the clarity of the questions,
        the audio, the translations, the perception of
        emphasis or rising contour, the difficulty of the
        task, or anything else you noticed during the
        experiment.
        """
    )


    st.markdown("---")


    st.markdown(
        "### Remarks"
    )


    remarks = st.text_area(
        "Please enter your comments or remarks",
        value=st.session_state.remarks,
        placeholder="Enter your remarks here...",
        height=180
    )


    st.session_state.remarks = (
        remarks
    )


    st.markdown("---")


    if st.button(
        "Submit Remarks",
        use_container_width=True
    ):

        success = save_remarks()

        if success:

            st.session_state.page = (
                "completed"
            )

            st.rerun()


# ============================================================
# COMPLETED PAGE
# ============================================================

elif st.session_state.page == "completed":

    st.markdown(
        '<div class="main-title">'
        'Study Completed'
        '</div>',
        unsafe_allow_html=True
    )


    st.success(
        "Thank you for participating in the study!"
    )


    st.markdown(
        """
        Your responses and remarks have been recorded
        successfully.

        You may now close this page.
        """
    )