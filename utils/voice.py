from streamlit_mic_recorder import speech_to_text


def listen():

    try:

        text = speech_to_text(
            start_prompt="🎤 Start Recording",
            stop_prompt="⏹ Stop Recording",
            language="en-US",
            just_once=True,
            use_container_width=True,
            key="voice"
        )

        if text is None:
            return ""

        text = str(text).strip()

        if text == "":
            return ""

        return text

    except Exception as e:

        print("Voice Error:", e)
        return ""