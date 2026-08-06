import pyttsx3

engine = pyttsx3.init()

engine.setProperty("rate", 170)
engine.setProperty("volume", 1)

voices = engine.getProperty("voices")

for v in voices:
    print(v.id)

engine.say("Hello Saad. Voice is working.")
engine.runAndWait()