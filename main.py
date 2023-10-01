import  win32com.client as wincom

speak = wincom.Dispatch('SAPI.SpVoice')
while True:
    text = input("Enter your text to speak:")
    if text == 'q':
        speak.Speak("Bye Bye")
        break
    speak.Speak(text)
