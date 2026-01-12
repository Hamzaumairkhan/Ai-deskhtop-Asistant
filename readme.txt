run this in terminal for audio control 


py -3.12 -m pip install --upgrade pip
py -3.12 -m pip install pyaudio==0.2.14 --only-binary :all:
py -3.12 -m pip install numpy requests SpeechRecognition pyttsx3
py -3.12 main.py