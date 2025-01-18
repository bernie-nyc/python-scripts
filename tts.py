import pyttsx3

# Initialize the TTS engine
engine = pyttsx3.init()

# Load your text file
file_path = r'C:\some\path\file\filename.txt'  
# Adjust the path as needed
with open(file_path, 'r', encoding='utf-8') as file:
    text = file.read()

# Set voice properties (optional)
engine.setProperty('rate', 150)  # Speed
engine.setProperty('volume', 1.0)  # Volume

# Convert text to speech
engine.save_to_file(text, 'output_audio.mp3')  # Save as an audio file
engine.runAndWait()  # Run the TTS engine
