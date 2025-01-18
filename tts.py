import pyttsx3

# Initialize the TTS engine
engine = pyttsx3.init()

# Load your text file
file_path = 'c:\somewhere\file.txt'  # Replace with the path to your text file
with open(file_path, 'r', encoding='utf-8') as file:
    text = file.read()

# Set voice properties (optional)
engine.setProperty('rate', 150)  # Speed (default is usually 200)
engine.setProperty('volume', 1.0)  # Volume (max is 1.0)

# Choose a voice (optional)
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[0].id)  # Use voices[1].id for a different voice (e.g., female)

# Convert text to speech
engine.save_to_file(text, 'output_audio.mp3')  # Save as an audio file
engine.runAndWait()  # Run the TTS engine
