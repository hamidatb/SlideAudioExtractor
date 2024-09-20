import os
import zipfile
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox
import speech_recognition as sr
from pydub import AudioSegment
import sys

# Function to unzip pptx, delete unnecessary files, convert .m4a to .wav, and transcribe
def process_pptx():
    # Ask the user to select a .pptx file
    pptx_file = filedialog.askopenfilename(title="Select a PPTX file", filetypes=[("PPTX files", "*.pptx")])
    if not pptx_file:
        messagebox.showerror("Error", "No file selected!")
        return

    # Ask the user where to store the extracted .m4a files and their transcriptions
    save_folder = filedialog.askdirectory(title="Select Folder to Save Files")
    if not save_folder:
        messagebox.showerror("Error", "No save folder selected!")
        return

    print("Extracting PPTX...")
    # Create a folder to store the extracted contents in the chosen save folder
    output_folder = os.path.join(save_folder, "extracted_media")
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # Unzip the pptx file (treating it as a zip)
    with zipfile.ZipFile(pptx_file, 'r') as zip_ref:
        zip_ref.extractall(output_folder)

    # Define path to the media folder inside the extracted pptx contents
    media_folder = os.path.join(output_folder, "ppt", "media")

    # Check if the media folder exists
    if not os.path.exists(media_folder):
        messagebox.showerror("Error", "No media files found in this presentation!")
        return

    print("Processing .m4a files and transcribing...")
    # Initialize the recognizer for transcription
    recognizer = sr.Recognizer()
    print(os.listdir(media_folder))

    # Loop over each file in the media folder
    for file in os.listdir(media_folder):
        if file.endswith(".m4a"):
            # Path to the current .m4a file
            m4a_path = os.path.join(media_folder, file)
            wav_path = os.path.splitext(m4a_path)[0] + ".wav"  # Change the extension to .wav

            print(f"Converting {file} to .wav...")

            # Convert .m4a to .wav using pydub
            try:
                AudioSegment.from_file(m4a_path).export(wav_path, format="wav")
                print(f"Converted {file} to {wav_path}")
            except Exception as e:
                print(f"Error converting {file}: {e}")
                continue

            # Transcribe the .wav file using Google Speech Recognition
            print(f"Transcribing {file}...")
            try:
                with sr.AudioFile(wav_path) as source:
                    audio_data = recognizer.record(source)
                    transcription = recognizer.recognize_google(audio_data)

                # Save the transcription to a text file with the same name as the .m4a file
                transcript_path = os.path.join(save_folder, os.path.splitext(file)[0] + ".txt")
                with open(transcript_path, "w") as f:
                    f.write(transcription)
                print(f"Saved transcript for {file} as {transcript_path}")

            except sr.RequestError as e:
                print(f"Could not request results from Google Speech Recognition for {file}; {e}")
            except sr.UnknownValueError:
                print(f"Google Speech Recognition could not understand the audio from {file}")

    # Optionally, delete everything but the .m4a and .wav files and their transcripts
    print("Cleaning up unnecessary files...")
    for root, dirs, files in os.walk(output_folder):
        for name in files:
            if not name.endswith(('.m4a', '.wav', '.txt')):
                os.remove(os.path.join(root, name))
        for name in dirs:
            if name != "media":
                shutil.rmtree(os.path.join(root, name))

    messagebox.showinfo("Success", f"Processed all audio files and saved transcriptions in: {save_folder}")
    print("Process completed successfully!")

    sys.exit()

# Create the Tkinter window
root = tk.Tk()
root.title("PPTX Audio Transcription")
root.geometry("300x150")

# Create a button to trigger the PPTX processing
process_button = tk.Button(root, text="Select PPTX and Process", command=process_pptx)
process_button.pack(pady=30)

# Start the Tkinter event loop
root.mainloop()
