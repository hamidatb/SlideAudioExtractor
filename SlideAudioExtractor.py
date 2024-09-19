import os
import tempfile
import zipfile
from pydub import AudioSegment
from tkinter import Tk
from tkinter.filedialog import askopenfilename, asksaveasfilename
from gtts import gTTS

def extract_audio_from_pptx_as_zip(pptx_path):
    # A list to store the audio segments
    audio_segments = []

    # Unzip the .pptx file and extract media files
    with zipfile.ZipFile(pptx_path, 'r') as pptx_zip:
        # Extract media files that end with .m4a and sort by the numerical part of their name
        media_files = sorted(
            [f for f in pptx_zip.namelist() if f.startswith('ppt/media/') and f.endswith('.m4a')],
            key=lambda x: int(''.join(filter(str.isdigit, x)))
        )

        for i, media_file in enumerate(media_files, start=1):
            # Add 'Slide X' audio between each slide
            slide_number_audio = generate_slide_number_audio(i)
            audio_segments.append(slide_number_audio)

            # Extract the audio file from the zip
            with pptx_zip.open(media_file) as media:
                # Convert the file to an AudioSegment using pydub (with ffmpeg support)
                audio_segment = AudioSegment.from_file(media, format="m4a")
                audio_segments.append(audio_segment)

    # Concatenate all audio segments into one long segment
    if audio_segments:
        final_audio = sum(audio_segments)
        return final_audio
    else:
        return None

def generate_slide_number_audio(slide_number):
    """Generate 'Slide X' audio using gTTS"""
    text = f"Slide {slide_number}"
    tts = gTTS(text)
    
    # Create a temporary audio file for the slide number
    with tempfile.NamedTemporaryFile(delete=False, suffix='.mp3') as temp_audio:
        temp_audio_path = temp_audio.name
        tts.save(temp_audio_path)
    
    # Load the generated audio file as an AudioSegment
    slide_audio = AudioSegment.from_file(temp_audio_path)
    
    # Remove the temporary file
    os.remove(temp_audio_path)
    
    return slide_audio

def main():
    # Tkinter setup
    root = Tk()
    root.withdraw()  # Hide the root window

    # Ask the user to select the PPTX file
    pptx_path = askopenfilename(
        title="Select the PowerPoint file",
        filetypes=[("PowerPoint files", "*.pptx")]
    )

    if not pptx_path:
        print("No file selected, exiting.")
        return

    # Extract the audio from the selected PPTX file
    final_audio = extract_audio_from_pptx_as_zip(pptx_path)

    if final_audio:
        # Pre-fill the filename with 'AUDIO_<original_filename>.wav'
        original_filename = os.path.splitext(os.path.basename(pptx_path))[0]
        default_filename = f"AUDIO_{original_filename}.wav"

        # Ask where to save the resulting audio
        save_path = asksaveasfilename(
            title="Save the combined audio",
            defaultextension=".wav",
            initialfile=default_filename,
            filetypes=[("WAV files", "*.wav")]
        )

        if save_path:
            final_audio.export(save_path, format="wav")
            print(f"Audio saved to {save_path}")
        else:
            print("No save location provided.")
    else:
        print("No audio found in the PowerPoint file.")

if __name__ == "__main__":
    main()
