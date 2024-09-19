import os
from pptx import Presentation
from pydub import AudioSegment
from tkinter import Tk
from tkinter.filedialog import askopenfilename, asksaveasfilename
import tempfile

def extract_audio_from_pptx(pptx_path):
    # Load the presentation
    prs = Presentation(pptx_path)
    
    # A list to store the audio segments
    audio_segments = []

    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_media and shape.media:
                # Save the embedded media to a temp file
                with tempfile.NamedTemporaryFile(delete=False, suffix='.wav') as temp_audio:
                    temp_audio.write(shape.media.blob)
                    temp_audio_path = temp_audio.name
                    
                # Load the audio segment
                audio = AudioSegment.from_file(temp_audio_path)
                audio_segments.append(audio)
                
                # Remove the temporary file after loading the audio
                os.remove(temp_audio_path)
    
    # Concatenate all audio segments into one long segment
    if audio_segments:
        final_audio = sum(audio_segments)
        return final_audio
    else:
        return None

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
    final_audio = extract_audio_from_pptx(pptx_path)

    if final_audio:
        # Ask where to save the resulting audio
        save_path = asksaveasfilename(
            title="Save the combined audio",
            defaultextension=".wav",
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
