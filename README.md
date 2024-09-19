# SlideAudioExtractor

SlideAudioExtractor is a Python script designed to extract and combine embedded audio from PowerPoint presentations (`.pptx`) into a single `.wav` file. This tool is especially useful for students and professionals who need to review or organize audio from lecture slides or presentations where the audio is embedded as clickable elements on each slide.

## Why I Made This

This script was created because many professors upload slides with embedded audio that you have to click to play. Navigating through each slide to listen to the audio can be cumbersome. SlideAudioExtractor solves this problem by extracting all embedded audio clips and combining them into a single audio file, while also announcing the slide number before each clip.

## Features

- **Extracts Embedded Audio**: Extracts and combines audio files embedded as clickable elements on PowerPoint slides into a single `.wav` file.
- **Supports `.m4a` Audio**: Specifically designed to handle `.m4a` audio files commonly used in PowerPoint presentations for embedded audio.
- **Slide Announcements**: Automatically inserts "Slide X" announcements before each audio file using Google Text-to-Speech (`gTTS`).
- **Chronological Order**: Ensures that the audio is extracted and combined in the correct order, matching the slide sequence.

## Audio Formats

This script is built to handle the following types of audio files that are embedded as clickable elements on PowerPoint slides:

- **.m4a**: PowerPoint often embeds audio in `.m4a` format, especially for the clickable audio icons on the slide. The script extracts these files and ensures they are combined in the correct order.
- **.mp3** and **.wav** (if present): Although `.m4a` is the primary format, if `.mp3` or `.wav` files are embedded as clickable audio on the slides, they will also be extracted and included in the output.

## Requirements

- **Python 3.6** or later
- **pydub**: For handling audio files and combining them.
- **gTTS**: For generating "Slide X" announcements.
- **tkinter**: For providing the graphical interface to select files and save locations.
- **ffmpeg**: Required by `pydub` to handle `.m4a` audio files.

### Installing ffmpeg (on macOS)

To install `ffmpeg` on macOS, you can use Homebrew:

```bash
brew install ffmpeg
```

If you are using another operating system, check [ffmpeg's official installation guide](https://ffmpeg.org/download.html) for details.

## Installation

To install the required Python libraries, run:

```bash
pip install -r requirements.txt
```

Make sure you also have `ffmpeg` installed to handle `.m4a` files.

## Usage

1. Run the script:

   ```bash
   python SlideAudioExtractor.py
   ```

2. Select the PowerPoint (`.pptx`) file that contains the embedded audio.

   - **Note**: The script is designed to handle audio files embedded as clickable elements on the slides. These are the icons you click to play the audio when viewing the slide in PowerPoint.

3. The script will extract the audio in the order of the slides, add "Slide X" announcements before each clip, and prompt you to save the combined audio as a `.wav` file.

### Default Save Filename

The default save filename will be pre-filled as `AUDIO_<original_pptx_filename>.wav`, where `<original_pptx_filename>` is the name of the PowerPoint file you selected.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for more information.

