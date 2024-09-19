# SlideAudioExtractor

SlideAudioExtractor is a Python script designed to help students extract and combine audio embedded in PowerPoint slides into a single `.wav` file. This tool was created for students who, like me, have professors who upload lecture slides with embedded audio. It allows you to easily collect all the slide audio in one file, so you can listen to lectures more efficiently without flipping through individual slides.

## Why I Made This

I built this tool because my professor uploads slides with embedded audio, and it was frustrating to listen to each slide one by one. With this tool, you can extract all the audio from the slides and create a single audio file for easier review.

## Features

- Extracts all embedded audio from PowerPoint presentations.
- Combines audio clips from each slide into one long `.wav` file.
- Simple, user-friendly interface using `tkinter`.

## Requirements

- **Python 3.6** or later
- **python-pptx**: for reading PowerPoint files.
- **pydub**: for handling and combining audio files.
- **tkinter**: for providing a graphical interface.
- **ffmpeg**: for handling audio processing (required by `pydub`).

### Installing ffmpeg (on macOS)

If you don't have `ffmpeg` installed, you can easily install it using Homebrew:

```bash
brew install ffmpeg
