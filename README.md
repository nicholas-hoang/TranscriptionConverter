# Meeting Transcription Cleaner

The Meeting Transcription Cleaner is a Python application that takes a DOCX file containing meeting transcriptions, processes the data, and provides a cleaned transcript. It uses the Gradio library for the user interface and python-docx for handling DOCX files.

## Usage

**Upload File:** Click on the "Upload your DOCX file" button to upload your meeting transcription file in DOCX format.

**Processing:** The application processes the uploaded file, cleans the transcript data, and generates a new DOCX file with the cleaned transcription.

**Download:** Once the processing is complete, click on the "Download cleaned transcript" button to download the cleaned transcript DOCX file.

## How It Works

The application processes the meeting transcription in the following steps:

- Upload: Users upload a DOCX file containing meeting transcriptions.
- Processing: The application processes the DOCX file, extracting text data.
- Cleaning: The text data is cleaned and formatted, including splitting the text into speakers, timestamps, and actual text content.
- Output: The cleaned transcription data is saved into a new DOCX file.
- Download: Users can download the cleaned transcription file.

## Requirements
Python 3.x
Gradio
python-docx
pandas

## Installation

Clone the repository:

```bash
git clone <https://github.com/nvh232/meeting-transcription-cleaner.git>
```
Install the required packages:
```bash
pip install gradio python-docx pandas
```
Run the application:
```bash
python app.py
```
## Contributors
Nick Hoang

## License

This project is licensed under the MIT License.
