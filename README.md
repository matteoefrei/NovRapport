# NovRapport Web

A simple Flask application to generate DOCX reports online. Users can add text or images, preview the result and download the final document. Images are stored in memory while reports are saved under `novrapport_web/docs`.

## Installation

1. Create a virtual environment and activate it.
2. Install dependencies:
   ```bash
   pip install -r novrapport_web/requirements.txt
   ```

## Usage

Run the application:
```bash
python novrapport_web/app.py
```
Open `http://localhost:5000` in your browser.

## Notes

Uploaded images and generated documents are not committed to the repository thanks to the `.gitignore` rules.
