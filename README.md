# Document Macro Operations - Flask Application

Ye application DOCX aur ODT files ko process karti hai aur multiple document macros apply karti hai.

## Project Structure

```
flask_document_processor/
├── app.py                 # Main Flask application
├── templates/
│   └── index.html        # Frontend HTML
├── requirements.txt      # Python dependencies
└── uploads/              # Auto-generated folder for processed files
```

## Installation

### 1. Python Environment Setup

```bash
# Virtual environment create karein (optional but recommended)
python -m venv venv

# Virtual environment activate karein
# Windows:
venv\Scripts\activate
# Linux/Mac:
source venv/bin/activate
```

### 2. Dependencies Install Karein

```bash
pip install -r requirements.txt
```

## How to Run

### Application Start Karein

```bash
python app.py
```

Application start ho jayegi aur ye message dikhega:
```
* Running on http://0.0.0.0:5000
```

### Browser Mein Open Karein

Apne browser mein ja kar ye URL open karein:
```
http://localhost:5000
```

## Usage

1. **File Upload**: DOCX ya ODT file upload karein (drag-drop ya click karke)
2. **Process**: "Apply Macros" button click karein
3. **Download**: Processing complete hone ke baad processed file download karein

## Features

### For ODT Files:
- Question numbering fix (1, 2, 3...)
- "Question No:" ko "Question:" mein convert
- Bracket text remove ([People], [Process] etc.)
- Valid question types preserve

### For DOCX Files:
- Question numbers fix aur brackets remove
- Proper spacing ensure
- Text operations (Question No: → Question:, References → Reference)
- Question types ko next line mein move
- Option spacing normalize (A. text, B. text)
- Explanation tags add
- Line spacing add after Question: aur Answer:

## API Endpoints

### GET `/`
- Main page render karti hai

### POST `/process`
- File upload aur process karti hai
- **Request**: multipart/form-data with 'file' field
- **Response**: JSON with processing status

### GET `/download/<filename>`
- Processed file download karti hai

## Configuration

`app.py` mein ye settings change kar sakte hain:

```python
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # Max file size (16MB)
app.config['UPLOAD_FOLDER'] = 'uploads'               # Upload folder path
```

## Troubleshooting

### Port already in use error
Agar port 5000 already use mein hai, to app.py ki last line change karein:
```python
app.run(debug=True, host='0.0.0.0', port=5001)  # Different port use karein
```

### Import errors
Make sure sare dependencies install hain:
```bash
pip install -r requirements.txt
```

## Requirements

- Python 3.7+
- Flask 3.0.0
- python-docx 1.1.0
- lxml 5.1.0
- Werkzeug 3.0.1

## Notes

- Streamlit code ko completely Flask mein convert kiya gaya hai
- Real-time processing status updates
- Modern, responsive UI with drag-and-drop support
- Error handling aur validation included
- Automatic cleanup of temporary files