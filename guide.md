# ODT Question Processor - Complete Documentation

## 📋 Overview (Tafseel)

Yeh Flask-based web application ODT (OpenDocument Text) files ko process karti hai aur questions ko properly format karti hai. Yeh application **sabse pehle validate** karti hai ke uploaded file actual ODT file hai ya nahi, aur agar koi issues hain to unko list kar deti hai.

## 🎯 Key Features (Khaas Khasusiyat)

### 1. **ODT File Validation** (Sabse Pehle)
Application upload hone wali file ko thoroughly check karti hai:

- ✅ File exist karti hai ya nahi
- ✅ File empty to nahi (0 bytes)
- ✅ Valid ZIP format hai (ODT files ZIP format mein hoti hain)
- ✅ Correct magic bytes hain (PK signature)
- ✅ `mimetype` file mojood hai
- ✅ Correct mimetype hai (`application/vnd.oasis.opendocument.text`)
- ✅ `content.xml` file mojood hai
- ✅ `content.xml` parse ho sakta hai

**Agar koi bhi check fail hota hai**, application user ko clear error message dikhati hai ke kya issue hai.

### 2. **Processing Operations** (7 Steps)

File valid hone ke baad, yeh operations perform hote hain:

#### Step 1: Fix Question Numbers & Remove Brackets
- Question numbers ko ascending order mein fix karta hai (1, 2, 3...)
- Duplicate question numbers ko remove karta hai
- Bracket text `[People]`, `[Process]` etc. ko remove karta hai
- Valid question types (HOTSPOT, SIMULATION, DRAG DROP, etc.) ko preserve karta hai

**Input:**
```
Question No: 45 HOTSPOT [People]
Question No: 45 SIMULATION
Question: 12 [Process]
```

**Output:**
```
Question: 1 HOTSPOT
Question: 2 SIMULATION
Question: 3
```

#### Step 2: Ensure Spacing Before Questions
- Questions se pehle proper spacing ensure karta hai
- Agar question directly text ke baad hai to spacing add karta hai

#### Step 3: Combined Text Operations
Ek hi pass mein multiple text operations:
- `QUESTION NO:` → `Question:` (replace karta hai)
- `Explanation:` tags remove karta hai (unnecessary ones)
- `References:` → `Reference:` (plural to singular)
- `<map>` tags remove karta hai

**Input:**
```
QUESTION NO: 1
Explanation: Some text
References: Link here
<map>data</map>
```

**Output:**
```
Question: 1
Some text
Reference: Link here

```

#### Step 4: Question Types to Next Line
Question types ko separate line par shift karta hai

**Input:**
```
Question: 1 HOTSPOT Some text here
```

**Output:**
```
Question: 1
HOTSPOT Some text here
```

#### Step 5: Normalize Option Spacing
Options (A, B, C, D) ka formatting consistent banata hai

**Input:**
```
A.Option text
B .  Text here
C    .Text
```

**Output:**
```
A. Option text
B. Text here
C. Text
```

#### Step 6: Add Explanation Tags
Answer ke baad agar text hai aur "Explanation:" nahi hai to add karta hai

**Input:**
```
Answer: A
This is the explanation text.
```

**Output:**
```
Answer: A
Explanation:
This is the explanation text.
```

#### Step 7: Add Line Spacing
Question: aur Answer: ke baad proper line spacing add karta hai for better readability

## 🚀 Installation & Usage

### Prerequisites
```bash
Python 3.8+
pip (Python package manager)
```

### Installation Steps

1. **Dependencies install karein:**
```bash
pip install -r requirements.txt
```

2. **Application run karein:**
```bash
python app_odt.py
```

3. **Browser mein open karein:**
```
http://localhost:5000
```

## 📁 File Structure

```
project/
│
├── app_odt.py              # Main Flask application
├── requirements.txt         # Python dependencies
├── templates/
│   └── index.html          # Web interface
└── uploads/                # Processed files save hoti hain (auto-created)
```

## 🔍 How Validation Works (Kaise Kaam Karta Hai)

### Valid ODT File ki Requirements:

1. **ZIP Structure**: ODT file ek ZIP archive hoti hai
```python
# Check karta hai zipfile.is_zipfile()
```

2. **Magic Bytes**: File ke first 2 bytes `PK` (0x504B) hone chahiye
```python
# Reads first 4 bytes: b'PK\x03\x04'
```

3. **Mimetype File**: Archive mein `mimetype` file honi chahiye with exact content:
```
application/vnd.oasis.opendocument.text
```

4. **Content.xml**: Main content file honi chahiye jo valid XML ho

### Invalid File Examples:

#### Example 1: Simple XML File (Not ODT)
```xml
<?xml version="1.0"?>
<document>Content here</document>
```
**Error**: "File is not a valid ZIP archive"

#### Example 2: Wrong Mimetype
Agar `mimetype` file mein:
```
application/xml
```
**Error**: "Invalid mimetype: 'application/xml' (expected: 'application/vnd.oasis.opendocument.text')"

#### Example 3: Missing Files
Agar ZIP hai lekin `content.xml` missing hai:
**Error**: "Missing 'content.xml' file in ODT structure"

## 🖥️ User Interface Features

### Real-time Status Updates
- Har processing step ka live status
- Time taken for each operation
- Visual indicators (⏳ in progress, ✅ completed, ❌ failed)

### Error Handling
- Clear validation errors with issue list
- Processing errors with detailed messages
- User-friendly error display

### Download
- Automatic processed file download
- Original filename preserved with "processed_" prefix
- Statistics display (total time, questions fixed)

## 🔧 Technical Details

### ODT File Structure:
```
ODT File (ZIP Archive)
│
├── mimetype                          # MIME type identifier
├── content.xml                       # Main document content
├── styles.xml                        # Document styles
├── meta.xml                          # Metadata
└── META-INF/
    └── manifest.xml                  # File manifest
```

### XML Namespaces Used:
```python
ODT_NAMESPACES = {
    'text': 'urn:oasis:names:tc:opendocument:xmlns:text:1.0',
    'office': 'urn:oasis:names:tc:opendocument:xmlns:office:1.0'
}
```

### Processing Flow:
```
Upload File
    ↓
Validate ODT Structure
    ↓
Extract to Temp Directory
    ↓
Parse content.xml
    ↓
Apply 7 Processing Steps
    ↓
Save Modified XML
    ↓
Re-pack to ODT (ZIP)
    ↓
Return Processed File
```

## ⚠️ Important Notes

1. **File Size Limit**: Maximum 200MB
2. **Only ODT**: Application sirf ODT files accept karti hai
3. **Temporary Files**: Processing ke baad automatically cleanup hoti hai
4. **Thread-Safe**: Background processing with proper thread handling

## 🐛 Troubleshooting

### Issue: "File is not a valid ZIP archive"
**Solution**: File actual mein ODT nahi hai. Proper ODT file export/save karein.

### Issue: "Invalid mimetype"
**Solution**: File ko ODT format mein properly save karein (LibreOffice ya OpenOffice use karein).

### Issue: "Cannot parse content.xml"
**Solution**: File corrupt hai. Original file se dobara try karein.

## 📞 Support

Agar koi issue ho to check karein:
1. Python version 3.8+ hai
2. Sahi dependencies install hain
3. ODT file properly saved hai
4. File corrupt nahi hai

## 🎓 Learning Points

Yeh code demonstrate karta hai:
- ODT file format understanding
- XML parsing with lxml
- Flask SSE (Server-Sent Events) for real-time updates
- File validation best practices
- Background processing in Flask
- Error handling and user feedback

---

**Note**: Yeh application production-ready hai with proper error handling, validation, and user-friendly interface.