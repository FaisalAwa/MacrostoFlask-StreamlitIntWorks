from flask import Flask, render_template, request, send_file, jsonify, Response
from io import BytesIO
import re
import time
import zipfile
from lxml import etree
import tempfile
import os
import shutil
from werkzeug.utils import secure_filename
import json
import queue
import threading

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 200 * 1024 * 1024  # 200MB max file size
app.config['UPLOAD_FOLDER'] = 'uploads'

# Create uploads folder if it doesn't exist
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
 
# Queue for real-time status updates
status_queue = queue.Queue()

# Pre-compile regex patterns for better performance
QUESTION_PATTERN = re.compile(r"^\s*Question\s*:?\s*\d*\s*$")
QUESTION_NUMBER_PATTERN = re.compile(r"Question\s*:\s*\d+")
QUESTION_TYPE_PATTERN = re.compile(r"(Question\s*:\s*\d+)\s+(HOTSPOT|SIMULATION|DRAG DROP)(.*)")
OPTION_PATTERN = re.compile(r"^\s*([A-J])\s*\.\s*(.+)$")
MAP_TAG_PATTERN = re.compile(r"<map>.*?</map>", flags=re.DOTALL)

# Valid question types
VALID_QUESTION_TYPES = [
    'DRAGDROP',
    'DRAG DROP',
    'DROPDOWN',
    'HOTSPOT',
    'FILLINTHEBLANK',
    'SIMULATION',
    'POSITIONEDDRAGDROP',
    'POSITIONEDDROPDOWN'
]

# ODT Namespaces
ODT_NAMESPACES = {
    'text': 'urn:oasis:names:tc:opendocument:xmlns:text:1.0',
    'office': 'urn:oasis:names:tc:opendocument:xmlns:office:1.0'
}


# ==================== ODT VALIDATION ====================

def validate_odt_file(file_path):
    """
    Validates if file is a proper ODT file
    Returns: (is_valid, issues_list)
    """
    issues = []
    
    # Check 1: File exists
    if not os.path.exists(file_path):
        issues.append("File does not exist")
        return False, issues
    
    # Check 2: File size > 0
    file_size = os.path.getsize(file_path)
    if file_size == 0:
        issues.append("File is empty (0 bytes)")
        return False, issues
    
    # Check 3: Is it a valid ZIP file?
    if not zipfile.is_zipfile(file_path):
        issues.append("File is not a valid ZIP archive (ODT files must be ZIP format)")
        return False, issues
    
    # Check 4: Read magic bytes
    try:
        with open(file_path, "rb") as f:
            magic = f.read(4)
            # ZIP files start with PK (0x504B)
            if magic[:2] != b'PK':
                issues.append(f"Invalid file signature: {magic.hex()} (expected: 504B for ZIP)")
    except Exception as e:
        issues.append(f"Cannot read file signature: {str(e)}")
        return False, issues
    
    # Check 5: Has mimetype file?
    try:
        with zipfile.ZipFile(file_path, 'r') as z:
            file_list = z.namelist()
            
            if 'mimetype' not in file_list:
                issues.append("Missing 'mimetype' file in ODT structure")
                return False, issues
            
            # Check 6: Correct mimetype content
            mimetype = z.read('mimetype').decode('utf-8', 'ignore').strip()
            expected_mimetype = 'application/vnd.oasis.opendocument.text'
            
            if mimetype != expected_mimetype:
                issues.append(f"Invalid mimetype: '{mimetype}' (expected: '{expected_mimetype}')")
                return False, issues
            
            # Check 7: Has content.xml?
            if 'content.xml' not in file_list:
                issues.append("Missing 'content.xml' file in ODT structure")
                return False, issues
            
            # Check 8: Can parse content.xml?
            try:
                content_xml = z.read('content.xml')
                etree.fromstring(content_xml)
            except Exception as e:
                issues.append(f"Cannot parse content.xml: {str(e)}")
                return False, issues
    
    except zipfile.BadZipFile:
        issues.append("Corrupted ZIP structure")
        return False, issues
    except Exception as e:
        issues.append(f"Error during validation: {str(e)}")
        return False, issues
    
    # All checks passed
    return True, []


# ==================== ODT HELPER FUNCTIONS ====================

def extract_valid_question_type(text):
    """
    Text se valid question type extract karta hai.
    Agar valid type nahi mila to None return karta hai.
    """
    text_upper = text.upper()
    
    for q_type in VALID_QUESTION_TYPES:
        if q_type in text_upper:
            return q_type
    
    return None


def get_para_text(para):
    """Extract text from ODT paragraph element"""
    return ''.join(para.itertext())


def clear_para_content(para):
    """Clear all child elements from paragraph"""
    for child in list(para):
        para.remove(child)
    para.text = None
    para.tail = None


def set_para_text(para, text):
    """Set text for ODT paragraph"""
    clear_para_content(para)
    para.text = text


# ==================== ODT PROCESSING FUNCTIONS ====================

def fix_odt_question_numbers_and_brackets(root, namespaces):
    """
    ODT file mein:
    1. Question numbers ko fix karta hai (ascending order: 1, 2, 3...)
    2. Bracket text [People], [Process] etc. ko remove karta hai
    3. Valid question types ko preserve karta hai
    """
    question_counter = 1
    
    # Pattern for both "Question No:" and "Question:"
    pattern = re.compile(r'^Question\s*(No)?:\s*\d+', re.IGNORECASE)
    
    paragraphs = root.xpath('//text:p', namespaces=namespaces)
    
    for para in paragraphs:
        para_text = get_para_text(para).strip()
        
        if pattern.match(para_text):
            # Extract question type agar hai to
            question_type = extract_valid_question_type(para_text)
            
            # Set new text with proper numbering (without brackets)
            if question_type:
                set_para_text(para, f"Question: {question_counter} {question_type}")
            else:
                set_para_text(para, f"Question: {question_counter}")
            
            question_counter += 1
    
    return root, question_counter - 1


def ensure_spacing_before_question_tags_odt(root, namespaces):
    """
    Ensures there's proper spacing before QUESTION NO: tags in ODT.
    """
    paragraphs = root.xpath('//text:p', namespaces=namespaces)
    
    for para in paragraphs:
        para_text = get_para_text(para)
        
        # Check if paragraph contains QUESTION NO: but doesn't start with it
        if "QUESTION NO:" in para_text and not para_text.strip().startswith("QUESTION NO:"):
            # Split the text at QUESTION NO:
            parts = para_text.split("QUESTION NO:")
            
            if len(parts) > 1:
                # For ODT, we'll keep it simple - just set the text properly
                new_text = parts[0].rstrip()
                if new_text:
                    new_text += "\n\n"
                new_text += "QUESTION NO:" + "QUESTION NO:".join(parts[1:])
                
                set_para_text(para, new_text)
    
    return root


def combined_text_operations_odt(root, namespaces):
    """
    Combined text operations for ODT:
    - Replace QUESTION NO: with Question:
    - Remove Explanation: tags
    - Replace References with Reference
    - Remove <map> tags
    """
    paragraphs = root.xpath('//text:p', namespaces=namespaces)
    
    for para in paragraphs:
        para_text = get_para_text(para)
        modified = False
        
        # Replace QUESTION NO: with Question:
        if "QUESTION NO:" in para_text:
            para_text = para_text.replace("QUESTION NO:", "Question:")
            modified = True
        
        # Remove Explanation: tags (case insensitive)
        if re.search(r'\bExplanation\s*:\s*', para_text, re.IGNORECASE):
            para_text = re.sub(r'\bExplanation\s*:\s*', '', para_text, flags=re.IGNORECASE)
            modified = True
        
        # Replace References with Reference (multiple to singular)
        if "References:" in para_text:
            para_text = para_text.replace("References:", "Reference:")
            modified = True
        
        # Remove <map> tags
        if "<map>" in para_text.lower():
            para_text = MAP_TAG_PATTERN.sub('', para_text)
            modified = True
        
        if modified:
            set_para_text(para, para_text)
    
    return root


# def shift_question_types_to_next_line_odt(root, namespaces):
#     """
#     Question types (HOTSPOT, SIMULATION, DRAG DROP) ko next line mein shift karta hai
#     """
#     paragraphs = root.xpath('//text:p', namespaces=namespaces)
    
#     for i, para in enumerate(paragraphs):
#         para_text = get_para_text(para)
        
#         # Check if this matches Question: X TYPE pattern
#         match = QUESTION_TYPE_PATTERN.match(para_text.strip())
        
#         if match:
#             question_part = match.group(1)  # "Question: X"
#             question_type = match.group(2)  # "HOTSPOT" etc.
#             remaining = match.group(3)      # Any remaining text
            
#             # Set current paragraph to just question number
#             set_para_text(para, question_part)
            
#             # Insert new paragraph after current with question type
#             # In ODT, we need to insert in parent
#             parent = para.getparent()
#             if parent is not None:
#                 index = list(parent).index(para)
                
#                 # Create new paragraph element
#                 new_para = etree.Element(f"{{{namespaces['text']}}}p")
#                 new_para.text = question_type + (remaining if remaining else "")
                
#                 # Insert after current paragraph
#                 parent.insert(index + 1, new_para)
    
#     return root


def shift_question_types_to_next_line_odt(root, namespaces):
    """
    Question types (HOTSPOT, SIMULATION, DRAG DROP) ko next line mein shift karta hai
    FIXED: Properly handles parent list modifications
    """
    paragraphs = root.xpath('//text:p', namespaces=namespaces)
    
    # Collect all changes first to avoid modifying list during iteration
    changes_to_make = []
    
    for para in paragraphs:
        para_text = get_para_text(para)
        
        # Check if this matches Question: X TYPE pattern
        match = QUESTION_TYPE_PATTERN.match(para_text.strip())
        
        if match:
            question_part = match.group(1)  # "Question: X"
            question_type = match.group(2)  # "HOTSPOT" etc.
            remaining = match.group(3)      # Any remaining text
            
            changes_to_make.append({
                'para': para,
                'question_part': question_part,
                'question_type': question_type,
                'remaining': remaining
            })
    
    # Apply all changes after collection
    for change in changes_to_make:
        para = change['para']
        question_part = change['question_part']
        question_type = change['question_type']
        remaining = change['remaining']
        
        # Set current paragraph to just question number
        set_para_text(para, question_part)
        
        # Insert new paragraph after current with question type
        parent = para.getparent()
        if parent is not None:
            try:
                # Get fresh list of children
                children = list(parent)
                if para in children:
                    index = children.index(para)
                    
                    # Create new paragraph element
                    new_para = etree.Element(f"{{{namespaces['text']}}}p")
                    new_para.text = question_type + (remaining if remaining else "")
                    
                    # Insert after current paragraph
                    parent.insert(index + 1, new_para)
            except ValueError:
                # Element not found in parent - skip
                continue
    
    return root


def normalize_option_spacing_odt(root, namespaces):
    """
    Options (A. B. C. etc.) ke formatting ko normalize karta hai
    Format: "A. " (capital letter, period, single space)
    """
    paragraphs = root.xpath('//text:p', namespaces=namespaces)
    
    for para in paragraphs:
        para_text = get_para_text(para)
        
        # Check if paragraph matches option pattern
        match = OPTION_PATTERN.match(para_text.strip())
        
        if match:
            option_letter = match.group(1)  # A, B, C, etc.
            option_text = match.group(2)    # Rest of the text
            
            # Normalize: "A. text"
            normalized_text = f"{option_letter}. {option_text}"
            set_para_text(para, normalized_text)
    
    return root


# def add_explanation_tags_if_text_present_odt(root, namespaces):
#     """
#     Answer ke baad agar text hai to "Explanation:" tag add karta hai
#     """
#     paragraphs = root.xpath('//text:p', namespaces=namespaces)
#     answer_found = False
#     explanation_added = False
    
#     for i, para in enumerate(paragraphs):
#         para_text = get_para_text(para).strip()
        
#         # Check if this is Answer: line
#         if para_text.startswith("Answer:"):
#             answer_found = True
#             explanation_added = False
#             continue
        
#         # Check if this is next Question:
#         if para_text.startswith("Question:"):
#             answer_found = False
#             explanation_added = False
#             continue
        
#         # If we're after Answer: and haven't added Explanation yet
#         if answer_found and not explanation_added and para_text:
#             # Check if line already has Explanation:
#             if not para_text.startswith("Explanation:") and not para_text.startswith("Reference:"):
#                 # Insert new paragraph with "Explanation:" before this one
#                 parent = para.getparent()
#                 if parent is not None:
#                     index = list(parent).index(para)
                    
#                     # Create new paragraph with Explanation:
#                     explanation_para = etree.Element(f"{{{namespaces['text']}}}p")
#                     explanation_para.text = "Explanation:"
                    
#                     # Insert before current paragraph
#                     parent.insert(index, explanation_para)
                    
#                     explanation_added = True
    
#     return root


def add_explanation_tags_if_text_present_odt(root, namespaces):
    """
    Answer ke baad agar text hai to "Explanation:" tag add karta hai
    FIXED: Better parent handling
    """
    paragraphs = root.xpath('//text:p', namespaces=namespaces)
    
    # Collect paragraphs that need explanation tags
    explanation_insertions = []
    
    answer_found = False
    explanation_added = False
    
    for i, para in enumerate(paragraphs):
        para_text = get_para_text(para).strip()
        
        # Check if this is Answer: line
        if para_text.startswith("Answer:"):
            answer_found = True
            explanation_added = False
            continue
        
        # Check if this is next Question:
        if para_text.startswith("Question:"):
            answer_found = False
            explanation_added = False
            continue
        
        # If we're after Answer: and haven't added Explanation yet
        if answer_found and not explanation_added and para_text:
            # Check if line already has Explanation:
            if not para_text.startswith("Explanation:") and not para_text.startswith("Reference:"):
                explanation_insertions.append({
                    'para': para,
                    'before': True  # Insert before this paragraph
                })
                explanation_added = True
    
    # Apply all insertions
    for insertion in explanation_insertions:
        para = insertion['para']
        parent = para.getparent()
        
        if parent is not None:
            try:
                children = list(parent)
                if para in children:
                    index = children.index(para)
                    
                    # Create new paragraph with Explanation:
                    explanation_para = etree.Element(f"{{{namespaces['text']}}}p")
                    explanation_para.text = "Explanation:"
                    
                    # Insert before current paragraph
                    parent.insert(index, explanation_para)
            except ValueError:
                # Element not found - skip
                continue
    
    return root


# def add_line_spacing_after_question_answer_odt(root, namespaces):
#     """
#     Question: aur Answer: ke baad proper line spacing add karta hai
#     """
#     paragraphs = root.xpath('//text:p', namespaces=namespaces)
#     parent = None
    
#     # Find parent element
#     for para in paragraphs:
#         parent = para.getparent()
#         if parent is not None:
#             break
    
#     if parent is None:
#         return root
    
#     # Process paragraphs in reverse to avoid index issues when inserting
#     paragraphs_list = list(paragraphs)
    
#     for i in range(len(paragraphs_list) - 1, -1, -1):
#         para = paragraphs_list[i]
#         para_text = get_para_text(para).strip()
        
#         # Add spacing after Question: lines
#         if para_text.startswith("Question:"):
#             index = list(parent).index(para)
            
#             # Add empty paragraph after
#             empty_para = etree.Element(f"{{{namespaces['text']}}}p")
#             empty_para.text = ""
#             parent.insert(index + 1, empty_para)
        
#         # Add spacing after Answer: lines
#         elif para_text.startswith("Answer:"):
#             index = list(parent).index(para)
            
#             # Add empty paragraph after
#             empty_para = etree.Element(f"{{{namespaces['text']}}}p")
#             empty_para.text = ""
#             parent.insert(index + 1, empty_para)
    
#     return root


def add_line_spacing_after_question_answer_odt(root, namespaces):
    """
    Question: aur Answer: ke baad proper line spacing add karta hai
    FIXED: Handles parent modifications correctly
    """
    paragraphs = root.xpath('//text:p', namespaces=namespaces)
    
    # Collect all spacing insertions
    spacing_insertions = []
    
    for para in paragraphs:
        para_text = get_para_text(para).strip()
        
        # Mark paragraphs that need spacing after them
        if para_text.startswith("Question:") or para_text.startswith("Answer:"):
            spacing_insertions.append(para)
    
    # Apply spacing insertions in reverse order to maintain indices
    # Reverse order ensures that indices remain valid
    for para in reversed(spacing_insertions):
        parent = para.getparent()
        
        if parent is not None:
            try:
                children = list(parent)
                if para in children:
                    index = children.index(para)
                    
                    # Create empty paragraph for spacing
                    empty_para = etree.Element(f"{{{namespaces['text']}}}p")
                    empty_para.text = ""
                    
                    # Insert after current paragraph
                    parent.insert(index + 1, empty_para)
            except ValueError:
                # Element not found - skip
                continue
    
    return root


# ==================== MAIN ODT PROCESSING ====================

def process_odt_file(input_file, output_file, send_status_update):
    """
    Main function to process ODT file with all operations
    """
    temp_dir = tempfile.mkdtemp()
    
    try:
        # Step 0: Validate ODT file
        send_status_update({
            'type': 'status',
            'name': 'Validate ODT File',
            'status': 'in_progress',
            'time': '0.00s'
        })
        start_time = time.time()
        
        is_valid, issues = validate_odt_file(input_file)
        
        if not is_valid:
            send_status_update({
                'type': 'status',
                'name': 'Validate ODT File',
                'status': 'failed',
                'time': f"{(time.time() - start_time):.2f}s"
            })
            raise Exception(f"Invalid ODT file. Issues found:\n" + "\n".join(f"- {issue}" for issue in issues))
        
        send_status_update({
            'type': 'status',
            'name': 'Validate ODT File',
            'status': 'completed',
            'time': f"{(time.time() - start_time):.2f}s"
        })
        
        # Extract ODT file
        with zipfile.ZipFile(input_file, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)
        
        # Load and parse content.xml
        content_file = os.path.join(temp_dir, 'content.xml')
        tree = etree.parse(content_file)
        root = tree.getroot()
        
        # Step 1: Fix Question Numbers and Remove Brackets
        send_status_update({
            'type': 'status',
            'name': 'Fix Question Numbers & Remove Brackets',
            'status': 'in_progress',
            'time': '0.00s'
        })
        start_time = time.time()
        root, questions_fixed = fix_odt_question_numbers_and_brackets(root, ODT_NAMESPACES)
        send_status_update({
            'type': 'status',
            'name': 'Fix Question Numbers & Remove Brackets',
            'status': 'completed',
            'time': f"{(time.time() - start_time):.2f}s"
        })
        
        # Step 2: Ensure Spacing Before Questions
        send_status_update({
            'type': 'status',
            'name': 'Ensure Spacing Before Questions',
            'status': 'in_progress',
            'time': '0.00s'
        })
        start_time = time.time()
        root = ensure_spacing_before_question_tags_odt(root, ODT_NAMESPACES)
        send_status_update({
            'type': 'status',
            'name': 'Ensure Spacing Before Questions',
            'status': 'completed',
            'time': f"{(time.time() - start_time):.2f}s"
        })
        
        # Step 3: Combined Text Operations
        send_status_update({
            'type': 'status',
            'name': 'Combined Text Operations',
            'status': 'in_progress',
            'time': '0.00s'
        })
        start_time = time.time()
        root = combined_text_operations_odt(root, ODT_NAMESPACES)
        send_status_update({
            'type': 'status',
            'name': 'Combined Text Operations',
            'status': 'completed',
            'time': f"{(time.time() - start_time):.2f}s"
        })
        
        # Step 4: Question Types to Next Line
        send_status_update({
            'type': 'status',
            'name': 'Question Types to Next Line',
            'status': 'in_progress',
            'time': '0.00s'
        })
        start_time = time.time()
        root = shift_question_types_to_next_line_odt(root, ODT_NAMESPACES)
        send_status_update({
            'type': 'status',
            'name': 'Question Types to Next Line',
            'status': 'completed',
            'time': f"{(time.time() - start_time):.2f}s"
        })
        
        # Step 5: Normalize Option Spacing
        send_status_update({
            'type': 'status',
            'name': 'Normalize Option Spacing',
            'status': 'in_progress',
            'time': '0.00s'
        })
        start_time = time.time()
        root = normalize_option_spacing_odt(root, ODT_NAMESPACES)
        send_status_update({
            'type': 'status',
            'name': 'Normalize Option Spacing',
            'status': 'completed',
            'time': f"{(time.time() - start_time):.2f}s"
        })
        
        # Step 6: Add Explanation Tags
        send_status_update({
            'type': 'status',
            'name': 'Add Explanation Tags',
            'status': 'in_progress',
            'time': '0.00s'
        })
        start_time = time.time()
        root = add_explanation_tags_if_text_present_odt(root, ODT_NAMESPACES)
        send_status_update({
            'type': 'status',
            'name': 'Add Explanation Tags',
            'status': 'completed',
            'time': f"{(time.time() - start_time):.2f}s"
        })
        
        # Step 7: Add Line Spacing
        send_status_update({
            'type': 'status',
            'name': 'Add Line Spacing',
            'status': 'in_progress',
            'time': '0.00s'
        })
        start_time = time.time()
        root = add_line_spacing_after_question_answer_odt(root, ODT_NAMESPACES)
        send_status_update({
            'type': 'status',
            'name': 'Add Line Spacing',
            'status': 'completed',
            'time': f"{(time.time() - start_time):.2f}s"
        })
        
        # Save modified XML
        tree.write(content_file, xml_declaration=True, encoding='UTF-8', pretty_print=True)
        
        # Create output ODT file
        with zipfile.ZipFile(output_file, 'w', zipfile.ZIP_DEFLATED) as zip_ref:
            for foldername, subfolders, filenames in os.walk(temp_dir):
                for filename in filenames:
                    file_path = os.path.join(foldername, filename)
                    arcname = os.path.relpath(file_path, temp_dir)
                    zip_ref.write(file_path, arcname)
        
        return questions_fixed
    
    except Exception as e:
        raise Exception(f"ODT processing failed: {str(e)}")
    
    finally:
        # Cleanup temp directory
        shutil.rmtree(temp_dir, ignore_errors=True)


# ==================== FLASK ROUTES ====================

@app.route('/')
def index():
    return render_template('index.html')


@app.route('/status')
def status():
    """SSE endpoint for real-time status updates"""
    def generate():
        while True:
            try:
                # Wait for status update with timeout
                update = status_queue.get(timeout=30)
                yield f"data: {json.dumps(update)}\n\n"
                
                # If processing is complete or error, stop the stream
                if update.get('type') in ['complete', 'error']:
                    break
            except queue.Empty:
                # Send keepalive
                yield f"data: {json.dumps({'type': 'keepalive'})}\n\n"
    
    return Response(generate(), mimetype='text/event-stream')


def send_status_update(update):
    """Helper function to send status updates"""
    status_queue.put(update)


@app.route('/upload', methods=['POST'])
def upload_file():
    """Handle file upload and processing"""
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400
    
    file = request.files['file']
    
    if file.filename == '':
        return jsonify({'error': 'No file selected'}), 400
    
    # Check file extension
    if not file.filename.lower().endswith('.odt'):
        return jsonify({'error': 'Only ODT files are supported'}), 400
    
    # Save uploaded file temporarily
    filename = secure_filename(file.filename)
    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.odt')
    temp_file_path = temp_file.name
    file.save(temp_file_path)
    temp_file.close()
    
    def process_in_background():
        """Background processing function"""
        try:
            start_total_time = time.time()
            
            # Create output file path
            temp_output = tempfile.NamedTemporaryFile(delete=False, suffix='.odt')
            temp_output.close()
            
            # Process the ODT file
            questions_fixed = process_odt_file(temp_file_path, temp_output.name, send_status_update)
            
            # Read processed file
            with open(temp_output.name, 'rb') as f:
                output_data = f.read()
            
            # Clean up temp files
            os.unlink(temp_file_path)
            os.unlink(temp_output.name)
            
            total_time = time.time() - start_total_time
            
            # Save to uploads folder
            output_filename = f"processed_{filename}"
            output_path = os.path.join(app.config['UPLOAD_FOLDER'], output_filename)
            with open(output_path, 'wb') as f:
                f.write(output_data)
            
            send_status_update({
                'type': 'complete',
                'filename': output_filename,
                'total_time': f"{total_time:.2f}",
                'questions_fixed': questions_fixed
            })
        
        except Exception as e:
            send_status_update({
                'type': 'error',
                'message': str(e)
            })
            # Clean up temp file on error
            if os.path.exists(temp_file_path):
                os.unlink(temp_file_path)
    
    # Start background thread
    thread = threading.Thread(target=process_in_background)
    thread.daemon = True
    thread.start()
    
    return jsonify({'success': True, 'message': 'Processing started'})


@app.route('/download/<filename>')
def download_file(filename):
    """Download processed file"""
    try:
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        return send_file(file_path, as_attachment=True)
    except Exception as e:
        return jsonify({'error': str(e)}), 404


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)