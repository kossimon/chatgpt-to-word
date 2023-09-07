import re
import docx
from docx.shared import Inches


def preprocess(markdown_text):
    """
    Preprocess the markdown text to replace certain patterns with a standardized notation.
    """
    # Split the text into lines
    lines = markdown_text.split('\n')

    # Find lines that are marked to be converted to heading 2
    heading_2_indices = set()
    for i in range(0,len(lines)):
        if lines[i].startswith("---") and lines[i - 1].strip():
            lines[i] = ""
            heading_2_indices.add(i - 1)
            
    # Apply the replace_pattern function on each line and mark certain lines as heading 2
    for i, line in enumerate(lines):
        lines[i] = replace_pattern(line)
        if i in heading_2_indices:
            print(lines[i])
            print(f'## {lines[i]}')
            lines[i] = "## " + lines[i]
            print(lines[i])
    
    # Join the lines back into a single string
    return lines

def replace_pattern(s):
    return re.sub(r'\[\^(.*?)\^\]', r'^\1', s)

def process_for_formatting(line):
    '''
    Process the lines and finds standard notation for formatting and creates runs.
    '''
    runs = []
    i = 0

    # Special handling for lines starting with ^\d+ for bold numbers
    if line.startswith('^') and re.match(r"^\^\d+", line):
        match = re.match(r"^\^(\d+)", line)
        if match:
            runs.append((match.group(1), 'bold'))
            i += len(match.group(0))

    while i < len(line):
        # Superscript within the line (we'll handle it later in process_for_superscript)
        if line[i] == '^':
            end_sup = line.find(' ', i)  # We assume that superscript ends at a space
            if end_sup != -1:
                runs.append((line[i:end_sup], 'normal'))
                i = end_sup
            else:
                runs.append((line[i:], 'normal'))
                i = len(line)
            continue

        if line[i:i+3] == '***':  # Bold and Italic
            end_bold_italic = line.find('***', i + 3)
            if end_bold_italic != -1:
                runs.append((line[i+3:end_bold_italic], 'bold_italic'))
                i = end_bold_italic + 3
            else:
                runs.append((line[i], 'normal'))
                i += 1

        elif line[i:i+2] == '**':  # Bold
            end_bold = line.find('**', i + 2)
            if end_bold != -1:
                inner_runs = process_for_formatting(line[i+2:end_bold])  # Recursive call to handle nested formatting
                for run_text, run_style in inner_runs:
                    if run_style == 'italic':
                        runs.append((run_text, 'bold_italic'))
                    else:
                        runs.append((run_text, 'bold'))
                i = end_bold + 2
            else:
                runs.append((line[i], 'normal'))
                i += 1

        elif line[i] == '*':  # Italic
            end_italic = line.find('*', i + 1)
            if end_italic != -1:
                runs.append((line[i+1:end_italic], 'italic'))
                i = end_italic + 1
            else:
                runs.append((line[i], 'normal'))
                i += 1

        else:
            next_bold = line.find('**', i)
            next_italic = line.find('*', i)
            next_bold_italic = line.find('***', i)
            next_superscript = line.find('^', i)

            next_special = min([pos for pos in [next_bold, next_italic, next_bold_italic, next_superscript] if pos != -1], default=len(line))

            runs.append((line[i:next_special], 'normal'))
            i = next_special

    return runs

def process_for_superscript(runs):
    """
    Post-process the runs to handle superscript.
    """
    processed_runs = []

    for run_text, run_style in runs:
        i = 0
        while i < len(run_text):
            if run_text[i] == '^':
                match = re.match(r"\^(\d+)", run_text[i:])
                if match:
                    superscript_text = match.group(1)
                    if run_style == 'normal':
                        processed_runs.append((superscript_text, 'superscript'))
                    elif run_style == 'bold':
                        processed_runs.append((superscript_text, 'bold_superscript'))
                    elif run_style == 'italic':
                        processed_runs.append((superscript_text, 'italic_superscript'))
                    elif run_style == 'bold_italic':
                        processed_runs.append((superscript_text, 'bold_italic_superscript'))
                    i += len(match.group(0))
                    continue
            processed_runs.append((run_text[i], run_style))
            i += 1

    return processed_runs

def process_for_numbered_lists(lines):
    """
    Post-process the lines to handle numbered lists.
    """
    processed_lines = []
    in_numbered_list = False

    for line in lines:
        if re.match(r"\d+\.", line.strip()):
            if not in_numbered_list:
                in_numbered_list = True
                processed_lines.append({"type": "start_numbered_list"})
            # Extract the number and the content after it
            num, content = line.strip().split('.', 1)
            processed_lines.append({"type": "list_item", "content": content.strip()})
        else:
            if in_numbered_list:
                in_numbered_list = False
                processed_lines.append({"type": "end_numbered_list"})
            processed_lines.append({"type": "normal", "content": line})

    if in_numbered_list:
        processed_lines.append({"type": "end_numbered_list"})

    return processed_lines

def process_for_bullet_points(lines):
    """
    Post-process the lines to handle bullet points.
    """
    processed_lines = []
    in_bullet_point_list = False

    for line in lines:
        if isinstance(line, dict):
            processed_lines.append(line)
            continue

        if line.strip().startswith('- ') or line.strip().startswith('* '):
            if not in_bullet_point_list:
                in_bullet_point_list = True
                processed_lines.append({"type": "start_bullet_point_list"})
            content = line.strip()[2:]
            processed_lines.append({"type": "bullet_point", "content": content})
        else:
            if in_bullet_point_list:
                in_bullet_point_list = False
                processed_lines.append({"type": "end_bullet_point_list"})
            processed_lines.append(line)

    if in_bullet_point_list:
        processed_lines.append({"type": "end_bullet_point_list"})

    return processed_lines


def get_heading_level(content):
    """
    Check for headings and return the level (0-5) and content without the markdown. 
    If not a heading, return None.
    """
    headings = [
        ("# ", 0),
        ("## ", 1),
        ("### ", 2),
        ("#### ", 3),
        ("##### ", 4),
        ("###### ", 5)
    ]
    for prefix, level in headings:
        if content.startswith(prefix):
            return level, content[len(prefix):]
    return None, content

def process_run_styles(run, run_style):
    """
    Set the styles for a run based on the given run_style.
    """
    if 'bold' in run_style:
        run.bold = True
    if 'italic' in run_style:
        run.italic = True
    if 'superscript' in run_style:
        run.font.superscript = True

def add_runs_to_paragraph(paragraph, runs):
    """
    Process the runs and add them to the given paragraph.
    """
    for run_text, run_style in runs:
        run = paragraph.add_run(run_text)
        process_run_styles(run, run_style)

def process_line(doc, line, in_numbered_list, in_bullet_point_list):
    """
    Process a line and add its content to the document.
    Return the updated in_numbered_list and in_bullet_point_list flags.
    """
    if line["type"] == "start_numbered_list":
        return True, in_bullet_point_list
    elif line["type"] == "end_numbered_list":
        return False, in_bullet_point_list
    elif line["type"] == "start_bullet_point_list":
        return in_numbered_list, True
    elif line["type"] == "end_bullet_point_list":
        return in_numbered_list, False

    # Check for headings using the get_heading_level function
    heading_level, content = get_heading_level(line["content"])

    # Check for leading spaces
    leading_spaces = len(content) - len(content.lstrip(' '))
    indentation_level = leading_spaces // 2 * 0.25  # For each two spaces, we indent 0.25 inches

    # Remove the leading spaces from the content
    content = content.lstrip()

    # Process the line for formatting
    runs = process_for_formatting(content)
    # Further process the runs for superscript
    runs = process_for_superscript(runs)

    # If the line was identified as a heading, add it as a heading
    if heading_level is not None:
        heading = doc.add_heading('', level=heading_level)
        add_runs_to_paragraph(heading, runs)
    elif in_numbered_list or line["type"] == "list_item":
        paragraph = doc.add_paragraph(style='ListNumber')
        add_runs_to_paragraph(paragraph, runs)
    elif in_bullet_point_list or line["type"] == "bullet_point":
        paragraph = doc.add_paragraph(style='ListBullet')
        add_runs_to_paragraph(paragraph, runs)
    else:
        # For normal lines
        paragraph = doc.add_paragraph()
        # Set the indentation for the paragraph if there were leading spaces
        if indentation_level:
            paragraph.paragraph_format.left_indent = docx.shared.Inches(indentation_level)
        add_runs_to_paragraph(paragraph, runs)

    return in_numbered_list, in_bullet_point_list

def markdown_to_docx(markdown_text,filename):
    doc = docx.Document()

    # Preprocess the markdown text
    lines = preprocess(markdown_text)

    # Post-process for numbered lists
    lines = process_for_numbered_lists(lines)
    
    # Post-process for bullet points
    lines = process_for_bullet_points(lines)

    in_numbered_list = False
    in_bullet_point_list = False

    for line in lines:
        in_numbered_list, in_bullet_point_list = process_line(doc, line, in_numbered_list, in_bullet_point_list)
    
    filename = filename[:42]
    file_path = f"{filename}.docx"
    doc.save(file_path)

    return file_path
