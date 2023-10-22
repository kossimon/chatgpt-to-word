import re
import docx
from docx.shared import Inches

def indentate_lines(markdown_text):
    lines = markdown_text.split('\n')
    indented = []

     # Initialize the count of spaces and the indent level.
    # Iterate through each character in the line.
    for line in lines:
        space_count = 0
        for char in line:
            if char == ' ':
                # If the character is a space, increment the space_count.
                space_count += 1
            else:
                # If a non-space character is encountered, break the loop.
                break

        # Map count of spaces to indent level
        mapping = {0:0,2:1,3:1,4:2,5:2,6:2}
        indent_level = mapping.get(space_count, 0)

        # Create and return the dictionary with the content and indent level.
        indented.append({'content': line.strip(), 'indent': indent_level})
    return indented


def preprocess(indentated_lines):
    # We're directly dealing with dictionaries inside the list here.
    # So, each 'line_info' is a dictionary with 'content' and 'indent'.
    
    # First, we process the lines for heading styles.
    heading_2_indices = set()
    for i, line_info in enumerate(indentated_lines):
        if line_info['content'].startswith("---") and i > 0 and indentated_lines[i - 1]['content'].strip():
            indentated_lines[i]['content'] = ""  # Clear the current '---' line
            heading_2_indices.add(i - 1)

    # Apply transformations on each line's content.
    for i, line_info in enumerate(indentated_lines):
        line = replace_pattern(line_info['content'])
        if i in heading_2_indices:
            line = "## " + line  # Add Markdown heading syntax for level 2 headings
        indentated_lines[i]['content'] = line  # Update the content after transformations

    return indentated_lines


def replace_pattern(s):
    s = re.sub(r'\[\^(.*?)\^\]', r'^\1', s)
    s = re.sub(r'\[(.*?)\]', r'^\1', s)
    s = re.sub(r'^\*\s', '- ', s)
    return s

def process_for_formatting(line):
    runs = []
    i = 0
    inside_bold = False
    inside_italic = False

    while i < len(line):
        if line[i:i+3] == '***':
            if not inside_bold and not inside_italic:
                start = i + 3
                end = line.find('***', start)
                if end != -1:
                    runs.append((line[start:end], 'bold_italic'))
                    i = end + 3
                    continue
            # if the end marker was not found or we're inside other styles
            runs.append((line[i:i+3], 'normal'))
            i += 3

        elif line[i:i+2] == '**':
            inside_bold = not inside_bold  # toggle the state
            i += 2  # move past this marker

        elif line[i] == '*':
            inside_italic = not inside_italic  # toggle the state
            i += 1  # move past this marker

        else:
            # Find the next marker, considering we are not at the start of a marker now
            next_marker = len(line)
            for marker in ['***', '**', '*']:
                marker_pos = line.find(marker, i)
                if marker_pos != -1:
                    next_marker = min(next_marker, marker_pos)

            style = 'normal'
            if inside_bold and inside_italic:
                style = 'bold_italic'
            elif inside_bold:
                style = 'bold'
            elif inside_italic:
                style = 'italic'

            runs.append((line[i:next_marker], style))
            i = next_marker  # move to the next marker position

    return runs



def process_for_superscript(runs):
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



def add_run(paragraph, text, style):
    run = paragraph.add_run(text)
    if style == 'bold':
        run.bold = True
    elif style == 'italic':
        run.italic = True
    elif style == 'bold_italic':
        run.bold = True
        run.italic = True
    elif style == 'superscript':
        run.font.superscript = True
    elif style == 'bold_superscript':
        run.bold = True
        run.font.superscript = True
    elif style == 'italic_superscript':
        run.italic = True
        run.font.superscript = True
    elif style == 'bold_italic_superscript':
        run.bold = True
        run.italic = True
        run.font.superscript = True

def process_headings(line, doc):
    # Define markdown heading levels
    headings = [
        ("# ", 1),
        ("## ", 2),
        ("### ", 3),
        ("#### ", 4),
        ("##### ", 5),
        ("###### ", 6)
    ]

    # Check if the current line starts with a markdown heading prefix
    for prefix, level in headings:
        if line.startswith(prefix):
            # Remove the markdown prefix from the heading text
            text = line[len(prefix):]
            # Add the heading to the document with the appropriate level
            doc.add_heading(text, level=level)
            return True  # Heading was processed

    return False  # No heading was found

def process_bullets(line_info, doc):
    # Check if the current line starts with the bullet point markdown
    line = line_info['content']
    indent_level = line_info['indent']
    if line.startswith("- "):
        # Extract the actual content, excluding the markdown bullet point ("- ")
        content = line[2:]

        # Depending on the indent level, we might want to adjust the style
        if indent_level == 0:
            style = 'List Bullet'
        elif indent_level == 1:
            style = 'List Bullet 2'  # Assuming 'List Bullet 2' is defined in your Word styles
        elif indent_level == 2:
            style = 'List Bullet 3'  # Assuming 'List Bullet 3' is defined in your Word styles
        else:
            style = 'List Bullet'  # Default to normal bullets for deeper indents or unexpected cases

        # Create a new bullet point with the specified style
        bullet = doc.add_paragraph(style=style)
        bullet.paragraph_format.line_spacing = 1.15
        bullet.paragraph_format.space_after = Inches(0.05)

        # Instead of adding text directly, process the content for formatting
        runs = process_for_formatting(content)
        runs = process_for_superscript(runs)
        for run_text, run_style in runs:
            add_run(bullet, run_text, run_style)  # Apply the styles per run

        return True  # The line was a bullet point and was processed

    return False  # The line was not a bullet point

def process_numbered_lists(line_info, doc):
    # Check if the current line starts with a markdown numbered list format
    line = line_info['content']
    indent_level = line_info['indent']

    # This regular expression matches lines that start with "1. ", "1.1. ", etc.
    # It handles multi-digit numbers and a single following space.
    match = re.match(r'(\d+\.)+\s', line)
    if match:
        # Extract the actual content, including the markdown numbered prefix (e.g., "1. ")
        number = line[:match.end()]  # Include the numbers and period
        content = line[match.end():]  # Extract the content after the numbers

        number = '*' + number + '* '
        whole = number + content

        # Create a new paragraph and set the line height to 1.0
        number_item = doc.add_paragraph()
        number_item.paragraph_format.line_spacing = 1.15
        number_item.paragraph_format.space_after = Inches(0.05)

        # Apply the indent level
        number_item.paragraph_format.left_indent = Inches(0.25 * indent_level)  # 36 points per indent level (adjust as needed)

         # Instead of adding text directly, process the content for formatting
        runs = process_for_formatting(whole)
        runs = process_for_superscript(runs)
        for run_text, run_style in runs:
            add_run(number_item, run_text, run_style)  # Apply the styles per run

        return True  # The line was a bullet point and was processed

    return False  # The line was not a numbered list item




def process_lines(doc, lines):
    for line_info in lines:
        # If the line is a heading, a bullet point, or a numbered list, it's processed inside these functions
        if process_headings(line_info['content'], doc):
            continue
        elif process_bullets(line_info, doc):
            continue
        elif process_numbered_lists(line_info, doc):
            continue
        else:
            # If the line is a regular line (not a heading, bullet, or numbered list), create a new paragraph for it
            paragraph = doc.add_paragraph()

            # Adjust the left indent based on the indentation level.
            if line_info['indent'] > 0:
                paragraph.paragraph_format.left_indent = Inches(0.25 * line_info['indent'])

            # Process the content for formatting (bold, italic, etc.)
            runs = process_for_formatting(line_info['content'])
            runs = process_for_superscript(runs)
            for run_text, run_style in runs:
                add_run(paragraph, run_text, run_style)




def markdown_to_docx(markdown_text, filename):
    doc = docx.Document()

    # Indentate lines before processing
    indentated_lines = indentate_lines(markdown_text)

    # Preprocess the lines. Note that we're passing the list of line dictionaries here,
    # not a single string.
    preprocessed_lines = preprocess(indentated_lines)

    # The remaining part of your function stays the same
    process_lines(doc, preprocessed_lines)

    filename = filename[:42]  # Ensure the filename is not excessively long
    file_path = f"{filename}.docx"
    doc.save(file_path)

    return file_path


