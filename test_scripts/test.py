from format import markdown_to_docx
import re

# Load the new markdown file and convert it
with open("md.md", "r") as file:
    markdown_text = file.read()

first_line = markdown_text.split('\n', 1)[0]
filename = re.sub(r'\W+', '_', first_line)

docx_file_path = markdown_to_docx(markdown_text,filename)
