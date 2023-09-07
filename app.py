import streamlit as st
from format import markdown_to_docx
import re

st.title('chatGPT to Microsoft Word')

# Step 3: Create UI elements
markdown_text = st.text_area('Copy and paste and click convert:', height=500)
convert_button = st.button('Convert to Microsoft Word')

# Step 4: Handle button click event
if convert_button:
    if markdown_text:
        # Get the first line of the markdown text to use as the filename
        first_line = markdown_text.split('\n', 1)[0]
        filename = re.sub(r'\W+', '_', first_line)

        # Call the markdown_to_docx function to convert the input text to a DOCX file
        docx_file_path = markdown_to_docx(markdown_text, filename)
        
        # Step 5: Provide a button to download the DOCX file
        with open(docx_file_path, 'rb') as file:
            st.download_button(
                label="Download Word",
                data=file,
                file_name=f"{filename}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
    else:
        st.warning('Please enter some Markdown text before trying to convert it.')
