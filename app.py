import streamlit as st
from format import markdown_to_docx
import re

hide_streamlit_style = """
<style>
    #root > div:nth-child(1) > div > div > div > div > section > div {padding-top: 1rem;}
</style>

"""
st.markdown(hide_streamlit_style, unsafe_allow_html=True)

# Step 3: Create UI elements
markdown_text = st.text_area('', height=500)
convert_button = st.button('Vytvořit Microsoft Word')

# Step 4: Handle button click event
if convert_button:
  if markdown_text:
    first_line = markdown_text.split('\n', 1)[0]
    filename = re.sub(r'\W+', '_', first_line)

    # Create a temporary directory to store the user's file
    with tempfile.TemporaryDirectory() as tempdir:
      # The temporary directory is automatically cleaned up when you exit this block

      # Call the markdown_to_docx function, but inside the temporary directory
      docx_file_path = markdown_to_docx(markdown_text, filename)

      # Use the full path for the temporary file
      temp_file_path = os.path.join(tempdir, docx_file_path)

      # Move the generated docx to the temporary directory
      shutil.move(docx_file_path, temp_file_path)

      # Serve the file for download from the temporary directory
      with open(temp_file_path, 'rb') as file:
        st.download_button(
            label="Stáhnout Microsoft Word",
            data=file,
            file_name=f"{filename}.docx",
            mime=
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

      # The temporary file and directory will be automatically deleted after serving the download

  else:
    st.warning('Please enter some Markdown text before trying to convert it.')
