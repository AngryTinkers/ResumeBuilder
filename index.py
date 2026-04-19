import os
from docx import Document


def replace_text_in_docx(file_path, search_text, replace_text):
  doc = Document(file_path)

  for paragraph in doc.paragraphs:
    for run in paragraph.runs:
      if search_text in run.text:
        run.text = run.text.replace(search_text, replace_text)

  # Tables need special handling
  for table in doc.tables:
    for row in table.rows:
      for cell in row.cells:
        for paragraph in cell.paragraphs:
          for run in paragraph.runs:
            # print(f":::::::::::::: {run.text}")
            if search_text in run.text:
              run.text = run.text.replace(search_text, replace_text)

  doc.save(file_path)

def process_folder(root_folder, search_text, replace_text):
  for root, dirs, files in os.walk(root_folder):
    for file in files:
      if file.startswith("~$"):
        continue
      if file.endswith(".docx"):
        file_path = os.path.join(root, file)
        replace_text_in_docx(file_path, search_text, replace_text)

# Usage
process_folder(
  root_folder="resumes/",
  search_text="justintech130",
  replace_text="justin.raines.dev"
)