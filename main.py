from docx import Document
import os

# Create a new Word document
doc = Document()

my_string = "Hello World! Here is written by Python."
doc_name = "example.docx"

# Add a title to the document
doc.add_paragraph(my_string)

# Save the documente
doc.save(doc_name)

# Check if the document was created successfully
if os.path.exists(doc_name):
    print("The Word file was created successfully, and saved as" + doc_name)
else:
    print("Happend a problem while creating the Word file")
