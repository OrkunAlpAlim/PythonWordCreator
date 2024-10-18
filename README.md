# Creating a Word Document with Python - `python-docx` Library

This project demonstrates how to create Word documents using the `python-docx` library in Python. It covers the essential functions and basic text formatting features required to build Word documents programmatically.

## Installation

To use the `python-docx` library, follow the installation instructions below.

### Automatic Installation

The script checks if the `python-docx` library is installed and installs it if necessary:

```python
import subprocess
import sys

def check_install(package): 
    try:
        __import__(package)
        print(f"{package} is already installed.")
    except ImportError:
        print(f"{package} is not installed. Installing...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])

check_install("python-docx")
```

### Manual Installation

To manually install the `python-docx` library:

```bash
pip install python-docx
```

## Required Imports

The following modules are used for creating and formatting Word documents:

```python
from docx import Document
from docx.shared import Pt, RGBColor
```

## Project Details

### 1. Creating a Word Document

To create a Word document, use the `Document()` class:

```python
doc = Document()
```

### 2. Adding a Heading

To add a heading to the document, use the `add_heading()` function:

```python
doc.add_heading('Word Document Heading', level=1)
```

### 3. Adding and Formatting Text in a Paragraph

Text within a paragraph can be added and formatted using `add_paragraph()` and `add_run()`:

```python
paragraph = doc.add_paragraph()
run = paragraph.add_run('This text is bold and blue.')
run.font.bold = True
run.font.color.rgb = RGBColor(0, 0, 255)
```

### 4. Font Properties

You can customize the text with the following font properties:

- **Font Size**: Set the font size using `Pt`:
  ```python
  run.font.size = Pt(20)
  ```

- **Font Family**: Specify the font name:
  ```python
  run.font.name = 'Arial'
  ```

- **Bold Text**: Make the text bold:
  ```python
  run.font.bold = True
  ```

- **Italic Text**: Make the text italic:
  ```python
  run.font.italic = True
  ```

- **Text Color**: Set the text color in RGB format:
  ```python
  run.font.color.rgb = RGBColor(255, 0, 0)
  ```

### 5. Saving the File

To save the document to a specific path:

```python
doc.save('example.docx')
```
