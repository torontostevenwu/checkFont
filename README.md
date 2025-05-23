# Document Font Checker

This tool allows you to upload Microsoft Word (.docx) and PDF (.pdf) documents to identify the font types used within them.

## Installation

1.  Clone this repository:
    ```bash
    git clone <repository_url> 
    cd <repository_directory>
    ```
2.  Install the required Python packages:
    ```bash
    pip install -r requirements.txt
    ```

## Usage

1.  Run the Streamlit application:
    ```bash
    streamlit run checkFontTypes.py
    ```
2.  Open the application in your web browser (usually at `http://localhost:8501`).
3.  Upload your `.docx` or `.pdf` file using the file uploader.
4.  The application will display the list of unique fonts found in the document.

## Running Tests

To run the unit tests, execute the following command in the root directory:
```bash
python -m unittest test_checkFontTypes.py
```
