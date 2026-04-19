# PDF Data Extraction Web Application

## How to Download and Run

### 1. Prerequisites
Before running the application, make sure you have the following installed:
- **Python 3.9+**: [Download Python](https://www.python.org/downloads/)
- **Tesseract OCR** (Required for scanned PDFs):
  - **Windows**: [Download Installer](https://github.com/UB-Mannheim/tesseract/wiki)
  - **Mac**: `brew install tesseract`
  - **Linux**: `sudo apt-get install tesseract-ocr`

### 2. Installation

1.  **Download the project code** to your local machine.

2.  **Open a terminal** (Command Prompt or PowerShell) in the project folder.

3.  **Create a virtual environment** (recommended):
    ```bash
    python -m venv venv
    ```

4.  **Activate the virtual environment**:
    - **Windows**:
      ```bash
      venv\Scripts\activate
      ```
    - **Mac/Linux**:
      ```bash
      source venv/bin/activate
      ```

5.  **Install the dependencies**:
    ```bash
    pip install -r requirements.txt
    ```

### 3. Running the Application

1.  Make sure your virtual environment is activated.

2.  Run the application:
    ```bash
    python app/app.py
    ```

3.  Open your web browser and go to:
    ```
    http://127.0.0.1:5000
    ```
