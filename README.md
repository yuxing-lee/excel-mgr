# Excel Manager

This simple Flask application provides a web interface to perform CRUD (create, read, update, delete) operations on a local Excel file. Each record includes an **Option** field with values `test1`, `test2` or `test`.

## Setup

1. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
2. Run the server:
   ```bash
   python app.py
   ```
3. Open `http://localhost:5000` in your browser.

The application automatically creates `data.xlsx` if it does not exist.

## Building a Windows executable

You can package the server as a stand-alone Windows executable using
`PyInstaller`:

```bash
pip install pyinstaller
pyinstaller --onefile --add-data "templates;templates" app.py
```

The generated `dist/app.exe` starts the server and creates `data.xlsx` in the
same directory when first run.
