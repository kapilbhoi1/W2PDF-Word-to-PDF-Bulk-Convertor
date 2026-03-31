# W2PDF - Word to PDF Bulk Converter

A sleek, dark-themed Streamlit web application that converts multiple Word documents (`.docx` / `.doc`) into individual PDF files in bulk — with real-time progress tracking, individual downloads, and a single ZIP bundle option.

---

## Features

- **Bulk Conversion** — Upload multiple Word files and convert them all to PDF in one click.
- **High-Fidelity Output** — Uses Microsoft Word's COM automation (`docx2pdf`) for pixel-perfect PDF rendering.
- **Real-Time Progress** — Live progress bar and status text during conversion.
- **Flexible Downloads** — Download each PDF individually or grab all of them as a single `.zip` archive.
- **Dark Mode UI** — Professional, modern dark interface with purple/blue gradient accents.
- **100% Private** — All processing happens entirely on your local machine. No files leave your computer.

---

## Screenshots

> Upload your Word documents on the left. Preview and convert on the right.

---

## Prerequisites

Before running this application, make sure you have the following installed:

| Requirement         | Details                                                                 |
|---------------------|-------------------------------------------------------------------------|
| **Python 3.8+**     | [Download Python](https://www.python.org/downloads/)                    |
| **Microsoft Word**  | Must be installed on your machine (used internally by `docx2pdf`)       |
| **pip**             | Comes bundled with Python. Used to install dependencies.                |

> **Important:** This application requires **Microsoft Word** to be installed because `docx2pdf` uses Word's COM automation to perform the conversion. It will **not work** on systems without Word installed (e.g., most Linux servers).

---

## Project Structure

```
Word-to-PDF-Bulk/
├── .streamlit/
│   └── config.toml          # Streamlit theme configuration (dark mode)
├── app.py                    # Main application (UI + conversion logic)
├── requirements.txt          # Python dependencies
├── Run_Converter.bat         # One-click launcher for Windows
└── README.md                 # This file
```

---

## How to Run

### Option 1: Using the `.bat` File (Easiest — No IDE Required)

1. Navigate to the `Word-to-PDF-Bulk` folder in File Explorer.
2. **Double-click** `Run_Converter.bat`.
3. A terminal window will open and the app will launch in your default browser automatically.
4. To stop the server, simply close the terminal window.

> **Tip:** Right-click `Run_Converter.bat` → **Send to** → **Desktop (create shortcut)** to create a desktop shortcut for quick access.

---

### Option 2: Using VS Code (or any Python IDE)

1. **Open the project folder** in VS Code:
   ```
   File → Open Folder → Select "Word-to-PDF-Bulk"
   ```

2. **Open a terminal** in VS Code:
   - Press `` Ctrl + ` `` (backtick) or go to `Terminal → New Terminal`

3. **Install dependencies** (first time only):
   ```bash
   pip install -r requirements.txt
   ```

4. **Run the application:**
   ```bash
   streamlit run app.py
   ```

5. The app will open in your browser at `http://localhost:8501`.

6. To stop the server, press `Ctrl + C` in the terminal.

---

### Option 3: Using Command Prompt / PowerShell Directly

1. Open **Command Prompt** or **PowerShell**.

2. Navigate to the project folder:
   ```bash
   cd D:\Kapil_Dhole\Word-to-PDF-Bulk
   ```

3. Install dependencies (first time only):
   ```bash
   pip install -r requirements.txt
   ```

4. Run the app:
   ```bash
   streamlit run app.py
   ```

5. Open `http://localhost:8501` in your browser if it doesn't open automatically.

---

## How to Use

1. **Upload** — Drag and drop your `.docx` or `.doc` files into the upload area on the left side.
2. **Preview** — The right side will list all uploaded files with their sizes.
3. **Convert** — Click the **"Convert All to PDF"** button to start conversion.
4. **Download** — Once conversion is complete:
   - Click individual **"Download filename.pdf"** buttons for specific files.
   - Or click **"Download All as ZIP"** to get everything in one archive.

---

## Configuration

### Dark Theme

The dark theme is configured via `.streamlit/config.toml`:

```toml
[theme]
base="dark"
primaryColor="#667eea"
backgroundColor="#0f172a"
secondaryBackgroundColor="#1e293b"
textColor="#e2e8f0"
```

To switch back to a light theme, change `base="dark"` to `base="light"` and restart the server.

### Changing the Port

By default, Streamlit runs on port `8501`. To change it:

```bash
streamlit run app.py --server.port 8080
```

---

## Dependencies

| Package       | Purpose                                          |
|---------------|--------------------------------------------------|
| `streamlit`   | Web application framework for the UI             |
| `docx2pdf`    | Converts `.docx` files to `.pdf` via MS Word COM |
| `python-docx` | Reads Word document metadata                     |

Install all at once:
```bash
pip install -r requirements.txt
```

---

## Production & Deployment Notes

### Local Windows Deployment (Recommended)

This app is optimized for **local Windows environments** where Microsoft Word is installed. For shared team usage:

1. Place the project folder on a shared drive.
2. Each user can run `Run_Converter.bat` from their machine.
3. Ensure Microsoft Word is installed on each user's PC.

### Cloud / Linux Deployment

Since `docx2pdf` depends on Microsoft Word (COM automation), it **cannot run on Linux servers** directly. If you need to deploy on a Linux server (e.g., AWS, GCP, Azure):

1. **Replace `docx2pdf`** with a headless alternative like **LibreOffice**:
   ```bash
   sudo apt-get install libreoffice
   ```
2. Use `subprocess` to call LibreOffice for conversion:
   ```python
   import subprocess
   subprocess.run(["libreoffice", "--headless", "--convert-to", "pdf", input_file])
   ```
3. Update the `convert_word_to_pdf()` function in `app.py` accordingly.

### Running as a Background Service (Windows)

To keep the app running permanently on a Windows machine:

1. Use **NSSM** (Non-Sucking Service Manager) to register it as a Windows service:
   ```bash
   nssm install W2PDF "python" "-m" "streamlit" "run" "app.py"
   nssm set W2PDF AppDirectory "D:\Kapil_Dhole\Word-to-PDF-Bulk"
   nssm start W2PDF
   ```

2. The app will auto-start on system boot and run in the background.

---

## Troubleshooting

| Issue | Solution |
|-------|----------|
| **"docx2pdf failed"** | Ensure Microsoft Word is installed and not running in the background with open dialogs. |
| **App doesn't open in browser** | Manually go to `http://localhost:8501` in your browser. |
| **Port already in use** | Run with a different port: `streamlit run app.py --server.port 8502` |
| **Files download with wrong names** | Clear browser cache and try again. Avoid special characters in filenames. |
| **Dark theme not applying** | Restart the Streamlit server (`Ctrl+C`, then `streamlit run app.py` again). |

---

## License

This project is for internal/personal use.

---

**Built with Streamlit & docx2pdf**
