Here is the `README.md` file for your PyQt5-based PDF to DOCX converter project:

```markdown
# PDF to DOCX Converter (PyQt5)

A user-friendly GUI application to convert multiple PDF files into DOCX format using Adobe PDF Services API. The application supports **multi-threading**, **real-time progress tracking**, and **batch processing**.

---

## Features

✅ **Multi-PDF Selection** – Users can select multiple PDFs at once.  
✅ **Multithreading** – Each PDF conversion runs in a separate thread.  
✅ **Progress Tracking** – The UI shows a progress bar for each file.  
✅ **Auto-Saving** – Converted DOCX files are saved in the same directory as the PDFs.  
✅ **Start & Stop Control** – "Start" to begin conversion and "Stop" to terminate all processes.  
✅ **Modern UI** – Built with PyQt5, featuring a sleek and modern interface.

---

## Installation

### Prerequisites
- **Python 3.7+**
- **PyQt5**
- **Requests** (for API communication)

### Install Required Packages
```bash
pip install PyQt5 requests
```

---

## How to Use

1. **Run the Application**
   ```bash
   python gui.py
   ```

2. **Select PDF Files**  
   - Click **"Select PDFs"** to choose multiple PDF files.
   - The selected files will be listed in the UI.

3. **Start Conversion**  
   - Click **"Start Conversion"** to begin the process.
   - The progress bar updates in real-time for each file.

4. **Stop All Processes**  
   - Click **"Stop All"** to terminate ongoing conversions.

5. **View Results**  
   - Converted DOCX files will be saved in the **same directory** as the original PDFs.

---

## API Configuration

The application requires **Adobe PDF Services API credentials** for conversion.  

### Steps:
1. **Create an API Key**  
   - Go to [Adobe PDF Services API](https://developer.adobe.com/document-services/apis/pdf-services/)
   - Register and obtain **client credentials**.

2. **Save Credentials**  
   - Download the `pdfservices-api-credentials.json` file.
   - Place the file in a secure location on your system.

3. **Load Credentials in the App**  
   - Click **"🔑 Choose Key"** and select your `pdfservices-api-credentials.json` file.

---

## Project Structure

```
📂 pdf-to-docx-gui
 ┣ 📜 gui.py              # Main GUI application
 ┣ 📜 README.md           # Documentation
 ┣ 📜 requirements.txt     # List of dependencies
 ┗ 📜 pdfservices-api-credentials.json  # Adobe API credentials (user-provided)
```

---

## Screenshots

![image](https://github.com/user-attachments/assets/3ae02dd6-0c6e-4540-a542-2c2de34d4231)

Live:
![image](https://github.com/user-attachments/assets/e167a17f-d765-4726-8eae-1258e3ab3443)



---

## Troubleshooting

### 1. "API Authentication Failed"
- Ensure that `pdfservices-api-credentials.json` is valid.
- Make sure your API key has the correct permissions.

### 2. "Conversion Failed"  
- Check internet connectivity.
- Try again later in case of Adobe API downtime.

### 3. "Progress Bar Stuck"
- The process might be slow due to a large file.
- Try restarting the application.

---

## License
This project is **open-source** and available for modification.

---

## Author
Developed by **[Your Name]** 🚀
```

### Let me know if you want to customize anything! 🚀
