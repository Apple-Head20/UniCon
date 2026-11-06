# UniCon — Universal File Converter

### A Windows Metro-Inspired Multi-Format File Conversion Application

**UniCon** is a Python-based desktop application designed to simplify file conversion through a unified graphical interface. Inspired by the Windows 8 Metro design language, UniCon combines animated tiles, a dark-themed interface, drag-and-drop functionality, and a multi-engine conversion architecture.

The application provides a centralized workspace for converting documents, images, videos, audio files, archives, ebooks, structured data, and fonts.

Rather than implementing every conversion algorithm independently, UniCon integrates established tools such as **FFmpeg, LibreOffice, Pillow, Calibre, and Python's built-in libraries**.

---

## Features

- **Universal Conversion Interface** — Manage multiple file categories from a single application.
- **Windows Metro-Inspired UI** — Dark theme, colorful category tiles, animated backgrounds, and hover effects.
- **Batch File Conversion** — Queue multiple files and process them concurrently.
- **Drag-and-Drop Support** — Add files directly to the conversion queue when `tkinterdnd2` is available.
- **Automatic Format Suggestions** — Suggest output formats based on detected input categories.
- **Multi-Engine Architecture** — Routes files to different conversion engines based on their extensions.
- **Conversion Progress Tracking** — Animated progress bar with per-file status indicators.
- **Custom Output Directory** — Choose where converted files are saved.
- **Automatic Folder Opening** — Optionally open the destination folder after conversion.
- **Dependency Detection** — Displays availability indicators for several external conversion tools.
- **Queue Management** — Add, remove, and clear files before conversion.

## Supported File Categories

UniCon defines eight file categories.

| Category | Recognized input formats | Conversion engine |
|---|---|---|
| Documents | DOCX, DOC, ODT, RTF, TXT, PPTX, PPT, ODP, XLSX, XLS, ODS, CSV | LibreOffice |
| Images | JPG, JPEG, PNG, WEBP, BMP, TIFF, GIF, ICO, HEIC, HEIF, AVIF, SVG, PSD | Pillow |
| Video | MP4, MKV, AVI, MOV, WMV, FLV, WEBM, M4V, 3GP, TS, VOB, OGV | FFmpeg |
| Audio | MP3, WAV, FLAC, AAC, OGG, M4A, WMA, OPUS, AIFF, ALAC | FFmpeg |
| Archives | ZIP, TAR, GZ, BZ2, XZ, 7Z, RAR | Python archive libraries / 7-Zip |
| Ebooks | EPUB, MOBI, AZW3, FB2, PDF | Calibre |
| Data | CSV, JSON, XML, YAML, YML, TSV, XLSX, SQL | Python data conversion |
| Fonts | TTF, OTF, WOFF, WOFF2 | fontTools |

**Important:** These are formats recognized by the application's conversion map, not a guarantee that every input-to-output combination works. Actual conversion support depends on the selected engine, installed dependencies, codecs, and file contents.

## Technology Stack

| Technology | Purpose |
|---|---|
| Python | Core application and conversion orchestration |
| CustomTkinter | Modern desktop GUI |
| Tkinter | Native desktop components and file dialogs |
| tkinterdnd2 | Drag-and-drop functionality |
| Pillow | Image processing and conversion |
| pillow-heif | HEIF/HEIC image support |
| FFmpeg | Audio and video conversion |
| LibreOffice | Office document conversion |
| Calibre | Ebook conversion |
| PyYAML | YAML parsing and serialization |
| fontTools | Font file processing |
| concurrent.futures | Concurrent conversion tasks |
| threading | Background processing |

## Application Architecture

UniCon separates the user interface from its conversion logic.

```text
                   UniCon Desktop GUI
                           |
              File Selection / Drag & Drop
                           |
                      File Queue
                           |
                 Extension Detection
                           |
                  Category Resolution
                           |
                  Conversion Dispatcher
                           |
          +----------------+----------------+
          |                |                |
     LibreOffice         Pillow           FFmpeg
      Documents          Images         Video/Audio
          |                |                |
          +----------------+----------------+
                           |
             Other Conversion Engines
                           |
             Archives / Calibre / Data
                       / Fonts
                           |
                  Output Directory
                           |
                Progress and Status
```

### Conversion Workflow

1. **File selection:** Users browse for files or drag them into the application.
2. **Format detection:** UniCon identifies the category using the input file extension.
3. **Output selection:** The interface presents suggested destination formats.
4. **Destination selection:** Users choose the output directory.
5. **Engine dispatch:** Each file is passed to the conversion engine associated with its category.
6. **Background execution:** A thread pool processes queued conversions.
7. **Status reporting:** The interface updates individual file statuses and overall progress.
8. **Completion:** A summary is displayed, and the destination folder can be opened automatically.

## Installation

### Prerequisites

- Python 3 with compatible package versions
- Windows (primary intended platform)
- Git (optional, for cloning)
- External conversion utilities for the relevant file categories

### 1. Clone the Repository

```bash
git clone https://github.com/Apple-Head20/UniCon.git
cd UniCon
```

### 2. Create a Virtual Environment

On Windows PowerShell:

```powershell
python -m venv venv
```

Activate it:

```powershell
.\venv\Scripts\Activate.ps1
```

### 3. Install Python Dependencies

```powershell
python -m pip install -r requirements.txt
```

If installing dependencies manually:

```powershell
python -m pip install customtkinter tkinterdnd2 Pillow pillow-heif PyYAML fonttools
```

### 4. Install External Conversion Tools

Depending on the file types you intend to convert, install the following tools:

**FFmpeg — Audio and Video**

https://ffmpeg.org/download.html

Ensure the `ffmpeg` executable is available in your system PATH.

Verify:

```powershell
ffmpeg -version
```

**LibreOffice — Documents**

https://www.libreoffice.org/download/download-libreoffice/

UniCon checks standard Windows installation locations for `soffice.exe`.

**Calibre — Ebooks**

https://calibre-ebook.com/download

Ensure `ebook-convert` is accessible through PATH.

**7-Zip — Archives**

https://www.7-zip.org/

Ensure the `7z` command is available in PATH for archive formats that require it.

### 5. Run UniCon

```powershell
python unicon.py
```

The application opens its graphical interface with conversion categories, a file queue, and conversion controls.

## How to Use

### Step 1 — Add Files

Click **Browse Files** to select files from your computer.

Alternatively, drag files into the queue if drag-and-drop support is available.

### Step 2 — Select a Category or Output Format

Choose one of the category tiles:

- Documents
- Images
- Video
- Audio
- Archives
- Ebooks
- Data
- Fonts

Selecting a category updates the available output format options.

### Step 3 — Choose a Destination

Click **Choose Folder** and select the directory where converted files should be saved.

### Step 4 — Start Conversion

Click **Start Conversion**.

UniCon processes the queued files and displays their progress.

### Step 5 — Access Converted Files

Once conversion finishes, the application displays the number of successful conversions.

If **Auto-open folder** is enabled, the destination directory opens automatically.

---

## Project Structure

```text
UniCon/
|
|-- unicon.py
|-- requirements.txt
|-- .gitignore
|-- README.md
```

### File Descriptions

- `unicon.py` — Main application, graphical interface, conversion engines, queue management, and progress tracking.
- `requirements.txt` — Python dependencies.
- `.gitignore` — Git exclusion rules.
- `README.md` — Project documentation.

## Conversion Engines

### LibreOffice

Handles office document conversion using the LibreOffice command-line interface in headless mode.

A temporary LibreOffice user profile is created for each conversion process.

### Pillow

Handles image loading, image mode adjustments, and saving to supported output formats.

HEIC/HEIF support is enhanced through `pillow-heif` when installed.

### FFmpeg

Handles video transcoding, audio conversion, audio extraction, and GIF generation.

The implementation includes predefined encoding options for selected output formats, including MP4 and WEBM.

### Archive Engine

Uses Python's `zipfile` and `tarfile` modules to extract and repackage supported archives.

External 7-Zip commands are used for some other archive formats.

### Calibre

Uses the `ebook-convert` command-line utility for ebook format conversion.

### Data Conversion Engine

Uses Python's JSON, CSV, and XML libraries, with optional YAML support through PyYAML.

### Font Engine

Uses fontTools to load and save font files.

Font conversion compatibility depends on the source and destination formats and is not guaranteed for every combination.

## Current Limitations

UniCon is an experimental desktop application. Several aspects require further development and validation.

1. **External dependencies:** Document, video, audio, and ebook conversions require additional installed tools.
2. **Format compatibility:** Not every combination shown in the interface is supported by the underlying conversion engine.
3. **Extension-based detection:** Files are classified by their extensions rather than inspecting their actual contents.
4. **Mixed-category queues:** A single selected output format may not be appropriate for every file in a mixed queue.
5. **Data conversion:** Complex structures may not retain their original schema, metadata, or formatting.
6. **Font conversion:** Loading and saving a font under a different extension does not necessarily produce a valid converted font format.
7. **Archive security:** Archive extraction requires additional safeguards against malicious or untrusted archive contents.
8. **Output collisions:** Existing output files may be overwritten by some conversion engines.
9. **Platform compatibility:** The application is primarily designed for Windows; other operating systems have not been fully validated.
10. **Dependency detection:** Some optional dependency checks require refinement.
11. **Conversion reporting:** Success reporting is largely based on engine execution rather than comprehensive output-file validation.

## Planned Improvements

Potential future improvements include:

- More accurate file-format detection
- Compatibility checks before conversion
- Per-file output format selection
- Conversion presets and quality settings
- More robust error reporting
- Improved handling of unsupported formats
- Output collision detection
- Secure archive extraction
- Conversion history
- More reliable dependency detection
- Automated conversion tests
- Standalone Windows executable packaging
- Improved cross-platform compatibility

These are development goals rather than features currently guaranteed by the application.

## Security and Privacy

UniCon performs conversions locally through Python libraries and installed command-line utilities.

The current implementation does not include a cloud upload or remote conversion service.

However, users should exercise caution when converting files from untrusted sources, especially archives, complex documents, and media files processed by external tools.

## Author

**Mridul Mayank Jha**

B.Tech — Computer Science and Engineering  
Sikkim Manipal Institute of Technology

GitHub: [Apple-Head20](https://github.com/Apple-Head20)

Repository: [UniCon](https://github.com/Apple-Head20/UniCon)

---

**UniCon — One interface. Multiple formats. Simplified conversion.**
