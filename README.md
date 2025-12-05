# smart-file-organizer
A smart desktop file organizer built with Python and PySide6

# Smart File Organizer  
A desktop application that automatically organizes files using intelligent rules, metadata extraction, duplicate detection, and real-time folder monitoring.  
Built with Python, PySide6, Watchdog, and a modern GitHub-style UI.

---

## 🚀 Features

### 🔹 Smart Auto-Organization  
Sorts files based on:
- File type  
- EXIF metadata  
- PDF/DOCX metadata  
- Filename patterns  
- Content-based rules  
- Date modified  
- File size  
- Custom user-defined rules  

### 🔹 Real-Time Folder Watching  
Automatically organizes new files as soon as they appear.

### 🔹 Duplicate File Scanner  
Detects duplicate files using MD5 hashing.

### 🔹 Undo System (De-Organize)  
Restores files to their original paths using JSON logs.

### 🔹 Backup Before Organizing  
Creates ZIP backups of the entire folder before processing.

### 🔹 Modern UI  
- GitHub Light/Dark theme  
- Frameless window  
- Smooth animations  
- Toolbar shortcuts  
- Progress bar & logs  
- Drag-and-drop support  

---

## 🛠️ Technologies Used

- **Python**
- **PySide6 / Qt**
- **Watchdog**
- **PIL (Pillow)**
- **PyPDF2**
- **python-docx**
- **JSON & ZipFile**

---

## 📦 Installation

1. Clone the repository:
```bash
git clone https://github.com/yourusername/smart-file-organizer.git
