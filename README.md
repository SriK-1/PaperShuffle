# PaperShuffle
A desktop application designed to automate the generation of randomized exam question papers from Word, Excel, PDF, and text files. 

WORK IN PROGRESS - This project is currently in active development and testing: the initital source code is uploaded, but there are some additional test cases to be run before I upload the packaged executable.

## Features

- **Multi-Format Support**: Import questions from `.docx`, `.xlsx`, `.pdf`, and `.txt` files
- **AI-Powered Parsing**: Integrated with Groq's Llama-3 70B for intelligent question separation
- **Human-in-the-Loop**: Preview and manually correct parsed questions before generation
- **Smart Answer Detection**: Automatically highlights answer keys; optional student version strips answers
- **Flexible Output**: Generate multiple files or one combined document with page breaks
- **Modern UI**: Responsive dark-mode interface with progress tracking
  
## Installation

```bash
pip install customtkinter python-docx openpyxl PyPDF2 groq
