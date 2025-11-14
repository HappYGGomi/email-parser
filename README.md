# 📧 Email Parser

Parse `.eml` and `.msg` email files, extract Q&A pairs from email threads, and export to Excel.

[![Python](https://img.shields.io/badge/Python-3.7+-blue.svg)](https://www.python.org/)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

## ✨ Features

- 📥 **Multi-format Support**: Parse both `.eml` and `.msg` (Microsoft Outlook) files
- 🔄 **Thread Detection**: Automatically organize emails into conversation threads
- 💬 **Q&A Extraction**: Extract question-answer pairs from email conversations
- 📊 **Excel Export**: Export to beautifully formatted Excel files
- 🖥️ **Dual Interface**: Both GUI (tkinter) and CLI available
- 📁 **Batch Processing**: Process multiple files or entire directories at once

## 🚀 Quick Start

### Installation

```bash
# Clone the repository
git clone https://github.com/HappYGGomi/email-parser.git
cd email-parser

# Switch to email-parser branch
git checkout email-parser

# Install dependencies (Windows)
install.bat
```

Or install manually:
```bash
pip install openpyxl extract-msg
```

### Usage

#### GUI (Recommended)

```bash
python gui.py
```

Or double-click `run_gui.bat` on Windows.

#### CLI

```bash
# Parse a single file
python cli.py -f email.eml

# Parse multiple files (mixed formats)
python cli.py -f email1.eml email2.msg email3.eml

# Parse directory (recursive)
python cli.py -d ./emails

# Specify output file
python cli.py -d ./emails -o output.xlsx
```

## 📖 Documentation

- [Quick Start Guide](QUICK_START.md) - Get started in 1 minute
- [Full Documentation](EMAIL_PARSER_README.md) - Detailed usage instructions
- [Examples](example_usage.py) - Code examples

## 🎯 Use Cases

- **Customer Support**: Organize support emails into Q&A pairs
- **Email Archiving**: Convert email threads to structured Excel format
- **Data Analysis**: Extract and analyze email conversations
- **Documentation**: Create FAQ documents from support emails

## 📋 Requirements

- Python 3.7+
- openpyxl >= 3.1.2
- extract-msg >= 0.41.0

See [requirements_email_parser.txt](requirements_email_parser.txt) for full dependencies.

## 📂 Project Structure

```
email-parser/
├── email_parser.py          # Core parsing engine
├── excel_exporter.py        # Excel export functionality
├── cli.py                   # Command-line interface
├── gui.py                   # Graphical user interface
├── install.bat              # Windows installation script
├── run_gui.bat             # Windows GUI launcher
├── requirements_email_parser.txt  # Python dependencies
├── EMAIL_PARSER_README.md  # Detailed documentation
├── QUICK_START.md          # Quick start guide
└── example_usage.py        # Usage examples
```

## 🔧 How It Works

1. **Parse**: Reads `.eml` or `.msg` files and extracts headers, body, and metadata
2. **Thread**: Groups emails by subject and organizes them chronologically
3. **Extract**: Identifies Q&A pairs based on email sequence
4. **Export**: Creates formatted Excel files with proper styling

## 🤝 Contributing

Contributions are welcome! Feel free to:
- Report bugs
- Suggest features
- Submit pull requests

## 📝 License

This project is open source and available under the MIT License.

## 🙏 Acknowledgments

Built with:
- [openpyxl](https://openpyxl.readthedocs.io/) - Excel file handling
- [extract-msg](https://github.com/TeamMsgExtractor/msg-extractor) - MSG file parsing
- Python's built-in `email` library - EML file parsing

---

**Made with ❤️ by HappYGGomi**

🤖 *Generated with [Claude Code](https://claude.com/claude-code)*
