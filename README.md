# LaTeX to Word Converter

A modern web application that converts LaTeX files to Word documents with high fidelity formatting preservation and automatic image embedding.

![LaTeX to Word Converter](https://img.shields.io/badge/LaTeX-Word%20Converter-blue?style=for-the-badge&logo=latex)
![Python](https://img.shields.io/badge/Python-3.12+-green?style=for-the-badge&logo=python)
![Flask](https://img.shields.io/badge/Flask-3.0+-red?style=for-the-badge&logo=flask)
![License](https://img.shields.io/badge/License-MIT-yellow?style=for-the-badge)

## Features

- **Simple Upload Interface**: Drag & drop or click to upload `.tex` files
- **High Fidelity Conversion**: Preserves LaTeX formatting, structure, and styling
- **Image Embedding**: Automatically embeds images from LaTeX documents
- **Modern UI**: Clean, responsive design with smooth animations
- **File Management**: Automatic cleanup of temporary files
- **Security**: File size limits and type validation
- **Error Handling**: User-friendly error messages
- **Cross-platform**: Works on Windows, macOS, and Linux

## Live Demo

Visit the live application: [https://latex-to-word-converter-production.up.railway.app](https://latex-to-word-converter-production.up.railway.app)

## Screenshots

The application features a modern gradient background with a clean white upload card, drag-and-drop functionality, and real-time conversion status updates.

## Installation

### Prerequisites

- Python 3.12+
- Pandoc
- pip

### Setup

1. **Clone the repository**
   ```bash
   git clone https://github.com/yourusername/latex-to-word-converter.git
   cd latex-to-word-converter
   ```

2. **Create virtual environment**
   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```

3. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

4. **Run the application**
   ```bash
   python app.py
   ```

5. **Open your browser**
   Go to `http://localhost:5000`

## Usage

1. **Upload**: Select or drag your `.tex` file to the upload area
2. **Convert**: Click "Convert to Word" button
3. **Download**: Your `.docx` file will download automatically

## Project Structure

```
latex-to-word-converter/
├── app.py                 # Flask web application
├── latex_to_word.py       # Core conversion logic
├── templates/
│   └── index.html        # Main UI template
├── static/css/
│   └── style.css         # Modern styling
├── uploads/              # Temporary upload folder
├── outputs/              # Temporary output folder
├── requirements.txt      # Python dependencies
├── start_webapp.sh       # Startup script
└── README.md            # This file
```

## Technical Details

- **Backend**: Flask (Python)
- **Frontend**: HTML5, CSS3, JavaScript
- **Conversion**: Pandoc + python-docx
- **Styling**: Modern CSS with gradients and animations
- **File Handling**: Secure upload/download with UUID naming

## API Endpoints

- `GET /` - Main upload page
- `POST /convert` - File conversion endpoint
- `GET /health` - Health check
- `GET /status` - System status
- `GET /cleanup` - Manual cleanup

## Security Features

- File type validation (only `.tex` files)
- File size limits (10MB maximum)
- UUID-based file naming
- Automatic cleanup of temporary files
- Error handling and validation

## Deployment

The app is ready for deployment on any platform that supports Python/Flask:

- **Live Demo**: [https://latex-to-word-converter-production.up.railway.app](https://latex-to-word-converter-production.up.railway.app)
- **Local Development**: `python app.py`
- **Railway**: Deployed with Dockerfile including Pandoc installation
- **Production**: Use gunicorn or similar WSGI server
- **Docker**: Easy to containerize
- **Cloud**: Deploy to Heroku, AWS, Google Cloud, etc.

## Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Authors

**Arman Shirzad**
- GitHub: [@ArmanShirzad](https://github.com/ArmanShirzad)
- LinkedIn: [arman-shirzad](https://linkedin.com/in/arman-shirzad)
- Website: [armanshirzad.guru](https://armanshirzad.guru)

**Mahdi Ahmadi**
- Email: [mahdi73.ahmadi@gmail.com](mailto:mahdi73.ahmadi@gmail.com)

## Acknowledgments

- [Pandoc](https://pandoc.org/) for LaTeX to Word conversion
- [python-docx](https://python-docx.readthedocs.io/) for Word document manipulation
- [Flask](https://flask.palletsprojects.com/) for the web framework

## Roadmap

See [FEATURES.md](FEATURES.md) for planned features and improvements.

---

**Star this repository if you found it helpful!**