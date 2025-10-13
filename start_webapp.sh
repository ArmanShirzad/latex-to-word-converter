#!/bin/bash
# LaTeX to Word Converter - Web App Startup Script

echo "🚀 Starting LaTeX to Word Converter Web App..."
echo ""

# Check if virtual environment exists
if [ ! -d "venv" ]; then
    echo "❌ Virtual environment not found. Please run setup first."
    exit 1
fi

# Activate virtual environment
echo "📦 Activating virtual environment..."
source venv/bin/activate

# Check if Flask is installed
if ! python3 -c "import flask" 2>/dev/null; then
    echo "📥 Installing Flask..."
    pip install Flask==3.0.0
fi

# Create necessary directories
echo "📁 Creating directories..."
mkdir -p uploads outputs templates static/css

# Start the web app
echo ""
echo "🌐 Starting web server..."
echo "📱 Open your browser and go to: http://localhost:5000"
echo "🛑 Press Ctrl+C to stop the server"
echo ""

python3 app.py
