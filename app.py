#!/usr/bin/env python3
"""
LaTeX to Word Converter Web App
Simple Flask application for converting LaTeX files to Word documents
"""

from flask import Flask, request, render_template, send_file, flash, redirect, url_for
import os
import tempfile
from pathlib import Path
from latex_to_word import LaTeXToWordConverter
import uuid

app = Flask(__name__)
app.secret_key = 'your-secret-key-change-this'

# Configuration
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'outputs'
ALLOWED_EXTENSIONS = {'tex'}
MAX_FILE_SIZE = 10 * 1024 * 1024  # 10MB limit

# Create directories if they don't exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

def allowed_file(filename):
    """Check if file has allowed extension"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    """Main page with upload form"""
    return render_template('index.html')

@app.route('/convert', methods=['POST'])
def convert_file():
    """Handle file upload and conversion"""
    if 'file' not in request.files:
        flash('No file selected')
        return redirect(url_for('index'))
    
    file = request.files['file']
    
    if file.filename == '':
        flash('No file selected')
        return redirect(url_for('index'))
    
    if file and allowed_file(file.filename):
        try:
            # Check file size
            file.seek(0, 2)  # Seek to end
            file_size = file.tell()
            file.seek(0)  # Reset to beginning
            
            if file_size > MAX_FILE_SIZE:
                flash(f'File too large. Maximum size is {MAX_FILE_SIZE // (1024*1024)}MB.')
                return redirect(url_for('index'))
            
            # Generate unique filename
            unique_id = str(uuid.uuid4())
            tex_filename = f"{unique_id}.tex"
            docx_filename = f"{unique_id}.docx"
            
            # Save uploaded file
            tex_path = os.path.join(UPLOAD_FOLDER, tex_filename)
            file.save(tex_path)
            
            print(f"📁 Saved uploaded file: {tex_path}")
            
            # Copy photo to upload directory if it exists
            photo_source = "presidency photo.png"
            photo_dest = os.path.join(UPLOAD_FOLDER, "presidency photo.png")
            if os.path.exists(photo_source) and not os.path.exists(photo_dest):
                import shutil
                shutil.copy2(photo_source, photo_dest)
                print(f"📸 Copied photo to upload directory: {photo_dest}")
            
            # Convert to Word
            output_path = os.path.join(OUTPUT_FOLDER, docx_filename)
            converter = LaTeXToWordConverter(tex_path, output_path)
            
            print(f"🔄 Starting conversion: {tex_path} → {output_path}")
            success = converter.convert()
            
            if success:
                print(f"✅ Conversion successful: {output_path}")
                # Clean up uploaded file
                os.remove(tex_path)
                
                # Send converted file to user
                return send_file(
                    output_path,
                    as_attachment=True,
                    download_name=f"{Path(file.filename).stem}.docx",
                    mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
                )
            else:
                print(f"❌ Conversion failed for: {tex_path}")
                # Clean up uploaded file
                if os.path.exists(tex_path):
                    os.remove(tex_path)
                flash('Conversion failed. Please check your LaTeX file.')
                return redirect(url_for('index'))
                
        except Exception as e:
            print(f"💥 Exception during conversion: {str(e)}")
            # Clean up uploaded file
            if 'tex_path' in locals() and os.path.exists(tex_path):
                os.remove(tex_path)
            flash(f'Error during conversion: {str(e)}')
            return redirect(url_for('index'))
    
    else:
        flash('Invalid file type. Please upload a .tex file.')
        return redirect(url_for('index'))

@app.route('/cleanup')
def cleanup():
    """Clean up old files"""
    try:
        import time
        current_time = time.time()
        cleaned_count = 0
        
        # Clean up files older than 1 hour
        for folder in [UPLOAD_FOLDER, OUTPUT_FOLDER]:
            for filename in os.listdir(folder):
                filepath = os.path.join(folder, filename)
                if os.path.isfile(filepath):
                    file_age = current_time - os.path.getmtime(filepath)
                    if file_age > 3600:  # 1 hour
                        os.remove(filepath)
                        cleaned_count += 1
                        print(f"🗑️ Cleaned up old file: {filepath}")
        
        return {'status': 'success', 'message': f'Cleaned up {cleaned_count} old files'}
    except Exception as e:
        return {'status': 'error', 'message': str(e)}

@app.route('/status')
def status():
    """System status endpoint"""
    try:
        upload_count = len([f for f in os.listdir(UPLOAD_FOLDER) if f.endswith('.tex')])
        output_count = len([f for f in os.listdir(OUTPUT_FOLDER) if f.endswith('.docx')])
        
        return {
            'status': 'healthy',
            'message': 'LaTeX to Word Converter is running',
            'files': {
                'uploads': upload_count,
                'outputs': output_count
            },
            'limits': {
                'max_file_size_mb': MAX_FILE_SIZE // (1024 * 1024)
            }
        }
    except Exception as e:
        return {'status': 'error', 'message': str(e)}

@app.route('/health')
def health():
    """Health check endpoint"""
    return {'status': 'healthy', 'message': 'LaTeX to Word Converter is running'}

if __name__ == '__main__':
    print("🚀 Starting LaTeX to Word Converter Web App...")
    print("📁 Upload folder:", UPLOAD_FOLDER)
    print("📁 Output folder:", OUTPUT_FOLDER)
    
    # Get port from environment variable (Railway provides this)
    port = int(os.environ.get('PORT', 5000))
    debug = os.environ.get('FLASK_ENV') == 'development'
    
    print(f"🌐 Railway PORT environment variable: {os.environ.get('PORT', 'NOT SET')}")
    print(f"🌐 Starting server on port {port}")
    print(f"🌐 Debug mode: {debug}")
    app.run(debug=debug, host='0.0.0.0', port=port)
