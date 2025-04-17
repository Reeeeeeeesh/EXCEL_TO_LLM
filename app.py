import os
from flask import Flask, request, render_template, flash, redirect, url_for, send_file
from werkzeug.utils import secure_filename
from excel_to_llm_converter import ExcelToLLMConverter
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

app = Flask(__name__)
app.config['SECRET_KEY'] = 'your-secret-key-here'  # Change this to a secure secret key
app.config['UPLOAD_FOLDER'] = os.path.join(os.getcwd(), 'uploads')
app.config['OUTPUT_ROOT'] = os.path.join(os.getcwd(), 'output')
app.config['GOOGLE_API_KEY'] = os.getenv('GOOGLE_API_KEY')
if not app.config['GOOGLE_API_KEY']:
    raise ValueError("GOOGLE_API_KEY environment variable is not set")

# Ensure upload and output directories exist
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['OUTPUT_ROOT'], exist_ok=True)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() == 'xlsx'

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        if 'file' not in request.files:
            flash('No file part', 'error')
            return redirect(request.url)
        
        file = request.files['file']
        output_directory = request.form.get('output_directory', '').strip()
        
        if file.filename == '':
            flash('No selected file', 'error')
            return redirect(request.url)
        
        if not output_directory:
            flash('Output directory name is required', 'error')
            return redirect(request.url)
        
        if file and allowed_file(file.filename):
            try:
                # Save uploaded file
                filename = secure_filename(file.filename)
                filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                file.save(filepath)
                
                # Process Excel file
                output_path = os.path.join(app.config['OUTPUT_ROOT'], output_directory)
                os.makedirs(output_path, exist_ok=True)
                
                try:
                    # Initialize converter with input path, output directory, and API key
                    converter = ExcelToLLMConverter(
                        input_path=filepath,
                        output_dir=output_path,
                        api_key=app.config['GOOGLE_API_KEY']
                    )
                    converter.convert_all()
                    
                    # Get the Excel filename without extension
                    excel_filename = os.path.splitext(os.path.basename(filepath))[0]
                    
                    flash('File successfully processed', 'success')
                    return redirect(url_for('upload_file'))
                    
                except Exception as e:
                    flash(f'Error processing file: {str(e)}', 'error')
                    return redirect(request.url)
                
            except Exception as e:
                flash(f'Error saving file: {str(e)}', 'error')
                return redirect(request.url)
        else:
            flash('Invalid file type. Please upload an Excel file (.xlsx)', 'error')
            return redirect(request.url)
            
    return render_template('upload.html')

@app.route('/download/<path:filename>')
def download_file(filename):
    try:
        return send_file(os.path.join(app.config['OUTPUT_ROOT'], filename),
                        as_attachment=True)
    except Exception as e:
        flash(f'Error downloading file: {str(e)}', 'error')
        return redirect(url_for('upload_file'))

if __name__ == '__main__':
    app.run(debug=True)
