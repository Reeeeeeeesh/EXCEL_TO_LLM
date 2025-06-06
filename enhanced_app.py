import os
from flask import Flask, request, render_template, flash, redirect, url_for, send_file, jsonify
from werkzeug.utils import secure_filename
from enhanced_excel_converter import EnhancedExcelConverter
from dotenv import load_dotenv
import json
from pathlib import Path

# Load environment variables
load_dotenv()

app = Flask(__name__)
app.config['SECRET_KEY'] = 'your-secret-key-here'  # Change this to a secure secret key
app.config['UPLOAD_FOLDER'] = os.path.join(os.getcwd(), 'uploads')
app.config['OUTPUT_ROOT'] = os.path.join(os.getcwd(), 'OUTPUT')
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
        generate_prd = request.form.get('generate_prd') == 'on'
        
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
                
                # Process Excel file with enhanced converter
                output_path = os.path.join(app.config['OUTPUT_ROOT'], output_directory)
                os.makedirs(output_path, exist_ok=True)
                
                try:
                    # Initialize enhanced converter
                    converter = EnhancedExcelConverter(
                        input_path=filepath,
                        output_dir=output_path,
                        api_key=app.config['GOOGLE_API_KEY'],
                        generate_prd=generate_prd
                    )
                    converter.convert_all()
                    
                    # Get processing results for display
                    results = get_processing_results(output_path)
                    
                    success_message = 'File successfully processed!'
                    if generate_prd:
                        success_message += ' PRD document generated.'
                    
                    flash(success_message, 'success')
                    return render_template('results.html', results=results, output_dir=output_directory)
                    
                except Exception as e:
                    flash(f'Error processing file: {str(e)}', 'error')
                    return redirect(request.url)
                
            except Exception as e:
                flash(f'Error saving file: {str(e)}', 'error')
                return redirect(request.url)
        else:
            flash('Invalid file type. Please upload an Excel file (.xlsx)', 'error')
            return redirect(request.url)
            
    return render_template('enhanced_upload.html')

@app.route('/results/<output_dir>')
def view_results(output_dir):
    """View processing results for a specific output directory."""
    output_path = os.path.join(app.config['OUTPUT_ROOT'], output_dir)
    if not os.path.exists(output_path):
        flash('Output directory not found', 'error')
        return redirect(url_for('upload_file'))
    
    results = get_processing_results(output_path)
    return render_template('results.html', results=results, output_dir=output_dir)

@app.route('/api/analysis/<output_dir>')
def get_analysis_api(output_dir):
    """API endpoint to get analysis data in JSON format."""
    output_path = os.path.join(app.config['OUTPUT_ROOT'], output_dir)
    if not os.path.exists(output_path):
        return jsonify({'error': 'Output directory not found'}), 404
    
    results = get_processing_results(output_path)
    return jsonify(results)

@app.route('/download/<path:filename>')
def download_file(filename):
    """Download generated files."""
    try:
        file_path = os.path.join(app.config['OUTPUT_ROOT'], filename)
        if not os.path.exists(file_path):
            flash('File not found', 'error')
            return redirect(url_for('upload_file'))
        
        return send_file(file_path, as_attachment=True)
    except Exception as e:
        flash(f'Error downloading file: {str(e)}', 'error')
        return redirect(url_for('upload_file'))

@app.route('/preview/<path:filename>')
def preview_file(filename):
    """Preview markdown files in the browser."""
    try:
        file_path = os.path.join(app.config['OUTPUT_ROOT'], filename)
        if not os.path.exists(file_path):
            flash('File not found', 'error')
            return redirect(url_for('upload_file'))
        
        # Read file content
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Determine file type for proper rendering
        file_ext = Path(filename).suffix.lower()
        
        return render_template('preview.html', 
                             content=content, 
                             filename=filename, 
                             file_type=file_ext)
    except Exception as e:
        flash(f'Error previewing file: {str(e)}', 'error')
        return redirect(url_for('upload_file'))

@app.route('/compare/<output_dir>')
def compare_documents(output_dir):
    """Compare generated documents (user guide vs PRD)."""
    output_path = os.path.join(app.config['OUTPUT_ROOT'], output_dir)
    if not os.path.exists(output_path):
        flash('Output directory not found', 'error')
        return redirect(url_for('upload_file'))
    
    # Look for specific files to compare
    workbook_subdir = None
    for item in os.listdir(output_path):
        item_path = os.path.join(output_path, item)
        if os.path.isdir(item_path):
            workbook_subdir = item_path
            break
    
    if not workbook_subdir:
        flash('No workbook directory found', 'error')
        return redirect(url_for('upload_file'))
    
    # Read comparison files
    user_guide_path = os.path.join(workbook_subdir, 'llm_analysis_report.md')
    prd_path = os.path.join(workbook_subdir, 'software_prd.md')
    
    user_guide_content = ""
    prd_content = ""
    
    if os.path.exists(user_guide_path):
        with open(user_guide_path, 'r', encoding='utf-8') as f:
            user_guide_content = f.read()
    
    if os.path.exists(prd_path):
        with open(prd_path, 'r', encoding='utf-8') as f:
            prd_content = f.read()
    
    return render_template('compare.html', 
                         user_guide=user_guide_content,
                         prd=prd_content,
                         output_dir=output_dir)

def get_processing_results(output_path):
    """Extract processing results and metadata from output directory."""
    results = {
        'summary': {},
        'files': [],
        'workbooks': [],
        'has_prd': False,
        'has_user_guide': False,
        'complexity_metrics': {}
    }
    
    try:
        # Look for workbook directories
        for item in os.listdir(output_path):
            item_path = os.path.join(output_path, item)
            if os.path.isdir(item_path):
                workbook_data = analyze_workbook_directory(item_path, item)
                results['workbooks'].append(workbook_data)
                
                # Check for main documents
                if os.path.exists(os.path.join(item_path, 'software_prd.md')):
                    results['has_prd'] = True
                if os.path.exists(os.path.join(item_path, 'llm_analysis_report.md')):
                    results['has_user_guide'] = True
        
        # List all files in output directory
        for root, dirs, files in os.walk(output_path):
            for file in files:
                file_path = os.path.join(root, file)
                relative_path = os.path.relpath(file_path, output_path)
                file_size = os.path.getsize(file_path)
                
                results['files'].append({
                    'name': file,
                    'path': relative_path,
                    'size': format_file_size(file_size),
                    'type': get_file_type(file),
                    'can_preview': can_preview_file(file)
                })
        
        # Generate overall summary
        results['summary'] = generate_overall_summary(results)
        
    except Exception as e:
        print(f"Error analyzing results: {str(e)}")
    
    return results

def analyze_workbook_directory(workbook_path, workbook_name):
    """Analyze a specific workbook directory for metadata."""
    workbook_data = {
        'name': workbook_name,
        'sheets': [],
        'has_enhanced_summary': False,
        'implementation_estimates': {},
        'complexity_rating': 'Unknown'
    }
    
    try:
        # Read enhanced summary if available
        summary_path = os.path.join(workbook_path, 'enhanced_workbook_summary.md')
        if os.path.exists(summary_path):
            workbook_data['has_enhanced_summary'] = True
            with open(summary_path, 'r', encoding='utf-8') as f:
                summary_content = f.read()
                workbook_data = parse_enhanced_summary(summary_content, workbook_data)
        
        # Count sheet files
        sheet_files = [f for f in os.listdir(workbook_path) 
                      if f.endswith('.md') and f not in ['enhanced_workbook_summary.md', 
                                                         'workbook_summary.md',
                                                         'llm_analysis_report.md',
                                                         'software_prd.md',
                                                         'combined_enhanced_workbook.md']]
        workbook_data['sheet_count'] = len(sheet_files)
        workbook_data['sheets'] = sheet_files
        
    except Exception as e:
        print(f"Error analyzing workbook {workbook_name}: {str(e)}")
    
    return workbook_data

def parse_enhanced_summary(summary_content, workbook_data):
    """Parse enhanced summary content to extract metadata."""
    lines = summary_content.split('\n')
    
    for i, line in enumerate(lines):
        if 'Overall Complexity:' in line:
            workbook_data['complexity_rating'] = line.split(':')[1].strip()
        elif 'UI Components Required:' in line:
            try:
                count = int(line.split(':')[1].strip())
                workbook_data['implementation_estimates']['ui_components'] = count
            except:
                pass
        elif 'Business Rules to Implement:' in line:
            try:
                count = int(line.split(':')[1].strip())
                workbook_data['implementation_estimates']['business_rules'] = count
            except:
                pass
        elif 'Total Complexity Score:' in line:
            try:
                score = int(line.split(':')[1].strip())
                workbook_data['implementation_estimates']['complexity_score'] = score
            except:
                pass
    
    return workbook_data

def generate_overall_summary(results):
    """Generate overall summary statistics."""
    summary = {
        'total_workbooks': len(results['workbooks']),
        'total_files': len(results['files']),
        'total_sheets': sum(wb.get('sheet_count', 0) for wb in results['workbooks']),
        'avg_complexity': 'Unknown',
        'documentation_complete': results['has_prd'] and results['has_user_guide']
    }
    
    # Calculate average complexity
    complexity_scores = []
    for wb in results['workbooks']:
        if 'implementation_estimates' in wb and 'complexity_score' in wb['implementation_estimates']:
            complexity_scores.append(wb['implementation_estimates']['complexity_score'])
    
    if complexity_scores:
        avg_score = sum(complexity_scores) / len(complexity_scores)
        if avg_score > 500:
            summary['avg_complexity'] = 'Very High'
        elif avg_score > 200:
            summary['avg_complexity'] = 'High'
        elif avg_score > 50:
            summary['avg_complexity'] = 'Medium'
        else:
            summary['avg_complexity'] = 'Low'
    
    return summary

def format_file_size(size_bytes):
    """Format file size in human-readable format."""
    if size_bytes == 0:
        return "0B"
    size_names = ["B", "KB", "MB", "GB"]
    import math
    i = int(math.floor(math.log(size_bytes, 1024)))
    p = math.pow(1024, i)
    s = round(size_bytes / p, 2)
    return f"{s} {size_names[i]}"

def get_file_type(filename):
    """Determine file type from extension."""
    ext = Path(filename).suffix.lower()
    type_map = {
        '.md': 'Markdown',
        '.json': 'JSON Data',
        '.xlsx': 'Excel Workbook',
        '.pdf': 'PDF Document',
        '.html': 'HTML Document',
        '.txt': 'Text File'
    }
    return type_map.get(ext, 'Unknown')

def can_preview_file(filename):
    """Check if file can be previewed in browser."""
    ext = Path(filename).suffix.lower()
    previewable = ['.md', '.txt', '.json', '.html']
    return ext in previewable

if __name__ == '__main__':
    app.run(debug=True) 