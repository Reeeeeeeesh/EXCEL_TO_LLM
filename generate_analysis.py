import os
from pathlib import Path
from llm_analyzer import LLMAnalyzer
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Get API key
api_key = os.getenv('GOOGLE_API_KEY')
if not api_key:
    raise ValueError("GOOGLE_API_KEY environment variable is not set")

# Initialize the LLM analyzer with the updated code
llm_analyzer = LLMAnalyzer(api_key)

# Path to the existing combined workbook
workbook_dir = Path("C:/Users/sueho/Documents/EXCEL_TO_LLM/output/FinModo_BPFM_for_Corporates_END")
combined_file = workbook_dir / "combined_workbook.md"

print(f"Reading combined markdown file: {combined_file}")
try:
    with open(combined_file, 'r', encoding='utf-8') as f:
        markdown_content = f.read()
    print(f"Successfully read markdown content, length: {len(markdown_content)}")
    
    # Analyze with LLM using our updated method that handles chunking
    print("Analyzing with Gemini 2.5 Pro Preview LLM...")
    analysis_report = llm_analyzer.analyze_markdown(markdown_content)
    
    if analysis_report:
        # Save the analysis report
        report_path = llm_analyzer.save_report(analysis_report, str(workbook_dir))
        print(f"LLM analysis saved to: {report_path}")
    else:
        print("Error: LLM analysis failed")
except Exception as e:
    print(f"Error in LLM analysis: {str(e)}")
