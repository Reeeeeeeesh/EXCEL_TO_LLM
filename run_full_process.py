import os
from pathlib import Path
from excel_to_llm_converter import ExcelToLLMConverter
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Get API key
api_key = os.getenv('GOOGLE_API_KEY')
if not api_key:
    raise ValueError("GOOGLE_API_KEY environment variable is not set")

# Define input and output paths
INPUT_DIR = Path("C:/Users/sueho/Documents/EXCEL_TO_LLM/INPUT")
OUTPUT_DIR = Path("C:/Users/sueho/Documents/EXCEL_TO_LLM/output")

# Ensure output directory exists
os.makedirs(OUTPUT_DIR, exist_ok=True)

# Find Excel files in the input directory
excel_files = list(INPUT_DIR.glob("*.xlsx"))
if not excel_files:
    print("No Excel files found in the input directory!")
    exit(1)

# Process the first Excel file found
excel_file = excel_files[0]
print(f"Processing Excel file: {excel_file}")

# Initialize the converter with the input file, output directory, and API key
converter = ExcelToLLMConverter(
    input_path=str(excel_file),
    output_dir=str(OUTPUT_DIR),
    api_key=api_key
)

# Run the full conversion process
print("Starting conversion process...")
converter.convert_all()

print("Process completed. Check the output directory for results.")
