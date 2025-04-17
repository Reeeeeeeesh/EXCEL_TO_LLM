# Excel to LLM Converter

This application converts Excel spreadsheets into structured data for Large Language Model (LLM) processing. It extracts data from Excel files and prepares it for analysis using Google's Gemini API.

## Features

- Upload Excel (.xlsx) files through a web interface
- Process spreadsheet data into LLM-friendly formats
- Generate insights and analysis using LLM technology
- Download processed results

## Setup

1. Clone this repository
2. Install dependencies:
   ```
   pip install -r requirements.txt
   ```
3. Create a `.env` file with your Google API key:
   ```
   GOOGLE_API_KEY=your_api_key_here
   ```
4. Run the application:
   ```
   python app.py
   ```

## Project Structure

- `app.py`: Flask web application
- `excel_to_llm_converter.py`: Core conversion logic
- `llm_analyzer.py`: LLM integration for data analysis
- `templates/`: HTML templates for the web interface
- `uploads/`: Temporary storage for uploaded files
- `output/`: Generated output files

## Requirements

See `requirements.txt` for a list of dependencies.
