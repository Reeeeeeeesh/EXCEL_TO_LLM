import os
from pathlib import Path

def combine_markdown_files(input_dir: str, output_filename: str = "combined_workbook.md") -> str:
    """
    Combines all markdown files in the input directory into a single markdown file.
    
    Args:
        input_dir: Directory containing the markdown files
        output_filename: Name of the output combined markdown file
        
    Returns:
        str: Path to the combined markdown file, or None if no files were combined
    """
    input_path = Path(input_dir)
    output_path = input_path / output_filename
    
    # Get all markdown files
    markdown_files = sorted(input_path.glob("*.md"))
    
    if not markdown_files:
        print(f"No markdown files found in {input_dir}")
        return None
    
    print(f"Found {len(markdown_files)} markdown files. Combining them...")
    
    # Combine the content
    with output_path.open('w', encoding='utf-8') as outfile:
        # Write a header for the combined file
        outfile.write(f"# Combined Workbook Analysis\n\n")
        outfile.write(f"Source directory: {input_dir}\n\n")
        outfile.write("## Table of Contents\n")
        
        # Create table of contents
        for i, md_file in enumerate(markdown_files, 1):
            worksheet_name = md_file.stem
            outfile.write(f"{i}. [{worksheet_name}](#worksheet-{i})\n")
        
        outfile.write("\n---\n\n")
        
        # Process each markdown file
        for i, md_file in enumerate(markdown_files, 1):
            print(f"Processing: {md_file.name}")
            worksheet_name = md_file.stem
            
            # Add a clear worksheet separator and header
            outfile.write(f"\n\n{'='*80}\n\n")
            outfile.write(f"<a name='worksheet-{i}'></a>\n")
            outfile.write(f"# Worksheet {i}: {worksheet_name}\n\n")
            
            # Read and write the content, skipping the original sheet header
            with md_file.open('r', encoding='utf-8') as infile:
                content = infile.read()
                # Skip the original "# Sheet: name" line and start from the dimensions
                content_lines = content.split('\n')
                if content_lines and content_lines[0].startswith('# Sheet:'):
                    content = '\n'.join(content_lines[1:])
                outfile.write(content.strip() + '\n')
    
    return str(output_path)

if __name__ == "__main__":
    # Directory containing the markdown files
    input_dir = r"C:\Users\sueho\Documents\EXCEL_TO_LLM\OUTPUT\FinModo BPFM for Corporates END"
    result = combine_markdown_files(input_dir)
    print(f"Combined file created at: {result}")
