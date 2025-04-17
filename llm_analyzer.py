import os
import google.generativeai as genai
from typing import Optional
import tiktoken
from pathlib import Path

class LLMAnalyzer:
    def __init__(self, api_key: str):
        genai.configure(api_key=api_key)
        self.model = genai.GenerativeModel('gemini-2.5-pro-preview-03-25')  # Updated to gemini-2.5-pro-preview-03-25
        self.system_prompt = """You are an advanced analytical assistant tasked with analyzing a converted Markdown representation of an Excel spreadsheet. Carefully absorb the provided content and produce a detailed report on the workings of the original spreadsheet. Your report must include:
1. A summary of the overall purpose of the spreadsheet.
2. A sheet-by-sheet breakdown with details on each sheet's purpose, formulas, data tables, and unique features.
3. An explanation of inter-sheet relationships and data dependencies.
4. Identification of the spreadsheet's unique features or innovations.
5. Recommendations for improving the spreadsheet's design or functionality."""

    def count_tokens(self, text: str) -> int:
        """Count the number of tokens in a text string."""
        encoding = tiktoken.get_encoding("cl100k_base")  # This is the encoding used by GPT-4
        return len(encoding.encode(text))

    def chunk_content(self, content: str, max_tokens: int = 30000) -> list:
        """Split content into chunks that fit within token limits."""
        chunks = []
        current_chunk = []
        current_tokens = 0
        
        # Split by lines to maintain markdown structure
        lines = content.split('\n')
        
        for line in lines:
            line_tokens = self.count_tokens(line + '\n')
            
            if current_tokens + line_tokens > max_tokens:
                # Save current chunk
                if current_chunk:
                    chunks.append('\n'.join(current_chunk))
                current_chunk = [line]
                current_tokens = line_tokens
            else:
                current_chunk.append(line)
                current_tokens += line_tokens
        
        # Add the last chunk
        if current_chunk:
            chunks.append('\n'.join(current_chunk))
        
        return chunks

    def analyze_markdown(self, markdown_content: str) -> Optional[str]:
        """
        Analyze the markdown content using Google's Gemini 2.5 Pro Preview model.
        Sends the entire content in one go as Gemini can handle larger contexts.
        """
        try:
            # Check if content is too large and chunk if necessary
            content_tokens = self.count_tokens(markdown_content)
            print(f"Total content tokens: {content_tokens}")
            
            if content_tokens > 1000000:  # Gemini's token limit is around 1M tokens
                print("Content too large, chunking for analysis...")
                chunks = self.chunk_content(markdown_content, max_tokens=800000)  # Use a safe limit
                print(f"Split content into {len(chunks)} chunks")
                
                # Process each chunk and combine results
                all_analyses = []
                for i, chunk in enumerate(chunks):
                    print(f"Processing chunk {i+1}/{len(chunks)}, tokens: {self.count_tokens(chunk)}")
                    
                    # Combine system prompt with chunk
                    chunk_prompt = f"{self.system_prompt}\n\nAnalyze this portion ({i+1}/{len(chunks)}) of the Excel spreadsheet content:\n\n{chunk}"
                    
                    # Generate analysis for this chunk
                    response = self.model.generate_content(
                        contents=chunk_prompt,
                        generation_config=genai.types.GenerationConfig(
                            temperature=0.7,
                            top_p=0.8,
                            top_k=40,
                            max_output_tokens=8192,
                        )
                    )
                    
                    if response.text:
                        all_analyses.append(response.text)
                    else:
                        print(f"Error: Empty response from Gemini for chunk {i+1}")
                
                # Combine all analyses
                if all_analyses:
                    combined_analysis = "\n\n## Analysis of Next Section\n\n".join(all_analyses)
                    return combined_analysis
                else:
                    print("Error: No successful analyses from any chunks")
                    return None
            else:
                # For smaller content, process as before
                # Combine system prompt with content
                full_prompt = f"{self.system_prompt}\n\nAnalyze this Excel spreadsheet content:\n\n{markdown_content}"
                
                # Generate analysis using the correct method
                response = self.model.generate_content(
                    contents=full_prompt,
                    generation_config=genai.types.GenerationConfig(
                        temperature=0.7,
                        top_p=0.8,
                        top_k=40,
                        max_output_tokens=8192,
                    )
                )
                
                if response.text:
                    return response.text
                else:
                    print("Error: Empty response from Gemini")
                    return None
                
        except Exception as e:
            print(f"Error in LLM analysis: {str(e)}")
            return None
            
    def save_report(self, report: str, workbook_dir: str) -> str:
        """Save the analysis report to a file."""
        try:
            report_filename = "llm_analysis_report.md"
            report_path = Path(workbook_dir) / report_filename
            
            with open(report_path, 'w', encoding='utf-8') as f:
                f.write("# Excel Workbook Analysis Report\n\n")
                f.write(report)
                
            return str(report_path)
        except Exception as e:
            print(f"Error saving report: {str(e)}")
            return ""
