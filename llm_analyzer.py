import os
import google.generativeai as genai
from typing import Optional
import tiktoken
from pathlib import Path

class LLMAnalyzer:
    def __init__(self, api_key: str):
        genai.configure(api_key=api_key)
        self.model = genai.GenerativeModel('gemini-2.5-pro-preview-03-25')  # Updated to gemini-2.5-pro-preview-03-25
        self.system_prompt = """You are an advanced analytical assistant tasked with creating a user guide for an Excel spreadsheet based on its Markdown representation. Your goal is to help a first-time user understand how to use this spreadsheet effectively. Produce a detailed, practical guide that includes:

1. EXECUTIVE SUMMARY: A brief overview of what this spreadsheet does and its primary purpose (2-3 sentences).

2. USER WORKFLOW: Step-by-step instructions on how a user should interact with the spreadsheet:
   - Which sheets to start with
   - What inputs they need to provide and where
   - How to navigate between sheets
   - How to interpret the outputs

3. KEY INPUTS: Identify and explain all important input cells/sections:
   - Where they are located (sheet and cell references if available)
   - What each input represents
   - Acceptable values or ranges
   - How changing these inputs affects the outputs

4. KEY OUTPUTS: Identify and explain all important output cells/sections:
   - Where they are located
   - What each output represents
   - How to interpret the results
   - Any visualizations or dashboards available

5. CALCULATION LOGIC: Explain in simple terms how the spreadsheet works:
   - How inputs are transformed into outputs
   - Key formulas and their purpose (in plain English, not Excel syntax)
   - Any macros, data tables, or scenarios and how to use them

6. TROUBLESHOOTING: Common issues users might encounter and how to resolve them.

7. RECOMMENDATIONS: Suggestions for best practices when using this spreadsheet.

Use clear, non-technical language suitable for business users who may not be Excel experts. Include specific cell references and sheet names whenever possible. Format your response with headers, bullet points, and numbered lists for readability."""

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
            
            # Always use chunking for large documents to be safe
            print("Chunking content for analysis...")
            chunks = self.chunk_content(markdown_content, max_tokens=500000)  # Much more conservative limit
            print(f"Split content into {len(chunks)} chunks")
            
            # Process each chunk and combine results
            all_analyses = []
            for i, chunk in enumerate(chunks):
                chunk_tokens = self.count_tokens(chunk)
                print(f"Processing chunk {i+1}/{len(chunks)}, tokens: {chunk_tokens}")
                
                # Combine system prompt with chunk
                chunk_prompt = f"{self.system_prompt}\n\nAnalyze this portion ({i+1}/{len(chunks)}) of the Excel spreadsheet content:\n\n{chunk}"
                
                # Generate analysis for this chunk
                try:
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
                        print(f"Successfully analyzed chunk {i+1}")
                    else:
                        print(f"Error: Empty response from Gemini for chunk {i+1}")
                except Exception as chunk_error:
                    print(f"Error processing chunk {i+1}: {str(chunk_error)}")
                    # Try with an even smaller chunk if possible
                    if chunk_tokens > 200000:
                        print(f"Attempting to split chunk {i+1} further...")
                        subchunks = self.chunk_content(chunk, max_tokens=200000)
                        print(f"Split chunk {i+1} into {len(subchunks)} subchunks")
                        
                        for j, subchunk in enumerate(subchunks):
                            try:
                                subchunk_prompt = f"{self.system_prompt}\n\nAnalyze this portion ({i+1}.{j+1}) of the Excel spreadsheet content:\n\n{subchunk}"
                                subresponse = self.model.generate_content(
                                    contents=subchunk_prompt,
                                    generation_config=genai.types.GenerationConfig(
                                        temperature=0.7,
                                        top_p=0.8,
                                        top_k=40,
                                        max_output_tokens=8192,
                                    )
                                )
                                
                                if subresponse.text:
                                    all_analyses.append(subresponse.text)
                                    print(f"Successfully analyzed subchunk {i+1}.{j+1}")
                                else:
                                    print(f"Error: Empty response from Gemini for subchunk {i+1}.{j+1}")
                            except Exception as subchunk_error:
                                print(f"Error processing subchunk {i+1}.{j+1}: {str(subchunk_error)}")
            
            # Combine all analyses
            if all_analyses:
                combined_analysis = "\n\n## Analysis of Next Section\n\n".join(all_analyses)
                return combined_analysis
            else:
                print("Error: No successful analyses from any chunks")
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
