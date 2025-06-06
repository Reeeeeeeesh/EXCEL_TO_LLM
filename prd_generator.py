import os
import google.generativeai as genai
from typing import Optional, Dict, List, Any
import tiktoken
from pathlib import Path
import json
from datetime import datetime

class PRDGenerator:
    def __init__(self, api_key: str):
        genai.configure(api_key=api_key)
        self.model = genai.GenerativeModel('gemini-2.5-pro-preview-03-25')
        self.system_prompt = """You are an expert software architect and product manager tasked with creating a comprehensive Product Requirements Document (PRD) for recreating Excel spreadsheet functionality in a software application. Based on the detailed Excel analysis provided, create an extremely detailed PRD that would guide an AI-driven IDE (like Cursor) to build a functionally equivalent software tool.

Your PRD should include the following sections:

## 1. EXECUTIVE SUMMARY
- Business purpose and value proposition
- Target users and use cases
- Success metrics and KPIs

## 2. FUNCTIONAL REQUIREMENTS

### 2.1 Data Model & Architecture
- Database schema design (tables, relationships, constraints)
- Data types and validation rules
- Calculated fields and business logic
- Integration requirements

### 2.2 User Interface Requirements
- Screen layouts and navigation flow
- Input forms with validation
- Dashboard and reporting views
- Mobile responsiveness requirements

### 2.3 Business Logic Implementation
- Detailed calculation algorithms (step-by-step)
- Workflow automation rules
- Scenario analysis capabilities
- Real-time updates and dependencies

### 2.4 Reporting & Analytics
- Standard report formats
- Export capabilities (PDF, Excel, CSV)
- Data visualization requirements
- Drill-down and filtering capabilities

## 3. TECHNICAL SPECIFICATIONS

### 3.1 System Architecture
- Frontend technology stack recommendations
- Backend services and APIs
- Database technology and structure
- Security and authentication requirements

### 3.2 API Design
- REST API endpoints with request/response schemas
- Data flow between components
- Error handling and validation
- Performance requirements

### 3.3 Integration Requirements
- External data sources
- Third-party service integrations
- File import/export capabilities
- Real-time data synchronization

## 4. USER EXPERIENCE DESIGN

### 4.1 User Journey Mapping
- Primary user workflows
- Secondary use cases
- Error scenarios and recovery
- Onboarding and help systems

### 4.2 Interface Design Specifications
- Wireframes and mockups (described in detail)
- Component library requirements
- Accessibility standards (WCAG compliance)
- Performance expectations (load times, responsiveness)

## 5. DATA REQUIREMENTS

### 5.1 Input Data Specifications
- Required input fields and formats
- Data validation rules
- Import mechanisms and file formats
- Default values and auto-population

### 5.2 Calculated Data Logic
- Formula implementation (convert Excel formulas to programming logic)
- Dependency mapping and calculation order
- Real-time vs. batch calculation requirements
- Caching and performance optimization

### 5.3 Output Data Formats
- Report layouts and styling
- Export format specifications
- Data aggregation and summarization
- Historical data retention

## 6. IMPLEMENTATION ROADMAP

### 6.1 Phase 1: Core Functionality (MVP)
- Essential features for basic functionality
- Critical user workflows
- Basic reporting capabilities

### 6.2 Phase 2: Advanced Features
- Advanced calculations and scenarios
- Enhanced reporting and visualization
- Integration capabilities

### 6.3 Phase 3: Optimization & Scale
- Performance improvements
- Advanced analytics
- Mobile applications

## 7. TESTING REQUIREMENTS

### 7.1 Functional Testing
- Unit test specifications for calculations
- Integration test scenarios
- User acceptance test cases
- Performance benchmarks

### 7.2 Data Validation Testing
- Input validation test cases
- Calculation accuracy verification
- Edge case handling
- Error condition testing

## 8. DEPLOYMENT & MAINTENANCE

### 8.1 Infrastructure Requirements
- Server specifications and scaling
- Database backup and recovery
- Monitoring and alerting
- Security considerations

### 8.2 Support Documentation
- User manuals and help systems
- Administrator guides
- Troubleshooting procedures
- Training materials

## 9. APPENDICES

### 9.1 Excel Formula Translations
- Mapping of Excel functions to programming equivalents
- Complex formula breakdowns
- Conditional logic implementations

### 9.2 Data Dictionary
- Complete field definitions
- Business rule documentation
- Validation rule specifications

### 9.3 Wireframes and Mockups
- Detailed screen descriptions
- User interaction flows
- Component specifications

---

IMPORTANT INSTRUCTIONS:
1. Be extremely specific and detailed - include exact field names, data types, validation rules
2. Provide complete technical specifications that a developer could implement directly
3. Include security, performance, and scalability considerations
4. Map Excel formulas to equivalent programming logic with examples
5. Specify user interface elements in detail (buttons, forms, tables, charts)
6. Include specific technology recommendations with justifications
7. Provide implementation estimates and complexity ratings
8. Include specific test cases and validation scenarios
9. Reference the original Excel functionality throughout to ensure completeness
10. Use technical language appropriate for software developers and architects

Format with clear headers, bullet points, code examples where applicable, and numbered requirements for easy reference."""

    def count_tokens(self, text: str) -> int:
        """Count the number of tokens in a text string."""
        encoding = tiktoken.get_encoding("cl100k_base")
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

    def generate_prd(self, markdown_content: str, spreadsheet_metadata: Dict[str, Any] = None) -> Optional[str]:
        """
        Generate a comprehensive PRD based on the Excel analysis and metadata.
        """
        try:
            # Check content size and chunk if necessary
            content_tokens = self.count_tokens(markdown_content)
            print(f"Total content tokens: {content_tokens}")
            
            # Prepare enhanced prompt with metadata if available
            enhanced_prompt = self.system_prompt
            
            if spreadsheet_metadata:
                enhanced_prompt += f"\n\nADDITIONAL CONTEXT:\n"
                enhanced_prompt += f"- Spreadsheet contains {spreadsheet_metadata.get('sheet_count', 'N/A')} worksheets\n"
                enhanced_prompt += f"- Primary business function: {spreadsheet_metadata.get('business_type', 'Financial modeling')}\n"
                enhanced_prompt += f"- Complexity level: {spreadsheet_metadata.get('complexity', 'High')}\n"
                enhanced_prompt += f"- Key formulas identified: {', '.join(spreadsheet_metadata.get('formula_patterns', {}).keys())}\n"
            
            # Process in chunks for large documents
            print("Chunking content for PRD generation...")
            chunks = self.chunk_content(markdown_content, max_tokens=400000)
            print(f"Split content into {len(chunks)} chunks")
            
            # Process each chunk and combine results
            all_analyses = []
            for i, chunk in enumerate(chunks):
                chunk_tokens = self.count_tokens(chunk)
                print(f"Processing chunk {i+1}/{len(chunks)}, tokens: {chunk_tokens}")
                
                # Create chunk-specific prompt
                chunk_prompt = f"{enhanced_prompt}\n\nAnalyze this portion ({i+1}/{len(chunks)}) of the Excel spreadsheet for PRD generation:\n\n{chunk}"
                
                # Add context for multi-chunk processing
                if len(chunks) > 1:
                    chunk_prompt += f"\n\nNOTE: This is chunk {i+1} of {len(chunks)}. Focus on the functional requirements and technical specifications for the components described in this chunk. Ensure your PRD section integrates well with other potential chunks."
                
                try:
                    response = self.model.generate_content(
                        contents=chunk_prompt,
                        generation_config=genai.types.GenerationConfig(
                            temperature=0.3,  # Lower temperature for more structured output
                            top_p=0.9,
                            top_k=40,
                            max_output_tokens=8192,
                        )
                    )
                    
                    if response.text:
                        all_analyses.append(response.text)
                        print(f"Successfully generated PRD section {i+1}")
                    else:
                        print(f"Error: Empty response from Gemini for chunk {i+1}")
                        
                except Exception as chunk_error:
                    print(f"Error processing chunk {i+1}: {str(chunk_error)}")
                    continue
            
            # Combine all PRD sections
            if all_analyses:
                # If multiple chunks, create a synthesis prompt
                if len(all_analyses) > 1:
                    print("Synthesizing multi-chunk PRD...")
                    synthesis_prompt = f"""You are tasked with creating a final, cohesive Product Requirements Document by synthesizing the following {len(all_analyses)} partial PRD sections. These sections were generated from different parts of a complex Excel spreadsheet analysis.

Your task:
1. Merge overlapping sections intelligently
2. Ensure consistency across all technical specifications
3. Create a unified data model and architecture
4. Consolidate functional requirements without duplication
5. Provide a coherent implementation roadmap
6. Ensure all Excel functionality is captured comprehensively

Here are the partial PRD sections to synthesize:

{"=" * 80}
""" + f"\n\n{'=' * 40} SECTION BREAK {'=' * 40}\n\n".join(all_analyses)

                    try:
                        synthesis_response = self.model.generate_content(
                            contents=synthesis_prompt,
                            generation_config=genai.types.GenerationConfig(
                                temperature=0.2,
                                top_p=0.9,
                                top_k=40,
                                max_output_tokens=8192,
                            )
                        )
                        
                        if synthesis_response.text:
                            return synthesis_response.text
                        else:
                            print("Synthesis failed, returning combined sections")
                            return "\n\n# PRD SECTION BREAK\n\n".join(all_analyses)
                            
                    except Exception as synthesis_error:
                        print(f"Error in synthesis: {str(synthesis_error)}")
                        return "\n\n# PRD SECTION BREAK\n\n".join(all_analyses)
                else:
                    return all_analyses[0]
            else:
                print("Error: No successful PRD generation from any chunks")
                return None
                
        except Exception as e:
            print(f"Error in PRD generation: {str(e)}")
            return None

    def save_prd(self, prd_content: str, workbook_dir: str, filename: str = "software_prd.md") -> str:
        """Save the PRD to a file with proper formatting and metadata."""
        try:
            prd_path = Path(workbook_dir) / filename
            
            # Add header with metadata
            header = f"""# Product Requirements Document
## Software Implementation of Excel Spreadsheet

**Generated:** {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
**Source:** Excel Workbook Analysis
**Purpose:** Complete functional specification for software replication

---

"""
            
            with open(prd_path, 'w', encoding='utf-8') as f:
                f.write(header)
                f.write(prd_content)
                
            return str(prd_path)
        except Exception as e:
            print(f"Error saving PRD: {str(e)}")
            return ""

    def extract_spreadsheet_metadata(self, workbook_summary_path: str) -> Dict[str, Any]:
        """Extract key metadata from workbook summary for enhanced PRD generation."""
        try:
            if not os.path.exists(workbook_summary_path):
                return {}
                
            with open(workbook_summary_path, 'r', encoding='utf-8') as f:
                content = f.read()
            
            # Parse basic metadata from workbook summary
            metadata = {
                'sheet_count': 0,
                'formula_patterns': {},
                'business_type': 'Financial modeling',
                'complexity': 'High'
            }
            
            # Extract sheet count
            if 'Total Sheets:' in content:
                try:
                    sheet_line = [line for line in content.split('\n') if 'Total Sheets:' in line][0]
                    metadata['sheet_count'] = int(sheet_line.split(':')[1].strip())
                except:
                    pass
            
            # Extract formula patterns
            if 'Common Formula Patterns:' in content:
                lines = content.split('\n')
                patterns_start = False
                for line in lines:
                    if 'Common Formula Patterns:' in line:
                        patterns_start = True
                        continue
                    if patterns_start and line.strip().startswith('- '):
                        try:
                            pattern_info = line.strip()[2:]  # Remove '- '
                            if ':' in pattern_info:
                                pattern, count = pattern_info.split(':')
                                metadata['formula_patterns'][pattern.strip()] = int(count.split()[0])
                        except:
                            continue
                    elif patterns_start and not line.strip():
                        break
            
            return metadata
            
        except Exception as e:
            print(f"Error extracting metadata: {str(e)}")
            return {} 