# AI-Powered PowerPoint Inconsistency Detector v4.0

AI-Powered PowerPoint Inconsistency Detector is a Python tool that scans multi-slide PowerPoint decks to detect factual, numerical, and logical inconsistencies. It combines rule-based checks for quick wins with Gemini AI for deep, context-aware analysis, making review faster, more accurate, and scalable for presentations of any size.

## Problem Statement

Modern presentations often contain complex data, metrics, and claims that span multiple slides. Manual review for consistency is time-consuming and error-prone, especially in large decks with numerical data, timelines, and interconnected claims. This tool addresses the need for automated detection of:

- **Conflicting numerical data** (revenue figures, percentages, time savings)
- **Contradictory textual claims** (opposing statements about market conditions, capabilities)
- **Timeline mismatches** (conflicting dates, forecasts, sequences)
- **Mathematical errors** (breakdown components not summing to totals)
- **Unit inconsistencies** (mixing time units, currency formats)

## 📄 Project Report

The comprehensive project report documenting the development process, technical approach, and evaluation results is available at:

```PPTX_Inconsistency_Detection/docs/```

The report includes detailed analysis of the tool's evolution to an enterprise-ready solution, performance metrics, and evaluation criteria alignment.

## Key Features

### 🔍 **Dual Analysis Approach**
- **Rule-based detection**: Fast, reliable pattern matching for common inconsistency types
- **AI-powered analysis**: Deep contextual understanding using Gemini 2.5 Flash API

### 📊 **Multi-source Text Extraction**
- **Primary**: Direct text extraction from PPTX files using python-pptx
- **Fallback**: OCR processing of slide images using Tesseract
- **Hybrid**: Combines both sources for maximum accuracy

### 🚀 **Enterprise-Grade Scalability**
The tool automatically detects presentation size and adapts its processing strategy:

#### For Large Presentations (50+ slides):
- **Memory-Efficient Chunked Processing**: Processes slides in batches of 15 to handle presentations with hundreds of slides without memory issues
- **Parallel OCR Processing**: Multi-threaded OCR with configurable worker pools (up to 8 workers) for faster image processing
- **OCR Result Caching**: MD5-based LRU cache prevents reprocessing identical images (common in templates with repeated headers/footers)
- **API Rate Limiting**: Smart throttling (12 calls/minute) with exponential backoff to prevent hitting Gemini API limits during large batch operations
- **Lightweight Output**: Generates summary reports instead of full slide data to manage file sizes

#### For Small Presentations (<50 slides):
- **Standard Processing**: Fast, straightforward analysis using 4 OCR workers without enterprise overhead
- **Full Data Retention**: Includes complete slide content in output for detailed analysis
- **Batch Processing**: Processes slides in batches of 6 for optimal API efficiency

*Note: Features like OCR caching and API rate limiting add complexity that's unnecessary for small-scale presentations but become critical for enterprise-level document processing.*

### 🎯 **Specialized Detectors**
- **Impact Value Conflicts**: Detects conflicting monetary amounts ($2M vs $3M) with currency normalization
- **Time Savings Analysis**: Identifies inconsistent time metrics with unit normalization (minutes ↔ hours)
- **Mathematical Validation**: Verifies that breakdown components sum to claimed totals
- **Unit Mixing Detection**: Flags confusion between different time/currency units (per year vs per month)
- **Contextual Numeric Analysis**: Context-aware detection of conflicting values within categories
- **Percentage Sum Validation**: Ensures percentages sum to approximately 100% (±2% tolerance)

### 📈 **Enhanced Reporting **
- **Detailed Issue Breakdown**: Each issue shows specific evidence, claimed vs actual values, and breakdown details
- **Numbered AI Issues**: LLM-detected issues are numbered and formatted with proper text wrapping
- **Evidence Display**: Shows specific conflicting values and their locations
- **Enhanced Summary**: Includes analysis summary with recommendations and context
- **Progress Indicators**: Clear step-by-step progress indicators
- **Proper Slide References**: All issues properly reference slide numbers
- Severity-based classification (High/Medium/Low priority)
- Structured JSON output for programmatic use
- Professional formatting with text wrapping and indentation

## Installation

### Option 1: Virtual Environment (Recommended)

```bash
# Create virtual environment
python -m venv venv

# Activate virtual environment
# Windows:
venv\Scripts\activate
# macOS/Linux:
source venv/bin/activate

# Install dependencies
pip install -r dependencies.txt
```

### Option 2: Conda Environment

```bash
# Create conda environment
conda create -n ppt-analyzer python=3.9

# Activate environment
conda activate ppt-analyzer

# Install dependencies
pip install -r dependencies.txt
```

### Prerequisites

- **Python 3.7+**
- **Tesseract OCR**: Required for image text extraction
  - Windows: Download from [GitHub releases](https://github.com/UB-Mannheim/tesseract/wiki)
  - macOS: `brew install tesseract`
  - Ubuntu/Debian: `sudo apt install tesseract-ocr`

## Configuration

### API Key Setup
1. Obtain a free Gemini 2.5 Flash API key from [AI Studio](https://aistudio.google.com/app/apikey)
2. Create a `.env` file in the project root:
```env
my_api_key=your_gemini_api_key_here
```

### Dependencies
The tool requires these packages (specified in `dependencies.txt`):
- `python-pptx` - PowerPoint file processing
- `Pillow` - Image handling
- `pytesseract` - OCR text extraction
- `dateparser` - Date parsing and normalization
- `google-generativeai` - Gemini AI integration
- `python-dotenv` - Environment variable management
- `tqdm` - Progress bars
- Additional supporting packages for enhanced functionality

## Usage

### Basic Usage
```bash
python pptx_tool_enhanced.py
```

The script will prompt for:
- PPTX file path (auto-detects if in project directory)
- Images folder path (for OCR fallback)

### Adaptive Processing
The tool automatically detects presentation characteristics and optimizes accordingly:

```
--- ENHANCED SCALABLE PowerPoint Inconsistency Detector ---
Default PPTX:   /path/to/presentation.pptx
Default Images: /path/to/images/
----------------------------------------------------------------

⚠️  Large presentation detected (75 slides)
   Enabling chunked processing and parallel OCR...

✅ Processing 75 slides...
OCR workers: 8
Chunked processing: Yes

Processing chunk: slides 1-15
✓ Processed 15/75 slides
⏳ Rate limiting: waiting 12.3s...
Processing chunk: slides 16-30
✓ Processed 30/75 slides
```

### Input Options
1. **PPTX + Images**: Best accuracy, uses both text extraction and OCR
2. **PPTX only**: Fast processing, relies on embedded text
3. **Images only**: OCR-based analysis for image-only presentations

### Output Files
- **Terminal**: Real-time progress and detailed summary of findings
- **JSON Report**: Detailed analysis saved as `inconsistencies_enhanced.json`

## Technical Approach

### Architecture
```
Input Sources → Text Extraction → Normalization → Analysis → Enhanced Reporting
     ↓              ↓               ↓             ↓              ↓
   PPTX +        python-pptx    Number/Date   Rule-based    v3.0 Style
   Images    →  Parallel OCR  →  Extraction →   + AI     →   Terminal
                 + MD5 Cache                                   Output
```

### Scalable Analysis Pipeline

1. **Adaptive Processing Selection**
   - Automatic detection of presentation size (50+ slides threshold)
   - Dynamic resource allocation (4-8 OCR workers)
   - Memory-efficient chunked processing (15 slides per chunk)

2. **Enhanced Text Extraction**
   - Parallel OCR processing with ThreadPoolExecutor
   - MD5-based LRU caching for identical images
   - Custom OCR configuration for business-relevant characters
   - Progress tracking with tqdm integration

3. **Smart API Management**
   - Rate limiting class with 12 calls per minute limit
   - Exponential backoff retry logic (2s, 4s, 8s delays)
   - Batch optimization (6 slides per batch for small, dynamic for large)
   - Enhanced error handling with graceful degradation

4. **Memory-Optimized Processing**
   - SlideProcessor class for chunked processing
   - Lightweight summary generation for large presentations
   - Efficient garbage collection between chunks
   - Dynamic batch sizing based on presentation size

## Performance Characteristics

### Small Presentations (<50 slides)
- **Processing Time**: ~2-5 seconds per slide
- **Memory Usage**: <100MB peak
- **API Calls**: ~1 per 6 slides (batched)
- **OCR Workers**: 4 workers
- **Features**: Full data retention, standard processing

### Large Presentations (50+ slides)
- **Processing Time**: ~1-3 seconds per slide (parallelized)
- **Memory Usage**: <200MB peak (chunked processing)
- **API Calls**: Rate-limited at 12/minute with backoff
- **OCR Workers**: Up to 8 workers
- **Chunk Size**: 15 slides per chunk
- **Features**: Chunked processing, MD5 caching, lightweight output

## Evaluation Criteria Alignment

### ✅ **Accuracy & Completeness**
- Enhanced rule-based detectors with 6 specialized detection functions
- Context-aware AI analysis with structured prompts
- Multi-source text extraction with OCR fallback
- Currency and time unit normalization for accurate comparisons

### ✅ **Clarity & Usability**
- v3.0 style descriptive terminal output with numbered issues
- Detailed evidence display with specific slide references
- Text-wrapped descriptions with professional formatting
- Clear priority classification and actionable summaries

### ✅ **Scalability & Robustness**
- Automatic scaling with 50-slide threshold detection
- Production-ready API rate limiting with retry logic
- Memory-efficient chunked processing for enterprise use
- Parallel OCR with configurable worker pools

### ✅ **Thoughtful Design**
- Adaptive resource allocation based on presentation size
- Enhanced detection functions with proper slide tracking
- Professional output formatting matching v3.0 style
- Comprehensive error handling and graceful degradation

## Limitations & Considerations

### Current Limitations
- **API Dependency**: Deep analysis requires Gemini API (free tier: ~250 requests/day)
- **OCR Accuracy**: Complex layouts or poor image quality may affect text extraction
- **Language Support**: Optimized for English presentations only
- **Rate Limits**: Large presentations may require extended processing time due to API throttling

### Scalability Considerations
- **Enterprise Features**: MD5 caching and rate limiting add complexity for small presentations
- **Memory Trade-offs**: Large presentation mode prioritizes efficiency over complete data retention
- **Processing Time**: Rate limiting extends processing time for large decks
- **Chunk Size**: 15-slide chunks balance memory efficiency with processing overhead

### Future Enhancements
- Support for additional file formats (PDF, Google Slides)
- Multi-language inconsistency detection
- Custom rule definition interface
- Integration with presentation authoring tools
- Distributed processing for extremely large document collections

## Example Output

### Small Presentation
```
--- ENHANCED SCALABLE PowerPoint Inconsistency Detector ---
✅ Processing 12 slides...
OCR workers: 4
Chunked processing: No

[1/6] Extracting text from PPTX slides...
[2/6] Running parallel OCR on slide images...
[3/6] Normalizing slides (extracting numbers/dates)...
[4/6] Running enhanced rule-based detectors...
[5/6] Running enhanced Gemini deep checks (batched)...
[6/6] Generating enhanced report...

=== Enhanced Analysis Complete ===
Slides processed: 12
Total issues found: 8
  - High priority: 2
  - Medium priority: 3
  - Low priority: 0
Rule-based issues: 5
LLM issues: 3

🚨 HIGH PRIORITY ISSUES DETECTED:
  - Sum Breakdown Mismatch (Slide 3): Claimed total (50) doesn't match sum of breakdown (80)
    → Claimed Total: 50
    → Actual Sum: 80
    → Breakdown: [30, 25, 25]

  - Impact Value Conflict (Slides 2, 5): Found conflicting impact values: [2000000.0, 3000000.0]
    → Evidence: [2000000.0, 3000000.0]

🔍 AI-DETECTED ISSUES:
  1. Major Monetary Conflict (Slides 2, 5):
     The 'Overall Productivity Gains' section claims an objective to 'save $2M in lost 
     productivity.' However, slide 5 references '$3M saved' creating a direct contradiction 
     in the monetary impact claims.

📊 ANALYSIS SUMMARY:
   • Total Issues Found: 8
   • Critical Issues: 2 (require immediate attention)
   • Warning Issues: 3 (should be reviewed)
   • Rule-based Detection: 5 issues
   • AI Analysis: 3 issues
```

### Large Presentation
```
⚠️  Large presentation detected (75 slides)
   Enabling chunked processing and parallel OCR...

Processing chunk: slides 1-15
✓ Processed 15/75 slides
Processing chunk: slides 16-30
⏳ Rate limiting: waiting 12.3s...
✓ Processed 30/75 slides

=== Enhanced Analysis Complete ===
Slides processed: 75
Total issues found: 18
Report saved to: inconsistencies_enhanced_v4.json
```

## Files Generated

1. **`inconsistencies_enhanced_v4.json`**: Main analysis report with detailed findings
2. **Temporary cache files**: Automatically cleaned up after processing (MD5-based OCR cache)

## Technical Implementation Notes

### Key Classes and Functions
- `SlideProcessor`: Handles memory-efficient chunked processing
- `GeminiRateLimiter`: Manages API rate limiting with intelligent backoff
- `ocr_image_bytes_cached()`: MD5-based caching for OCR results
- Enhanced detection functions with proper slide tracking
- `main_enhanced()`: Adaptive main function with v3.0 style output

### Performance Optimizations
- MD5 hash-based OCR caching prevents duplicate processing
- ThreadPoolExecutor for parallel OCR operations
- Custom OCR character whitelist for faster processing
- Dynamic batch sizing based on presentation characteristics
- LRU cache with configurable size limits

## License

This project is provided for evaluation purposes. Please ensure compliance with all third-party library licenses and API terms of service.

## Support

For technical issues or questions about implementation details, please refer to the comprehensive code comments and docstrings which provide detailed explanations of each component's functionality.