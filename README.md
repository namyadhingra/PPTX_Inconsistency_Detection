# AI-Powered PowerPoint Inconsistency Detector

An intelligent Python tool that analyzes multi-slide PowerPoint presentations to identify factual and logical inconsistencies across slides using both rule-based detection and AI-powered analysis.

## Problem Statement

Modern presentations often contain complex data, metrics, and claims that span multiple slides. Manual review for consistency is time-consuming and error-prone, especially in large decks with numerical data, timelines, and interconnected claims. This tool addresses the need for automated detection of:

- **Conflicting numerical data** (revenue figures, percentages, time savings)
- **Contradictory textual claims** (opposing statements about market conditions, capabilities)
- **Timeline mismatches** (conflicting dates, forecasts, sequences)
- **Mathematical errors** (breakdown components not summing to totals)
- **Unit inconsistencies** (mixing time units, currency formats)

## Key Features

### 🔍 **Dual Analysis Approach**
- **Rule-based detection**: Fast, reliable pattern matching for common inconsistency types
- **AI-powered analysis**: Deep contextual understanding using Gemini 2.5 Flash API

### 📊 **Multi-source Text Extraction**
- **Primary**: Direct text extraction from PPTX files using python-pptx
- **Fallback**: OCR processing of slide images using Tesseract
- **Hybrid**: Combines both sources for maximum accuracy

### 🎯 **Specialized Detectors**
- **Impact Value Conflicts**: Detects conflicting monetary amounts and savings claims
- **Time Savings Analysis**: Identifies inconsistent time metrics with unit normalization
- **Mathematical Validation**: Verifies that breakdown components sum to claimed totals
- **Unit Mixing Detection**: Flags confusion between different time/currency units
- **Contextual Analysis**: AI-driven detection of subtle logical inconsistencies

### 📈 **Enhanced Reporting**
- Severity-based classification (High/Medium/Low priority)
- Detailed slide references and context
- Structured JSON output for programmatic use
- Clear terminal summaries with actionable insights

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
- `python-pptx==1.0.2` - PowerPoint file processing
- `Pillow==11.3.0` - Image handling
- `pytesseract==0.3.13` - OCR text extraction
- `dateparser==1.2.2` - Date parsing and normalization
- `google-generativeai==0.8.5` - Gemini AI integration
- `python-dotenv==1.1.1` - Environment variable management
- `tqdm==4.67.1` - Progress bars
- `lxml==6.0.0` - XML processing
- `requests==2.32.4` - HTTP requests

## Usage

### Basic Usage
```bash
python inconsistency_detector.py
```

The script will prompt for:
- PPTX file path (auto-detects if in project directory)
- Images folder path (for OCR fallback)

### Input Options
1. **PPTX + Images**: Best accuracy, uses both text extraction and OCR
2. **PPTX only**: Fast processing, relies on embedded text
3. **Images only**: OCR-based analysis for image-only presentations

### Output
- **Terminal**: Real-time progress and summary of findings
- **JSON Report**: Detailed analysis saved as `inconsistencies_enhanced.json`

## Technical Approach

### Architecture
```
Input Sources → Text Extraction → Normalization → Analysis → Reporting
     ↓              ↓               ↓             ↓          ↓
   PPTX +        python-pptx    Number/Date   Rule-based  Priority-based
   Images    →    Tesseract  →  Extraction →   + AI     →   Summary
```

### Analysis Pipeline

1. **Text Extraction Phase**
   - Direct PPTX text extraction with embedded image OCR
   - Batch OCR processing of slide images with progress tracking

2. **Normalization Phase**
   - Number extraction with currency/percentage recognition
   - Date parsing with format standardization
   - Context preservation for semantic analysis

3. **Rule-based Detection**
   - Pattern matching for common inconsistency types
   - Mathematical validation of sums and breakdowns
   - Unit normalization and conflict detection

4. **AI-Enhanced Analysis**
   - Contextual understanding of claims and relationships
   - Batch processing to optimize API usage (8 slides per request)
   - Semantic inconsistency detection beyond pattern matching

## Evaluation Criteria Alignment

### ✅ **Accuracy & Completeness**
- Multi-layered detection combining rule-based and AI approaches
- Specialized detectors for common business presentation inconsistencies
- Context-aware analysis to reduce false positives

### ✅ **Clarity & Usability**
- Priority-based issue classification
- Clear terminal output with specific slide references
- Structured JSON for integration with other tools

### ✅ **Scalability & Robustness**
- Batch processing for large presentations
- Efficient API usage within free tier limits (250 requests/day)
- Graceful handling of various input formats and edge cases

### ✅ **Thoughtful Design**
- Hybrid text extraction maximizes data capture
- Progressive enhancement from fast rules to deep AI analysis
- Modular architecture for easy extension and customization

## Limitations & Considerations

### Current Limitations
- **API Dependency**: Deep analysis requires Gemini API (250 free requests/day)
- **OCR Accuracy**: Complex layouts or poor image quality may affect text extraction
- **Language Support**: Optimized for English presentations
- **Context Understanding**: AI may miss domain-specific nuances

### Performance Characteristics
- **Speed**: ~2-5 seconds per slide for full analysis
- **Memory**: Minimal footprint, processes slides incrementally
- **API Usage**: Batched requests to maximize free tier efficiency

### Future Enhancements
- Support for additional file formats (PDF, Google Slides)
- Multi-language inconsistency detection
- Integration with presentation authoring tools
- Custom rule definition interface

## Example Output

```
=== Enhanced Analysis Complete ===
Slides processed: 12
Total issues found: 5
  - High priority: 2
  - Medium priority: 2
  - Low priority: 1

🚨 HIGH PRIORITY ISSUES DETECTED:
  - Sum Breakdown Mismatch (Slide 8): Claimed total (25) doesn't match sum of breakdown (23)
  - Impact Value Conflict (Slide 5): Found conflicting impact values: [2000000.0, 3000000.0]

⚠️ MEDIUM PRIORITY ISSUES:
  - Time Per Slide Conflict (Slides 3, 7): Found conflicting time savings per slide values
```

## License

This project is provided for evaluation purposes. Please ensure compliance with all third-party library licenses and API terms of service.

## Support

For technical issues or questions about implementation details, please refer to the code comments and docstrings which provide detailed explanations of each component's functionality.