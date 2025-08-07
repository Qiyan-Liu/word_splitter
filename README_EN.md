# Word Document Splitter

[中文](README.md) | **English**

A high-performance Word document splitting tool that intelligently splits large Word documents into multiple independent smaller documents based on document structure. Particularly suitable for academic papers, technical documents, training materials, and other documents with clear chapter structures.

## 🌟 Core Features

- 🔍 **Smart Title Recognition**: Supports Word standard heading styles, outline levels, and custom styles (such as "Style 3", "Style 4", etc.)
- 📄 **Complete Format Preservation**: 100% preserves original document text, images, tables, formatting, styles, and layout
- 🧵 **Dual-layer Multi-threading**: Document-level and chapter-level dual parallel processing, significantly improving processing speed
- 📁 **Smart Batch Processing**: Automatically scans input directory and batch processes multiple Word documents
- ⚙️ **Flexible Configuration**: Supports custom split levels, thread counts, input/output directories, and other parameters
- 📊 **Real-time Progress Display**: Processing status display with progress bars, supports processing time statistics
- 🛡️ **Error Handling**: Comprehensive exception handling and logging mechanisms
- 🎯 **Chapter Title Recognition**: Supports Chinese chapter formats (一、二、三、) and numeric formats (1、2、3、)

## 📋 System Requirements

- **Python Version**: Python 3.7+ (Python 3.11+ recommended)
- **Operating System**: Windows, macOS, Linux (Windows recommended for best compatibility)
- **Memory**: 4GB+ recommended (more memory needed for processing large documents)
- **Storage**: Ensure sufficient disk space to store split documents
- **Document Format**: Supports .docx and .doc format Microsoft Word documents

## 🚀 Quick Start

### 1. Clone or Download Project

```bash
git clone https://github.com/Qiyan-Liu/word_splitter.git
cd word-document-splitter
```

### 2. Install Dependencies

```bash
pip install -r requirements.txt
```

**Dependencies**:
- `python-docx>=0.8.11`: Core Word document processing library
- `tqdm`: Progress bar display (optional, will use simple progress display if not installed)

### 3. Prepare Documents

Create input directory and place Word documents to be split:

```bash
mkdir input
# Copy your .docx or .doc files to the input folder
```

### 4. Run Program

```bash
python main.py
```

## 📁 Project Structure

```
word-document-splitter/
├── main.py                 # Main program entry (launcher)
├── src/                    # Source code directory
│   ├── __init__.py        # Package initialization file
│   ├── app.py             # Main program implementation, contains user configuration parameters
│   └── word_splitter.py   # Core splitting logic and document processor
├── requirements.txt        # Project dependency list
├── README.md              # Project documentation (Chinese)
├── README_EN.md           # Project documentation (English)
├── word_splitter.log      # Program run log (auto-generated)
├── __pycache__/           # Python cache directory (auto-generated)
├── input/                 # Input document directory (manually created)
│   ├── document1.docx
│   ├── document2.docx
│   └── ...
└── output/                # Output document directory (auto-created)
    ├── document1/
    │   ├── Chapter 1 Overview.docx
    │   ├── Chapter 2 Details.docx
    │   └── ...
    ├── document2/
    │   ├── 1. Introduction.docx
    │   ├── 2. Background.docx
    │   └── ...
    └── ...
```

## 📖 Detailed Usage Guide

### Basic Usage Flow

1. **Prepare Documents**: Place Word documents to be split in the `input` directory
2. **Configure Parameters**: Modify configuration parameters in `src/app.py` as needed
3. **Run Program**: Execute `python main.py`
4. **View Results**: Check split documents in the `output` directory

### Supported Document Formats

- ✅ `.docx` format (recommended)
- ✅ `.doc` format
- ❌ Password-protected documents
- ❌ Corrupted documents

### Title Recognition Rules

The tool can recognize the following types of titles:

1. **Word Standard Heading Styles**:
   - English: `Heading 1`, `Heading 2`, `Heading 3`...
   - Chinese: `标题 1`, `标题 2`, `标题 3`...

2. **Word Outline Levels**: Outline levels set through paragraph formatting

3. **Custom Styles**: 
   - `样式1`, `样式2`, `样式3`... (common in academic papers)
   - Content must match heading characteristics (such as chapter format, keywords, etc.)

4. **Chapter Format Recognition**:
   - Chinese: `一、`, `二、`, `三、`...
   - Numeric: `1、`, `2、`, `3、`...

### Splitting Logic Explanation

- **Minimum Split Level**: Default is level 5, meaning only titles at level 5 or deeper will be split
- **Non-continuous Level Handling**: If document levels are non-continuous (e.g., 1→3→5), automatically adjusts to minimum splittable level
- **Content Integrity**: Each split document contains all content from current title to before the next same-level or higher-level title
- **Format Preservation**: Completely preserves all formatting, styles, images, tables, and other elements from original document

## ⚙️ Configuration Parameters

In the `main()` function of the `main.py` file, you can modify the following configuration parameters:

```python
def main():
    # ==================== Configuration Parameters ====================
    
    # Input/Output Directories
    INPUT_DIR = "input"          # Input document directory
    OUTPUT_DIR = "output"        # Output document directory
    
    # Document Splitting Configuration
    MIN_LEVEL = 5               # Minimum split level (1-6)
    
    # Multi-threading Configuration
    FILE_THREAD_COUNT = 4       # Thread count for processing multiple documents (recommended 1-8)
    CHAPTER_THREAD_COUNT = 4    # Thread count for processing chapters in single document (recommended 1-4)
```

### 📋 Detailed Parameter Description

| Parameter | Type | Default | Range | Description |
|-----------|------|---------|-------|-------------|
| `INPUT_DIR` | string | "input" | - | Input document directory path |
| `OUTPUT_DIR` | string | "output" | - | Output directory path for split documents |
| `MIN_LEVEL` | int | 5 | 1-6 | Minimum split level, corresponding to Word headings 1-6 |
| `FILE_THREAD_COUNT` | int | 4 | 1-16 | Thread count for processing multiple documents simultaneously |
| `CHAPTER_THREAD_COUNT` | int | 4 | 1-8 | Thread count for processing chapters within single document |

### 🎯 MIN_LEVEL Split Level Description

- **1**: Split only first-level headings (chapters)
- **2**: Split to second-level headings (sections)
- **3**: Split to third-level headings (subsections)
- **4**: Split to fourth-level headings (sub-subsections)
- **5**: Split to fifth-level headings (detailed chapters) - **Default recommended**
- **6**: Split to sixth-level headings (finest granularity)

**Selection Recommendations**:
- Academic papers: Recommended 3-4 levels
- Technical documents: Recommended 4-5 levels
- Training materials: Recommended 2-3 levels

## 🚀 Multi-threading Performance Optimization

### Dual-layer Multi-threading Architecture

This tool uses a dual-layer multi-threading design:
- **Document-level Thread Pool**: Parallel processing of multiple Word documents
- **Chapter-level Thread Pool**: Parallel processing of multiple chapters within single document

**Total Threads = FILE_THREAD_COUNT × CHAPTER_THREAD_COUNT**

### 🎯 Thread Configuration Strategy

#### Adjust Based on Document Characteristics

| Document Type | Document Count | Recommended Config | Description |
|---------------|----------------|-------------------|-------------|
| Large Documents | 1-3 | `FILE=1-2, CHAPTER=3-4` | Focus on chapter processing optimization |
| Medium Documents | 4-10 | `FILE=2-4, CHAPTER=2-3` | Balance document and chapter processing |
| Small Documents | 10+ | `FILE=4-8, CHAPTER=1-2` | Focus on document parallelization |

#### Adjust Based on Hardware Configuration

| Hardware Type | Recommended Config | Reason |
|---------------|-------------------|--------|
| **SSD + 8GB+ Memory** | `FILE=4-8, CHAPTER=2-4` | High I/O performance, supports more concurrency |
| **SSD + 4GB Memory** | `FILE=2-4, CHAPTER=2-3` | Balance performance and memory usage |
| **HDD + 8GB+ Memory** | `FILE=2-4, CHAPTER=2-3` | Avoid disk I/O competition |
| **HDD + 4GB Memory** | `FILE=1-2, CHAPTER=1-2` | Conservative configuration, ensure stability |

### ⚡ Performance Optimization Recommendations

1. **Memory Optimization**:
   - When processing large documents, appropriately reduce thread count
   - Monitor memory usage, avoid memory shortage

2. **Disk I/O Optimization**:
   - SSD users can use higher `FILE_THREAD_COUNT`
   - HDD users should reduce concurrency to avoid disk competition

3. **CPU Optimization**:
   - Multi-core CPUs can appropriately increase total thread count
   - Single-core or dual-core CPUs should use conservative configuration

### 📊 Performance Monitoring

During program execution, it displays:
- Real-time processing progress
- Processing status of each document
- Total processing time statistics
- Detailed logs recorded in `word_splitter.log`

## 📋 Splitting Rules Detailed

### Core Splitting Logic

1. **Level Recognition**: Automatically recognizes title level structure in document
2. **Minimum Split Level**: Determines split depth based on `MIN_LEVEL` parameter
3. **Smart Level Adjustment**: If document levels are non-continuous, automatically adjusts to appropriate split level
4. **Content Integrity**: Ensures each split document contains complete chapter content
5. **Format Preservation**: 100% preserves all elements and formatting from original document

### Splitting Example

Assume the following document structure:
```
Chapter 1 Overview (Level 1)
├── 1.1 Background Introduction (Level 2)
│   ├── 1.1.1 Historical Development (Level 3)
│   └── 1.1.2 Current Analysis (Level 3)
└── 1.2 Research Significance (Level 2)
Chapter 2 Methods (Level 1)
├── 2.1 Theoretical Foundation (Level 2)
└── 2.2 Experimental Design (Level 2)
```

**When MIN_LEVEL=3**, it will generate:
- `Chapter 1 Overview - 1.1 Background Introduction - 1.1.1 Historical Development.docx`
- `Chapter 1 Overview - 1.1 Background Introduction - 1.1.2 Current Analysis.docx`
- `Chapter 1 Overview - 1.2 Research Significance.docx`
- `Chapter 2 Methods - 2.1 Theoretical Foundation.docx`
- `Chapter 2 Methods - 2.2 Experimental Design.docx`

## 📝 Logging

### Log File
- **Filename**: `word_splitter.log`
- **Encoding**: UTF-8
- **Location**: Project root directory

### Log Content
- Document processing start and end times
- Chapter recognition and splitting process
- Error and exception information
- Performance statistics

### Log Levels
- **INFO**: Normal processing information
- **WARNING**: Warning information (such as parameters out of range)
- **ERROR**: Error information (such as unable to open document)

## ⚠️ Important Notes

### Document Requirements
- ✅ Ensure Word documents are not password protected
- ✅ Document format is `.docx` or `.doc`
- ✅ Document structure is clear with distinct title hierarchy
- ✅ Document is not corrupted or has format errors

### System Resources
- 💾 Ensure sufficient disk space (recommend at least 2-3 times original document size)
- 🧠 Pay attention to memory usage when processing large documents
- ⏱️ Complex documents (many images, tables) take longer to process

### Data Security
- 🔒 **Strongly recommend** backing up important documents before processing
- 📁 Output directory will be auto-created, won't overwrite existing files
- 🗂️ Each run creates new subdirectory in output directory

## 🔧 Troubleshooting

### Installation Issues

**Issue**: `ImportError: No module named 'docx'`
```bash
# Solution
pip install python-docx
# or
pip install -r requirements.txt
```

**Issue**: `ModuleNotFoundError: No module named 'tqdm'`
```bash
# Solution (optional, program will auto-downgrade)
pip install tqdm
```

### Document Processing Issues

**Issue**: Cannot open document
- ✅ Check if document is corrupted
- ✅ Confirm document format is `.docx` or `.doc`
- ✅ Check if document is password protected
- ✅ Confirm document path is correct, no special characters

**Issue**: Inaccurate chapter recognition
- ✅ Check document heading style settings
- ✅ Confirm paragraph outline level settings are correct
- ✅ Try adjusting `MIN_LEVEL` parameter
- ✅ Check log file for detailed information

**Issue**: Format lost in split documents
- ✅ Confirm original document format is correct
- ✅ Check if unsupported Word features were used
- ✅ Try saving original document as standard `.docx` format

### Performance Issues

**Issue**: Slow processing speed
- ✅ Adjust thread configuration parameters
- ✅ Check disk space and memory usage
- ✅ Close other resource-intensive programs
- ✅ Consider batch processing large numbers of documents

**Issue**: Insufficient memory
- ✅ Reduce `CHAPTER_THREAD_COUNT` parameter
- ✅ Reduce number of documents processed simultaneously
- ✅ Close other programs to free memory

### Output Issues

**Issue**: Output directory is empty
- ✅ Check if input directory contains Word documents
- ✅ Confirm document has recognizable title structure
- ✅ Check console output and log file
- ✅ Try reducing `MIN_LEVEL` parameter

## 🔬 Technical Implementation

### Core Technology Stack
- **Document Processing**: `python-docx` - Word document reading/writing and operations
- **Multi-threading**: `concurrent.futures.ThreadPoolExecutor` - High-performance parallel processing
- **Progress Display**: `tqdm` - Real-time progress bars and status display
- **Logging**: `logging` - Complete logging and error tracking

### Key Algorithms
1. **Smart Title Recognition Algorithm**:
   - Word standard heading style recognition
   - Outline level analysis
   - Custom style pattern matching
   - Content feature analysis (chapter format, keywords, etc.)

2. **Document Structure Analysis**:
   - Recursive level analysis
   - Chapter range calculation
   - Content integrity guarantee

3. **Format Preservation Technology**:
   - Complete paragraph format copying
   - Image and table element preservation
   - Style and layout retention
   - Document property inheritance

### Performance Optimization
- **Dual-layer Thread Pool Design**: Document-level and chapter-level parallel processing
- **Memory Management**: Smart garbage collection and memory optimization
- **I/O Optimization**: Batch file operations and caching mechanisms

## 📊 Usage Examples

### Processing Academic Papers
```bash
# Configuration recommendation: MIN_LEVEL=3, FILE_THREAD_COUNT=2, CHAPTER_THREAD_COUNT=3
python main.py
```

### Processing Technical Documents
```bash
# Configuration recommendation: MIN_LEVEL=4, FILE_THREAD_COUNT=4, CHAPTER_THREAD_COUNT=2
python main.py
```

### Batch Processing Large Numbers of Documents
```bash
# Configuration recommendation: MIN_LEVEL=5, FILE_THREAD_COUNT=6, CHAPTER_THREAD_COUNT=1
python main.py
```

## 🤝 Contributing

Contributions and suggestions are welcome!

1. Fork this project
2. Create feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to branch (`git push origin feature/AmazingFeature`)
5. Open Pull Request

## Development Notes

### 🤖 AI-Assisted Development

This project was developed with the assistance of **Trae AI** intelligent programming assistant, experiencing the powerful capabilities of AI-assisted programming.

- **Development Tool**: [Trae AI](https://trae.ai/) - World-leading AI programming assistant
- **Technical Support**: Advanced AI technology provided by ByteDance
- **Development Efficiency**: AI assistance significantly improved development efficiency and code quality

### 📚 Usage Statement

**⚠️ Important Notice: This project is for learning and research purposes only, not for commercial use.**

- ✅ **Allowed**: Personal learning, technical research, educational purposes
- ❌ **Prohibited**: Commercial sales, profit-making services, commercial distribution
- 📖 **Purpose**: Demonstrate AI-assisted programming capabilities, promote technical exchange and learning

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

**Note**: Although the code uses MIT license, please follow the above usage statement and use this project for learning purposes only.

## 👨‍💻 Author

**QiyanLiu** - *Initial Development*

## 🙏 Acknowledgments

### 🙏 Special Thanks

- **ByteDance** - Providing advanced AI technology support
- **Trae AI Team** - Developing excellent AI programming assistant
- **AI Technology** - Making programming more efficient and intelligent

### 🔧 Technical Acknowledgments

- Thanks to `python-docx` library for powerful Word document processing capabilities
- Thanks to all test users for valuable feedback and suggestions
- Open source community support and contributions
- Development of AI-assisted programming technology

### 💡 Project Significance

This project demonstrates the practical application effects of AI-assisted programming:
- **Rapid Prototyping**: AI helps quickly build project framework
- **Code Optimization**: AI provides code improvement suggestions
- **Documentation Generation**: AI assists in generating complete project documentation
- **Problem Solving**: AI assists in debugging and troubleshooting

---

**Thanks to Trae AI for making programming more intelligent and efficient!** 🚀

---

**If this tool helps you, please give it a ⭐ Star!**