---
name: PDF Master Processor
description: Advanced PDF operations including creation, editing, form filling, data extraction, merging, splitting, and report generation. Use when user needs PDF manipulation, document processing, or report creation.
version: 1.0.0
dependencies: python>=3.8, PyPDF2>=3.0.0, reportlab>=4.0.0, pdfplumber>=0.9.0, fpdf2>=2.7.0
---

# PDF Master Processor

Comprehensive PDF manipulation and document processing skill providing complete control over PDF operations from single prompts. Creates, edits, extracts, and processes PDF documents with intelligent automation and professional formatting.

## Quick Start

```python
import PyPDF2
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter, A4
import pdfplumber
from fpdf import FPDF

# Advanced PDF operations with intelligent processing
# Full implementation in scripts/pdf_master.py
```

## Core Capabilities

### 1. Intelligent PDF Creation
- **Report Generation**: Professional reports from data sources with charts and tables
- **Document Assembly**: Combine text, images, charts into formatted documents
- **Template System**: Pre-built templates for common document types
- **Dynamic Content**: Generate PDFs from databases, APIs, or user input

### 2. Advanced PDF Manipulation
- **Smart Merging**: Combine PDFs with bookmark preservation and page optimization
- **Intelligent Splitting**: Extract specific pages, sections, or chapters
- **Rotation & Scaling**: Adjust page orientation and sizing automatically
- **Watermarking**: Add logos, stamps, or security markings

### 3. Data Extraction Engine
- **Text Mining**: Extract structured data from unstructured PDF documents
- **Table Recognition**: Identify and extract tabular data with column headers
- **Form Processing**: Extract data from fillable PDF forms automatically
- **OCR Integration**: Process scanned documents with text recognition

### 4. Form Management System
- **Form Creation**: Generate fillable PDF forms from specifications
- **Automated Filing**: Fill forms programmatically from data sources
- **Validation Rules**: Implement data validation for form fields
- **Signature Integration**: Add digital signature fields and processing

### 5. Document Intelligence
- **Content Analysis**: Analyze document structure, headings, and sections
- **Metadata Management**: Extract and modify PDF properties and information
- **Security Operations**: Add/remove passwords, permissions, and encryption
- **Quality Optimization**: Compress, optimize, and clean PDF documents

## Specialized Functions

### Business Document Processing
- **Invoice Processing**: Extract data from invoices for accounting systems
- **Contract Analysis**: Parse contracts for key terms and clauses
- **Report Automation**: Generate recurring business reports from templates
- **Compliance Documents**: Create regulatory filings and compliance reports

### Legal Document Management
- **Document Assembly**: Create legal documents from templates and data
- **Redaction Services**: Automatically redact sensitive information
- **Case File Processing**: Organize and process legal case documents
- **Discovery Management**: Process large volumes of legal documents

### Academic & Research Tools
- **Research Reports**: Format academic papers with citations and bibliographies
- **Data Visualization**: Create charts and graphs within PDF documents
- **Thesis Formatting**: Professional academic document formatting
- **Publication Ready**: Prepare documents for journal submission

### Financial Document Processing
- **Statement Analysis**: Extract data from financial statements
- **Tax Document Prep**: Process tax forms and supporting documents
- **Audit Reports**: Generate formatted audit reports with findings
- **Investment Reports**: Create professional investment analysis documents

## Usage Patterns

### Document Creation
```
User: "Create a professional business report with our Q3 financial data"
→ Generates: Multi-page PDF with executive summary, financial tables,
  charts, and professional formatting
```

### Data Extraction
```
User: "Extract all invoice data from these PDF files into a spreadsheet"
→ Processes: Multiple invoice PDFs, extracts structured data,
  creates Excel summary with totals and analysis
```

### Document Assembly
```
User: "Merge these contracts and create a master agreement document"
→ Produces: Combined PDF with proper bookmarks, page numbering,
  and table of contents
```

## File Organization

- `SKILL.md` - Main instructions (this file)
- `TEMPLATES.md` - Document templates and formatting guides
- `EXTRACTION.md` - Data extraction patterns and rules
- `SECURITY.md` - PDF security and encryption guidelines
- `scripts/pdf_master.py` - Core PDF processing engine
- `scripts/data_extractor.py` - Intelligent data extraction system
- `scripts/form_processor.py` - PDF form creation and processing
- `scripts/report_generator.py` - Professional report generation
- `templates/` - Pre-built PDF templates for common documents
- `examples/` - Sample PDFs and processing examples

## Advanced Features

### Intelligent Processing
- **Auto-Detection**: Identify document types and apply appropriate processing
- **Batch Operations**: Process multiple PDFs simultaneously with optimization
- **Error Recovery**: Handle corrupted or problematic PDF files gracefully
- **Memory Optimization**: Efficient processing of large PDF files

### Integration Capabilities
- **Database Connectivity**: Generate reports directly from database queries
- **API Integration**: Fetch data from REST APIs for document generation
- **Cloud Storage**: Process documents from cloud storage platforms
- **Email Integration**: Automatically process PDF attachments

### Quality Assurance
- **Validation Engine**: Verify PDF integrity and structure
- **Accessibility Compliance**: Ensure PDFs meet accessibility standards
- **Print Optimization**: Optimize documents for high-quality printing
- **Mobile Compatibility**: Ensure PDFs display properly on mobile devices

## Security Features

### Document Protection
- **Encryption Standards**: Support for AES-256 and RC4 encryption
- **Permission Control**: Fine-grained access control for editing, printing, copying
- **Digital Signatures**: Create and verify digital signatures
- **Audit Trails**: Track document access and modifications

### Privacy Protection
- **Metadata Scrubbing**: Remove sensitive metadata from documents
- **Content Redaction**: Automatically identify and redact sensitive data
- **Secure Processing**: All operations performed in secure environment
- **Compliance Support**: GDPR, HIPAA, and other regulatory compliance

## Usage Examples

### Simple PDF Creation
```
"Create a PDF report from this data"
→ Professional formatted report with charts and tables
```

### Complex Document Processing
```
"Extract customer data from these 50 invoice PDFs and create a summary report"
→ Processes all invoices, extracts data, creates comprehensive analysis
```

### Document Assembly
```
"Combine these PDF sections into a complete proposal with table of contents"
→ Merged document with professional formatting and navigation
```

## Integration with Other Skills

- **Excel Integration**: Extract PDF data directly into Excel spreadsheets
- **PowerPoint Integration**: Create presentations from PDF content
- **Email Integration**: Generate PDFs and attach to automated emails
- **Database Integration**: Store extracted data in databases automatically

## Performance Optimization

### Processing Efficiency
- **Streaming Operations**: Process large files without memory overflow  
- **Parallel Processing**: Handle multiple documents simultaneously
- **Caching System**: Cache frequently used templates and operations
- **Compression Optimization**: Reduce file sizes while maintaining quality

### Scalability Features
- **Batch Processing**: Handle thousands of documents efficiently
- **Queue Management**: Process documents in background queues
- **Resource Monitoring**: Optimize memory and CPU usage
- **Error Handling**: Robust error recovery and logging

For template specifications, see [TEMPLATES.md](TEMPLATES.md)
For data extraction patterns, see [EXTRACTION.md](EXTRACTION.md)
For security guidelines, see [SECURITY.md](SECURITY.md)