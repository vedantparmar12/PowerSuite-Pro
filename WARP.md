# WARP.md

This file provides guidance to WARP (warp.dev) when working with code in this repository.

## Project Overview

This repository contains **6 Claude Agent Skills** that create professional business documents from single prompts. These skills follow Claude's official Skills format with YAML frontmatter and progressive disclosure architecture.

### Core Skills Architecture
- **Professional PowerPoint Creator** - Create presentations from prompts
- **Excel Master Controller** - Generate spreadsheets with formulas and charts  
- **PDF Master Processor** - Process and create PDF documents
- **Financial Analytics & Modeling Engine** - Financial modeling and analysis
- **Web Intelligence & Content Analyzer** - Web research and competitive analysis
- **Communication Master & Email Automation** - Email templates and automation

## Development Commands

### Testing
```bash
# Run comprehensive skills validation
python test_skills.py

# Test individual skills directly
python professional-ppt-skill/scripts/generate_presentation.py "Create a business presentation"
python excel-master-skill/scripts/generate_spreadsheet.py "Create a budget tracker"
```

### Dependencies Installation
```bash
# Core dependencies for all skills
pip install python-pptx>=0.6.21 openpyxl>=3.1.0 pandas>=1.5.0 pillow>=9.0.0 xlsxwriter>=3.0.0

# Advanced analytics dependencies
pip install numpy>=1.21.0 scipy>=1.9.0 matplotlib>=3.6.0 seaborn>=0.12.0

# PDF processing dependencies  
pip install PyPDF2>=3.0.0 reportlab>=4.0.0 pdfplumber>=0.9.0 fpdf2>=2.7.0

# Web intelligence dependencies
pip install requests>=2.28.0 beautifulsoup4>=4.11.0 selenium>=4.8.0 nltk>=3.8.0 textstat>=0.7.0

# Communication automation dependencies
pip install jinja2>=3.1.0 schedule>=1.2.0
```

### Skill Development
```bash
# Create new skill directory structure
mkdir new-skill-name
mkdir new-skill-name/scripts

# Add required SKILL.md with proper YAML frontmatter
echo "---\nname: Skill Name\ndescription: Skill description and when to use it\n---\n\n# Skill Name\n\n## When to use this skill\n\n## Instructions\n\n## Examples" > new-skill-name/SKILL.md
```

## Architecture Overview

### Progressive Disclosure System
This skills architecture uses a 3-level progressive disclosure system that dramatically reduces context consumption:

1. **Level 1 (Metadata)**: Always loaded - ~100 tokens per skill
2. **Level 2 (Instructions)**: Loaded when triggered - ~5k tokens each  
3. **Level 3 (Resources)**: Executed via bash - 0 context cost

### Skills Directory Structure (Claude Official Format)
```
[skill-name]/
├── SKILL.md                 # Required: YAML frontmatter + instructions
├── [REFERENCE.md]           # Optional: Additional reference materials
├── [TEMPLATES.md]           # Optional: Templates and examples
├── scripts/
│   ├── generate_[output].py # Main executable script
│   └── [helper_modules].py  # Supporting code modules
└── resources/               # Optional: Assets and data files
```

### Cross-Skill Intelligence
Skills automatically coordinate when multiple capabilities are needed:
- Excel skill processes data → PowerPoint skill creates presentations
- PDF skill extracts data → Financial skill creates models  
- Web skill researches competitors → Communication skill creates reports

## Core Implementation Patterns

### Skill Activation
Skills are activated automatically when Claude detects relevant keywords:
- "presentation", "slides", "PPT" → PowerPoint skill
- "spreadsheet", "Excel", "budget" → Excel skill
- "financial model", "DCF", "valuation" → Financial Analytics skill
- "research", "competitors", "market analysis" → Web Intelligence skill

### SKILL.md Format (Required)
Every skill must have a SKILL.md file with YAML frontmatter:
```yaml
---
name: Skill Name (64 chars max)
description: What the skill does and when to use it (1024 chars max)
---

# Skill Name

## When to use this skill
[Clear description of when Claude should invoke this skill]

## Instructions
[Step-by-step instructions for Claude]

## Examples
[Concrete examples of usage]
```

### Output Generation
All skills follow consistent output patterns:
- Generate safe filenames from prompts
- Apply professional formatting and styling
- Return file paths for generated outputs
- Include comprehensive error handling

## File Processing Guidelines

### Excel Files (`.xlsx`)
- Use openpyxl for comprehensive Excel control
- Apply corporate color schemes consistently
- Include formulas, charts, and conditional formatting
- Generate summary dashboards for complex sheets

### PowerPoint Files (`.pptx`)
- Use python-pptx for slide generation
- Implement intelligent slide type detection
- Apply consistent color schemes and typography
- Include speaker notes and navigation

### PDF Files (`.pdf`)
- Use multiple libraries (PyPDF2, reportlab, pdfplumber) for different operations
- Handle both creation and extraction use cases
- Include security and compliance features
- Support batch processing operations

## Business Logic Patterns

### Financial Modeling
- Follow Wall Street standards for DCF, LBO, and M&A models
- Include sensitivity analysis and scenario modeling
- Use industry-standard formulas and methodologies
- Ensure regulatory compliance (GAAP, IFRS, Basel III)

### Content Generation
- Analyze audience and presentation context
- Generate structured content hierarchies
- Apply domain-specific knowledge (business, sales, educational)
- Include professional design principles

### Data Processing
- Auto-detect data types and relationships
- Apply intelligent formatting and styling
- Generate appropriate visualizations
- Include data validation and error checking

## Integration Points

### API Integration Patterns
Skills can integrate with external APIs for:
- Real-time market data (financial modeling)
- Web scraping and content analysis
- Email automation and CRM systems
- Cloud storage and database connections

### Cross-Platform Compatibility
- Windows PowerShell support (current environment)
- Cross-platform file path handling
- Multiple output format support
- Cloud deployment compatibility

## Quality Standards

### Professional Output
- Every output must meet enterprise quality standards
- Consistent branding and formatting across all generated content
- Error-free calculations and professional presentation
- Comprehensive documentation and audit trails

### Code Quality
- Type hints for all functions and methods
- Comprehensive error handling and logging
- Unit tests for core functionality
- Professional commenting and documentation

### Performance Optimization
- Efficient memory usage for large files
- Batch processing capabilities
- Caching for frequently used templates
- Parallel processing where appropriate

## Advanced Features

### Automation Capabilities
- Scheduled report generation
- Event-triggered document creation  
- Batch processing for multiple files
- Integration with business workflows

### Intelligence Features
- Natural language prompt analysis
- Automatic content structuring
- Intelligent template selection
- Cross-skill coordination and data flow

### Scalability Features
- Handle enterprise-scale document volumes
- Multi-user concurrent access support
- Cloud deployment and distributed processing
- API-first architecture for integration

## Notes for Development

- Skills operate in secure Claude environment with bash execution
- All code executes via file system without context consumption
- Focus on single-prompt to professional-output workflows
- Prioritize cross-skill synergy and intelligent automation
- Maintain enterprise-grade quality and compliance standards

This architecture represents the next evolution in AI-powered business automation, providing unprecedented capability with minimal overhead through intelligent progressive disclosure.