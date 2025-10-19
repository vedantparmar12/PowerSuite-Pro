---
name: Web Intelligence & Content Analyzer
description: Advanced web content analysis, competitive research, market intelligence, and automated report generation from web sources. Use for research, competitive analysis, market research, or web data extraction.
version: 1.0.0
dependencies: python>=3.8, requests>=2.28.0, beautifulsoup4>=4.11.0, selenium>=4.8.0, nltk>=3.8.0, textstat>=0.7.0
---

# Web Intelligence & Content Analyzer

Comprehensive web research and content analysis skill providing intelligent web scraping, competitive analysis, market research, and automated report generation from web sources. Transforms URLs and research queries into professional insights and reports.

## Quick Start

```python
import requests
from bs4 import BeautifulSoup
from selenium import webdriver
import nltk
from textstat import flesch_reading_ease

# Advanced web intelligence with professional analysis
# Full implementation in scripts/web_intelligence.py
```

## Core Capabilities

### 1. Intelligent Web Scraping
- **Smart Content Extraction**: Identify and extract meaningful content from any website
- **Dynamic Site Handling**: JavaScript-heavy sites using Selenium automation
- **Bulk URL Processing**: Process multiple websites simultaneously with optimization
- **Content Classification**: Automatically categorize and structure extracted data

### 2. Competitive Intelligence
- **Competitor Analysis**: Comprehensive competitor website and content analysis
- **Market Positioning**: Compare messaging, pricing, and positioning strategies
- **Feature Comparison**: Extract and compare product/service features across sites
- **Content Strategy Analysis**: Analyze content themes, frequency, and engagement

### 3. Market Research Automation
- **Industry Analysis**: Gather industry trends and market intelligence from multiple sources
- **News Monitoring**: Track industry news and company mentions across web sources
- **Social Media Intelligence**: Analyze social media presence and engagement metrics
- **Review & Sentiment Analysis**: Aggregate and analyze customer reviews and feedback

### 4. Content Analysis Engine
- **Readability Analysis**: Assess content complexity and reading level
- **SEO Analysis**: Evaluate on-page SEO factors and optimization opportunities
- **Content Gaps**: Identify content opportunities compared to competitors
- **Topic Modeling**: Extract key themes and topics from large content volumes

### 5. Automated Research Reports
- **Executive Summaries**: Generate C-level research summaries with key insights
- **Competitive Benchmarking**: Professional competitor comparison reports
- **Market Intelligence**: Industry trend reports with data visualization
- **Due Diligence Reports**: Comprehensive company and market analysis

## Specialized Research Functions

### Business Intelligence
- **Company Profiles**: Automated company research and profile generation
- **Financial Intelligence**: Extract financial information from public sources
- **Leadership Analysis**: Research executive teams and key personnel
- **Partnership Mapping**: Identify business relationships and partnerships

### Market Analysis
- **Trend Identification**: Detect emerging trends and market shifts
- **Consumer Behavior**: Analyze customer preferences and behavior patterns
- **Pricing Intelligence**: Monitor and analyze pricing strategies across markets
- **Product Launch Tracking**: Monitor new product announcements and releases

### Regulatory & Compliance Research
- **Regulatory Monitoring**: Track regulatory changes and compliance requirements
- **Legal Intelligence**: Monitor legal developments and court cases
- **Policy Analysis**: Research government policies and regulatory impacts
- **Risk Assessment**: Identify potential regulatory and business risks

### Academic & Technical Research
- **Literature Review**: Automated academic paper and research aggregation
- **Technical Analysis**: Extract and analyze technical specifications and capabilities
- **Patent Research**: Monitor patent filings and intellectual property developments
- **Innovation Tracking**: Track technological developments and innovations

## Advanced Analysis Features

### Natural Language Processing
- **Content Summarization**: Automatic summarization of lengthy web content
- **Keyword Extraction**: Identify key terms and concepts from large content volumes
- **Sentiment Analysis**: Analyze sentiment and tone across multiple sources
- **Language Detection**: Multi-language content processing and analysis

### Data Visualization
- **Trend Charts**: Visualize data trends and patterns over time
- **Competitive Maps**: Visual positioning and competitive landscape mapping
- **Network Analysis**: Relationship mapping and network visualization
- **Geographic Analysis**: Location-based analysis and mapping

### Machine Learning Integration
- **Content Classification**: Automated categorization of web content
- **Anomaly Detection**: Identify unusual patterns or outliers in data
- **Predictive Analytics**: Forecast trends based on historical web data
- **Recommendation Engines**: Suggest relevant content and research directions

## Usage Patterns

### Competitive Research
```
User: "Analyze our top 5 competitors' websites and create a competitive intelligence report"
→ Generates: Comprehensive analysis with feature comparisons,
  pricing intelligence, content strategy, and positioning insights
```

### Market Intelligence
```
User: "Research the fintech market trends and create an industry report"
→ Produces: Market size data, trend analysis, key players overview,
  regulatory landscape, and growth projections
```

### Due Diligence Research
```
User: "Research this company for potential acquisition - analyze financials, team, and market position"
→ Creates: Complete due diligence report with financial analysis,
  leadership assessment, competitive position, and risk factors
```

## File Organization

- `SKILL.md` - Main instructions (this file)
- `SOURCES.md` - Web source templates and extraction patterns
- `ANALYSIS.md` - Content analysis methodologies and frameworks
- `REPORTS.md` - Report templates and formatting guidelines
- `scripts/web_intelligence.py` - Core web scraping and analysis engine
- `scripts/content_analyzer.py` - Advanced content analysis and NLP
- `scripts/competitor_tracker.py` - Competitive intelligence automation
- `scripts/report_generator.py` - Professional report generation
- `scripts/data_visualizer.py` - Charts and visualization creation
- `templates/` - Pre-built report templates for different research types
- `datasets/` - Reference datasets and benchmark data

## Ethical & Legal Compliance

### Web Scraping Ethics
- **Robots.txt Compliance**: Respect website scraping policies and limitations
- **Rate Limiting**: Implement responsible scraping with appropriate delays
- **Terms of Service**: Compliance with website terms and conditions
- **Data Privacy**: Handle personal and sensitive data appropriately

### Legal Considerations
- **GDPR Compliance**: European data protection regulation compliance
- **CCPA Compliance**: California consumer privacy act adherence
- **Fair Use**: Appropriate use of copyrighted content for analysis
- **Attribution**: Proper source citation and attribution practices

### Quality Assurance
- **Data Validation**: Verify accuracy and reliability of extracted data
- **Source Credibility**: Assess and weight source reliability and authority
- **Bias Detection**: Identify and mitigate potential bias in analysis
- **Fact Checking**: Cross-reference information across multiple sources

## Integration Capabilities

### Data Export
- **Excel Reports**: Structured data export to Excel with formatting
- **PowerPoint Integration**: Automated presentation creation from research
- **PDF Reports**: Professional formatted reports with charts and analysis
- **Database Storage**: Store research data in structured databases

### API Integrations
- **Social Media APIs**: Twitter, LinkedIn, Facebook data integration
- **News APIs**: Real-time news monitoring and analysis
- **Financial APIs**: Market data and financial information integration
- **Search APIs**: Enhanced search capabilities across multiple engines

### Collaboration Tools
- **Team Sharing**: Share research findings and collaborate on analysis
- **Version Control**: Track research iterations and updates
- **Annotation System**: Add notes and insights to research findings
- **Workflow Integration**: Connect with project management and CRM systems

## Advanced Features

### Real-Time Monitoring
- **Alert Systems**: Automated alerts for specific content changes or mentions
- **Trend Monitoring**: Real-time tracking of emerging trends and topics
- **Competitive Monitoring**: Ongoing surveillance of competitor activities
- **News Monitoring**: Continuous monitoring of relevant news and updates

### Automation & Scheduling
- **Scheduled Research**: Automated research reports on regular schedules
- **Trigger-Based Analysis**: Automated analysis based on specific events
- **Workflow Automation**: Connect research to downstream business processes
- **Batch Processing**: Handle large-scale research projects efficiently

### Advanced Analytics
- **Predictive Modeling**: Forecast trends and outcomes based on web data
- **Statistical Analysis**: Statistical significance testing and analysis
- **Time Series Analysis**: Analyze trends and patterns over time
- **Correlation Analysis**: Identify relationships between different data points

## Usage Examples

### Simple Research
```
"Research renewable energy market trends"
→ Comprehensive market analysis with trends, players, and projections
```

### Complex Intelligence
```
"Analyze our competitor's product launch strategy across 20 websites and social channels"
→ Multi-source analysis with launch timeline, messaging analysis, and impact assessment
```

### Automated Monitoring
```
"Set up monitoring for mentions of our company and competitors across industry publications"
→ Ongoing monitoring system with weekly intelligence reports
```

## Performance & Scalability

### Processing Efficiency
- **Parallel Processing**: Concurrent analysis of multiple websites
- **Caching System**: Cache frequently accessed data for faster processing
- **Content Optimization**: Intelligent content filtering and prioritization
- **Resource Management**: Optimize memory and bandwidth usage

### Quality Controls
- **Error Handling**: Robust error recovery and retry mechanisms
- **Data Quality**: Automated data quality checks and validation
- **Source Reliability**: Continuous assessment of source credibility
- **Accuracy Verification**: Cross-source validation and fact-checking

For web source patterns, see [SOURCES.md](SOURCES.md)
For analysis methodologies, see [ANALYSIS.md](ANALYSIS.md)
For report templates, see [REPORTS.md](REPORTS.md)