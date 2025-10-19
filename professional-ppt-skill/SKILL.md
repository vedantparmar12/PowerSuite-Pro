---
name: Professional PowerPoint Creator
description: Create professional presentations from single prompts with comprehensive content, design, animations, and branding. Use when user requests presentations, PPTs, slides, or pitch decks.
version: 1.0.0
dependencies: python>=3.8, python-pptx>=0.6.21, pillow>=9.0.0
---

# Professional PowerPoint Creator

Transform any user request into a comprehensive, professional presentation with intelligent content generation, stunning visuals, and cohesive branding.

## Quick Start

```python
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor

# Create presentation with professional template
prs = Presentation()
# Implementation details in scripts/ppt_creator.py
```

## Core Capabilities

### 1. Intelligent Content Generation
- **Topic Analysis**: Break down user prompts into logical slide sequences
- **Content Expansion**: Generate comprehensive talking points and supporting details  
- **Structure Optimization**: Organize information hierarchically (title → main points → details)
- **Audience Adaptation**: Adjust tone and depth based on context clues

### 2. Professional Design System
- **Color Schemes**: Apply consistent, psychology-based color palettes
- **Typography**: Hierarchical font system with proper contrast ratios
- **Layout Templates**: Grid-based layouts optimized for readability
- **Visual Hierarchy**: Strategic use of size, color, and positioning

### 3. Smart Content Types
- **Executive Summaries**: Key points with data visualization
- **Process Flows**: Step-by-step workflows with connectors
- **Data Presentations**: Charts, graphs, and infographics
- **Comparison Slides**: Side-by-side analysis with visual emphasis
- **Call-to-Action**: Compelling conclusion slides with next steps

### 4. Advanced Features
- **Brand Integration**: Apply logos, colors, and fonts consistently
- **Animation Sequences**: Subtle transitions that enhance rather than distract
- **Speaker Notes**: Detailed notes for each slide
- **Interactive Elements**: Clickable navigation and embedded media support

## Usage Patterns

### Single Prompt to Complete Presentation
```
User: "Create a presentation about renewable energy for board meeting"
→ Generates: Title slide, executive summary, market analysis, technology overview, 
  financial projections, implementation roadmap, Q&A preparation
```

### Domain-Specific Adaptations
- **Business**: Financial data, market analysis, strategic planning
- **Educational**: Learning objectives, progressive concepts, knowledge checks
- **Sales**: Problem-solution-benefit structure, social proof, pricing
- **Technical**: Architecture diagrams, process flows, implementation steps

## File Organization

- `SKILL.md` - Main instructions (this file)
- `TEMPLATES.md` - Professional slide templates and layouts
- `BRANDING.md` - Brand guidelines and color psychology
- `ANIMATIONS.md` - Subtle animation and transition guidelines
- `scripts/ppt_creator.py` - Core presentation generation engine
- `scripts/content_analyzer.py` - Intelligent content structuring
- `scripts/design_system.py` - Professional styling and layouts
- `templates/` - Pre-built slide templates
- `assets/` - Icons, graphics, and design elements

## Key Principles

1. **Content First**: Always prioritize clear, valuable information
2. **Visual Hierarchy**: Guide viewer attention through strategic design
3. **Professional Polish**: Every element should look intentionally crafted
4. **Audience Focus**: Adapt complexity and style to intended viewers
5. **Actionable Outcomes**: Include clear next steps or takeaways

## Integration Notes

- Automatically detects presentation requests in user prompts
- Combines with Excel skill for data-driven slides
- Outputs ready-to-present .pptx files
- Includes speaker notes and presentation tips
- Supports both creative and corporate presentation styles

For detailed template specifications, see [TEMPLATES.md](TEMPLATES.md)
For branding guidelines, see [BRANDING.md](BRANDING.md)
For animation best practices, see [ANIMATIONS.md](ANIMATIONS.md)