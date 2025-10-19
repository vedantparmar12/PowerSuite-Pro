---
name: Professional PowerPoint Creator
description: Create professional presentations from prompts with intelligent content, design, and branding. Use when user requests presentations, PPTs, slides, pitch decks, or business presentations.
version: 1.0.0
dependencies: python>=3.8, python-pptx>=0.6.21, pillow>=9.0.0
---

# Professional PowerPoint Creator

Create comprehensive, professional presentations from single prompts with intelligent content generation and cohesive design.

## When to use this skill

Use this skill when users request:
- PowerPoint presentations or slides
- Business presentations or pitch decks
- Sales presentations or proposals
- Training materials or educational slides
- Executive briefings or board presentations

## Quick Start

Create a presentation by running the PowerPoint generator script:

```bash
python scripts/generate_presentation.py "[user prompt]" [output_filename.pptx]
```

## Core Features

- **Intelligent Content Generation**: Analyzes prompts to create structured slide sequences
- **Professional Design**: Applies consistent color schemes and typography
- **Multiple Templates**: Business, educational, sales, and technical presentation styles
- **Automatic Formatting**: Professional layouts with proper visual hierarchy

## Basic Usage

1. Analyze the user's prompt to determine presentation type and content
2. Generate appropriate slide structure (title, agenda, content, conclusion)
3. Apply professional formatting and design elements
4. Create .pptx file with speaker notes

## Advanced Features

**Design Templates**: See [TEMPLATES.md](TEMPLATES.md) for slide layouts and design patterns

**Brand Guidelines**: See [BRANDING.md](BRANDING.md) for color schemes and typography standards

**Animation Guide**: See [ANIMATIONS.md](ANIMATIONS.md) for transition and animation best practices

## Examples

**Business Presentation**:
```
Input: "Create a Q3 financial review presentation for the board"
Output: 10-slide presentation with executive summary, financial metrics, 
performance analysis, challenges, opportunities, and next steps
```

**Sales Pitch**:
```
Input: "Make a sales deck for our new product launch"
Output: 8-slide presentation with problem, solution, benefits, 
case studies, pricing, and call-to-action
```

## Reference Materials

**Template Library**: [TEMPLATES.md](TEMPLATES.md) - Slide layouts for different presentation types

**Design System**: [BRANDING.md](BRANDING.md) - Colors, fonts, and branding guidelines  

**Animation Guidelines**: [ANIMATIONS.md](ANIMATIONS.md) - Professional transition effects

## Scripts

- `scripts/generate_presentation.py` - Main presentation generation script
- `scripts/analyze_prompt.py` - Content analysis and structure planning
- `scripts/apply_design.py` - Design and formatting application
