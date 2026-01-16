# Slide Design Generator

Create a slide structure and save it as a markdown file for later pptx generation.

## Input
$ARGUMENTS

## Instructions

1. Parse the input to understand:
   - Presentation topic/purpose
   - Target audience (if mentioned)
   - Time constraint (if mentioned)
   - Any specific requirements

2. If critical information is missing, ask clarifying questions using AskUserQuestion tool:
   - What is the main purpose/goal?
   - Who is the audience?
   - How long is the presentation?

3. Design the slide structure:
   - Create a logical flow/storyline
   - Determine appropriate number of slides
   - Write title and key points for each slide
   - Consider the audience and time constraints

4. Generate markdown in the following format:

```markdown
---
title: [Presentation Title]
date: [YYYY-MM-DD]
audience: [Target Audience]
duration: [Expected Duration]
template: genda
---

# Slide Structure

## 1. [Slide Title]
- Key point 1
- Key point 2
- Key point 3

## 2. [Slide Title]
- Key point 1
- Key point 2

... (continue for all slides)

# Notes
[Any additional notes or considerations]
```

5. Save the file:
   - Directory: `./slides/design/`
   - Filename format: `YYYY-MM-DD_[slug].md` (slug = lowercase, hyphen-separated topic name)
   - Use Bash with Write tool to create the file

6. Output:
   - Show the proposed structure to the user
   - Confirm the file path where it was saved
   - Mention they can run `/sl` to generate the pptx

## Example Output

```
## Slide Structure (8 pages, ~15 min)

1. **Title** - Q1 Business Report
2. **Agenda** - Today's overview
3. **Executive Summary** - Key highlights
4. **Sales Performance** - Revenue & growth
5. **Product Updates** - New features launched
6. **Challenges** - Issues and learnings
7. **Q2 Plan** - Next quarter priorities
8. **Q&A** - Discussion

Saved to: `./slides/design/2025-01-16_q1-report.md`

Run `/sl` to generate the pptx file.
```
