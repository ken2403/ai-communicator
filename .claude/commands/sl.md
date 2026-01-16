# Slide Generator (pptx)

Generate a PowerPoint file from a slide design markdown file.

## Input
$ARGUMENTS

## Options
- `[filename]`: Design file name (without path/extension) or full path
- `-t [template]`: Template name (default: genda)

## Instructions

1. **Parse arguments**:
   - Extract filename if provided
   - Extract template name if `-t` flag is used (default: `genda`)

2. **Determine the design file**:

   Priority order:
   a. If filename argument provided → use that file
   b. If no argument → check conversation context for recent `/sl-design` output
   c. If still none → find most recently modified `.md` file in `./slides/design/`

   Use Bash to find files:
   ```bash
   ls -t ./slides/design/*.md | head -1
   ```

3. **Read the design file**:
   - Use Read tool to get the markdown content
   - Parse the frontmatter (title, date, template override)
   - Parse each slide section (## numbered headings)

4. **Generate pptx using Python**:

   Create and run a Python script that:
   - Uses python-pptx library
   - Opens the template from `./slides/templates/[template].pptx`
   - Creates slides based on the design structure
   - Saves to `./slides/output/[same-name-as-design].pptx`

   Run Python scripts using `uv run python`:
   ```bash
   uv run python script.py
   ```

   Example Python script structure:
   ```python
   from pptx import Presentation
   from pptx.util import Inches, Pt

   # Load template
   prs = Presentation('./slides/templates/genda.pptx')

   # Add slides based on design
   # ... (parse markdown and create slides)

   # Save
   prs.save('./slides/output/filename.pptx')
   ```

5. **Slide creation guidelines** (for genda.pptx template):
   - Layout 0 (TITLE_AND_BODY_1): Title slide - placeholder 0=title, 1=subtitle, 2=subtitle2
   - Layout 1 (TITLE_AND_BODY_1_1): Subtitle only
   - Layout 2 (TITLE_AND_BODY_1_1_1): Content slide - placeholder 0=title, 1=body content

   Recommended usage:
   - First slide (title): Use Layout 0
   - Content slides: Use Layout 2 (has title + body)
   - Set title from `## N. Title` line
   - Add bullet points from `- item` lines to body placeholder
   - Keep text readable (appropriate font sizes)

6. **Output**:
   - Confirm successful generation
   - Show the output file path
   - Mention total slide count

## Example Usage

```bash
# Auto-detect from context
/sl

# Specify design file
/sl q1-report

# Specify design file and template
/sl q1-report -t minimal

# Full path
/sl ./slides/design/2025-01-16_q1-report.md
```

## Example Output

```
Generated: ./slides/output/2025-01-16_q1-report.pptx
- Template: genda
- Slides: 8 pages
- Source: ./slides/design/2025-01-16_q1-report.md
```

## Error Handling

- If python-pptx not installed: Run `uv add python-pptx`
- Always use `uv run python` to execute Python scripts
- If design file not found: List available files in ./slides/design/
- If template not found: List available templates in ./slides/templates/
