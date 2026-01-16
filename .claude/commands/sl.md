# Slide Generator (pptx)

Generate a PowerPoint file from a slide design markdown file.

## Input
$ARGUMENTS

## Options
- `[filename]`: Design file name (without path/extension) or full path
- `-t [template]`: Template name (default: genda)

## Directory Structure

```
./slides/
├── design/          # Markdown design files (.md)
├── output/          # Generated PowerPoint files (.pptx)
├── scripts/         # Python generation scripts
│   ├── generate_pptx.py           # Base generator class
│   └── generate_*.py              # Specific slide generators
└── templates/       # PowerPoint templates (.pptx)
```

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

   Use the base generator class from `./slides/scripts/generate_pptx.py` or create a new script.

   Key principles:
   - **Delete existing template slides** before adding new ones
   - **Extract colors from template theme** (don't hardcode)
   - **Respect template layout** (avoid right side copyright area)

   Run Python scripts using `uv run python`:
   ```bash
   uv run python slides/scripts/generate_pptx.py design_file.md
   ```

   Base generator features:
   ```python
   from slides.scripts.generate_pptx import SlideGenerator

   gen = SlideGenerator('./slides/templates/genda.pptx')
   gen.load_template()
   gen.delete_all_slides()  # Always delete existing slides

   # Colors are loaded from template theme automatically
   gen.colors.gold          # Primary accent
   gen.colors.dark_navy     # Primary dark
   gen.colors.sage_green    # Secondary accent
   # etc.

   gen.add_title_slide("Title", "Subtitle")
   gen.add_content_slide("Slide Title")
   gen.save('./slides/output/filename.pptx')
   ```

5. **Slide creation guidelines** (for genda.pptx template):

   **Template dimensions**: 20" x 11.25" (widescreen)

   **Content area margins**:
   - Left: 1.0 inch
   - Right: 2.0 inches (to avoid © GENDA Inc. copyright area)
   - Top: 2.8 inches (after title)
   - Bottom: 1.0 inch
   - Usable width: 17 inches (from 1.0" to 18.0")

   **Layouts**:
   - Layout 0 (TITLE_AND_BODY_1): Title slide
   - Layout 1 (TITLE_AND_BODY_1_1): Subtitle only
   - Layout 2 (TITLE_AND_BODY_1_1_1): Content slide - use for most slides

   **Placeholders** (Layout 2):
   - Type 1 (TITLE): Set slide title here
   - Type 4 (SUBTITLE): Clear this to avoid "サブタイトルを入力" showing

   **Visual elements**:
   - Use `add_rounded_box()` for colored boxes
   - Use `add_text_box()` for plain text
   - Use `center_left(total_width)` to center content horizontally

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
