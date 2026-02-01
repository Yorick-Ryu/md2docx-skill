# md2docx-skill

A Cursor/Claude Agent Skill for converting Markdown to professionally formatted Word (DOCX) documents.

## Overview

This skill enables AI coding assistants (Cursor, Claude Code, etc.) to convert Markdown content to Word documents using the [DeepShare](https://ds.rick216.cn) conversion API. It provides a seamless way to export documentation, papers, and other content to DOCX format.

## Features

- **Two Conversion Modes**: URL mode for cloud environments, File mode for local environments
- **Multiple Templates**: Support for Chinese and English document templates
- **Academic Paper Support**: Specialized templates for papers and thesis
- **Math Formula Support**: LaTeX syntax (`$...$` and `$$...$$`)
- **Flexible API Key Configuration**: Environment variable, Skill variable, or trial key

## Installation

### For Cursor

1. Copy the `skills/md2docx` folder to your project's `.cursor/skills/` directory:

```bash
cp -r skills/md2docx /path/to/your/project/.cursor/skills/
```

2. (Optional) Set your API key:

```bash
export DEEP_SHARE_API_KEY="your_api_key_here"
```

### For Claude Code

1. Copy the `skills/md2docx` folder to your project's `.claude/skills/` directory:

```bash
cp -r skills/md2docx /path/to/your/project/.claude/skills/
```

2. (Optional) Set your API key as described above.

## Project Structure

```
md2docx-skill/
├── README.md
└── skills/
    └── md2docx/
        ├── SKILL.md          # Skill definition file
        └── scripts/
            └── convert.py    # Conversion script
```

## Usage

Once installed, the AI assistant can convert Markdown to Word documents when requested:

**Example prompts:**
- "Convert this markdown to Word"
- "Export my README.md to DOCX"
- "Create a Word document from these notes"

### Command Line Usage

The conversion script can also be used directly:

```bash
# URL mode (returns download URL)
python scripts/convert.py input.md --url

# File mode (saves file locally)
python scripts/convert.py input.md --file

# With template and language options
python scripts/convert.py paper.md --file --template 论文 --language zh
python scripts/convert.py thesis.md --file --template thesis --language en

# With custom output directory
python scripts/convert.py doc.md --file --output ./output
```

### Mode Selection

| Scenario | Mode | Description |
|----------|------|-------------|
| Cloud/Remote environment | `--url` | Returns a download URL |
| Local environment | `--file` | Saves DOCX file directly |

## Available Templates

### Chinese Templates (`language: zh`)

- `templates` - General purpose
- `论文` - Academic paper
- `论文-首行不缩进` - Paper without first-line indent
- `论文-标题加粗` - Paper with bold headings

### English Templates (`language: en`)

- `templates` - General purpose
- `article` - Article/report style
- `thesis` - Academic thesis

## API Key Configuration

The skill uses API keys in the following priority order:

1. **Environment Variable** (Highest Priority)
   ```bash
   export DEEP_SHARE_API_KEY="your_api_key_here"
   ```

2. **Skill Variable** (Medium Priority)
   Edit the `api_key` field in `SKILL.md`:
   ```yaml
   ---
   name: md2docx
   api_key: "your_api_key_here"
   ---
   ```

3. **Trial Key** (Fallback)
   A limited trial key is included for testing.

> ⚠️ **Note**: The trial key has limited quota. For production use, purchase an API key at: https://ds.rick216.cn/purchase

## Requirements

- Python 3.7+
- `requests` library

Install dependencies:

```bash
pip install requests
```

## API Endpoints

| Mode | Endpoint | Response |
|------|----------|----------|
| URL | `https://api.deepshare.app/convert-text-to-url` | `{"url": "..."}` |
| File | `https://api.deepshare.app/convert-text` | DOCX binary |

## Markdown Guidelines

For best conversion results:

- Use `#` syntax for headers
- Use `-` or `1.` for lists
- Use triple backticks for code blocks
- Use `$...$` for inline math, `$$...$$` for block math
- Use publicly accessible `https://` URLs for images
- Keep content under 10MB

## Error Handling

| Error Code | Meaning | Solution |
|------------|---------|----------|
| 401 | Invalid API key | Check your API key |
| 403 | Quota exceeded | Purchase more credits |
| 413 | Content too large | Reduce content size (max 10MB) |
| 500 | Server error | Retry later |

## License

MIT

## Links

- **Purchase API Key**: https://ds.rick216.cn/purchase
- **API Documentation**: https://api.deepshare.app/docs
