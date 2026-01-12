# Developer Guide – AutoDoc AI

## Architecture Overview
- Tkinter + ttkbootstrap UI
- OpenAI API wrapper with retry logic
- DOCX parsing and markdown-to-docx conversion
- Threaded worker with queue-based UI updates

## Key Files
- `main.py` – core application
- `templates.json` – AI system prompts
- `settings.json` – UI preferences

## Extending the App
- Add new profiles in `templates.json`
- Modify formatting in `parse_markdown_to_docx()`
- Adjust pricing in `MODEL_PRICING`