# Development Notes

Goal: have a templating engine for Microsoft Word files. Avoid using existing
Rust libraries for DOCX manipulation (both poiscript and bokuweb) - instead do
non-destructive XML manipulation based on a few ground rules.