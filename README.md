# Excel VBA + Local Ollama GUI Macro

This repository contains an import-ready VBA module and UserForm for Excel.

## Files

- `OllamaSpillModule.bas`
- `frmOllamaPrompt.frm`

## What it does

1. Takes one or more input cell/range references (comma-separated) and concatenates them into one prompt.
2. Lets you choose a local Ollama model by loading `/api/tags`.
3. Sends the prompt to `/api/generate` with `stream=false`.
4. Writes response text to the output start cell and spills downward when needed.
5. Splits large text logically using this priority:
   - nearest newline
   - nearest space
   - nearest underscore
   - hard split at max size if none exist

## Usage

1. Open Excel and press `ALT+F11`.
2. Import `OllamaSpillModule.bas`.
3. Import `frmOllamaPrompt.frm`.
4. Ensure **Trust access to the VBA project object model** is enabled if required in your environment.
5. Run macro: `ShowOllamaPromptGui`.

## Notes

- Assumes Ollama is running locally at `http://127.0.0.1:11434`.
- Excel cell text limit is 32,767 chars; this implementation uses ~32,000-char safe chunks.
