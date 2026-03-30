# Excel VBA + Local Ollama (GUI + Callable Automation)

This project gives you two ways to use a local Ollama model from Excel VBA:

1. **GUI workflow** (UserForm): pick ranges and run manually.
2. **Callable automation workflow** (VBA Sub): run repeatable batch jobs column-by-column (for example `B1:B56 -> B61`, then `C1:C56 -> C61`, etc.).

---

## Included files

- `OllamaSpillModule.bas` (core logic + callable automation procedures)
- `frmOllamaPrompt.frm` (GUI form, optional if you only want code-driven automation)

---

## Prerequisites

1. **Windows + Excel desktop** with VBA enabled.
2. **Ollama installed and running locally**.
3. At least one model pulled (example: `llama3`).
4. Ollama reachable at:
   - `http://127.0.0.1:11434`

Quick Ollama checks in a terminal:

```bash
ollama list
curl http://127.0.0.1:11434/api/tags
```

---

## Installation guide (detailed)

### 1) Download project files

Save these files locally:
- `OllamaSpillModule.bas`
- `frmOllamaPrompt.frm`

### 2) Open VBA editor

1. Open Excel.
2. Press **`ALT + F11`**.

### 3) Import module and form

1. In VBA editor, right-click your workbook project.
2. Click **Import File...**
3. Import `OllamaSpillModule.bas`.
4. Import `frmOllamaPrompt.frm`.

### 4) Save workbook as macro-enabled

Save as **`.xlsm`**.

### 5) Macro security settings

If required by your environment:
- Enable macros when opening workbook.
- Optionally enable **Trust access to the VBA project object model** (not usually required for the runtime flow in this repo, but some environments enforce it).

### 6) Verify installation

From VBA Immediate Window (`CTRL + G`), run:

```vb
? OLLAMA_BASE_URL
```

Expected result:
`http://127.0.0.1:11434`

---


## Can this be used without `frmOllamaPrompt.frm` (without GUI)?

Yes. The GUI form file is **optional**. You can import only `OllamaSpillModule.bas` and run everything through macros.

### Procedure (no GUI)

1. Import only `OllamaSpillModule.bas` into your workbook.
2. Do **not** import `frmOllamaPrompt.frm`.
3. Create a wrapper macro and run it (example below).

```vb
Sub RunWithoutGui()
    RunOllamaForSingleRange _
        inputRangeAddress:="A1:A20", _
        outputStartCellAddress:="B1", _
        modelName:="llama3"
End Sub
```

For repeated columns, use `RunOllamaColumnSeries(...)` exactly as shown in the automation section.

### Consequences of not importing the form

- `ShowOllamaPromptGui` will fail because `frmOllamaPrompt` does not exist in that workbook.
- You lose model/input/output pickers and status text on a form; all parameters must be passed in code.
- Core features still work: prompt building, calling Ollama, and writing/spilling output to cells.

If you want to keep one codebase that works with or without the form, avoid calling `ShowOllamaPromptGui` in your own macros unless the form is imported.

---
## User guide (GUI version)

### Launch

Run macro:

```vb
ShowOllamaPromptGui
```

### GUI fields

- **Input ranges (comma-separated)**  
  Example: `A1,A2:A4,D10`
- **Output start cell**  
  Example: `B1`
- **Model**  
  Loaded from `GET /api/tags`.

### What happens on Run

1. Input ranges are combined into one prompt (non-empty cells only, newline-separated).
2. Prompt is sent to Ollama `POST /api/generate` with `stream=false`.
3. Response is written starting at output cell.
4. If response is too long, it spills down into following rows.

### Spill behavior and text size handling

- Excel limit per cell: **32,767 chars**
- Safe chunk target used: **32,000 chars**
- Chunk split priority:
  1. newline
  2. space
  3. underscore
  4. hard split fallback

---

## User guide (callable automation version)

This is the additional installable version that can run **at the same time as the GUI version** (same module, separate entry points).

### New callable procedures

1. `RunOllamaForSingleRange(inputRangeAddress, outputStartCellAddress, modelName)`  
   Runs one input range and writes result to one output anchor.

2. `RunOllamaColumnSeries(modelName, startColumnIndex, inputStartRow, inputEndRow, outputRow, iterations)`  
   Repeats the process across columns.

3. `Example_RunBtoN()`  
   Ready-made sample:
   - Starts column **B** (`2`)
   - Input rows `1:56`
   - Output row `61`
   - Runs `13` iterations (B through N)

### Exact scenario you requested

To run:
- `B1:B56 -> B61`
- then `C1:C56 -> C61`
- and so on for a specified number of columns,

use:

```vb
Sub RunMyBatch()
    RunOllamaColumnSeries _
        modelName:="llama3", _
        startColumnIndex:=2, _
        inputStartRow:=1, _
        inputEndRow:=56, _
        outputRow:=61, _
        iterations:=20
End Sub
```

`iterations:=20` means:
- B, C, D ... up to 20 columns total from B.

### How to call this from VBA

1. In VBA editor, Insert -> Module.
2. Paste your wrapper macro (like `RunMyBatch` above).
3. Press `F5` while cursor is inside the Sub.

---

## Error handling behavior

You will get clear runtime errors for:

- Empty model name
- Empty prompt
- Invalid range syntax
- Invalid row/column/iteration arguments
- Non-2xx HTTP responses from Ollama

---

## Troubleshooting

### “Format class contained in frmOllamaPrompt.frm is not supported in VBA”
- Cause: the form header must use the VBA UserForm class GUID, not VB6 `Begin VB.UserForm` syntax.
- Fix in this repo: `frmOllamaPrompt.frm` now uses:
  - `Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOllamaPrompt`
- If you still see the error, delete the failed form object from your VBA project and re-import the updated `frmOllamaPrompt.frm` file.

### “System error &H80004005 …” then “Out of memory” during form import
- Most common causes:
  1. The `.frm` file has incompatible or corrupted line endings/encoding for your Excel build.
  2. A previously failed partial form import is still in the VBA project.
- Fix:
  1. In the VBA editor, delete any partially imported `frmOllamaPrompt` object first.
  2. Re-import the latest `frmOllamaPrompt.frm` from this repo.
  3. If it still fails, open the `.frm` in a text editor and resave as **ANSI or UTF-8 without BOM** with **CRLF** line endings, then import again.
  4. If your organization blocks UserForms, skip the form entirely and use the non-GUI procedure above (`OllamaSpillModule.bas` only).

### “Model load failed”
- Confirm Ollama is running.
- Check `http://127.0.0.1:11434/api/tags` in browser or curl.

### “No response field found”
- Check whether Ollama returned an error payload.
- Verify model exists (`ollama list`).

### Macro runs but no useful output
- Confirm input cells are populated.
- Verify selected model can follow your prompt style.

---

## Optional customization

- Change default Ollama endpoint by editing:
  - `Public Const OLLAMA_BASE_URL`
- Adjust spill chunk size:
  - `SAFE_CELL_CHARS`

---

## Quick start summary

1. Import `.bas` + `.frm`.
2. Run `ShowOllamaPromptGui` for manual use.
3. Run `RunOllamaColumnSeries(...)` for repeatable column automation.
