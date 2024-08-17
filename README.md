# Multitool
Excel one-liner calculator and topmost tool

### Excel Multitool includes the following tools:
- A one-line calculator which always stays on top of other windows (non-topmost) and can evaluate expressions copy-pasted from MS Word Equation;
- A topmost tool which makes any other application window topmost/non-topmost.

### How Excel Multitool appeared (skip if not interested):
I have to write a lot of equations in my reports. Thus, I often use MS Word Equation editor. Windows Calculator doesn't calculate copy-pasted expressions in most cases. Neither MS Excel does. Typing expressions again in Windows Calculator and/or MS Excel is error-prone and not very comfortable and sometimes is time-consuming (since, large expressions in linear form doesn't look user-friendly). Installing third-party applications on a corporate PC isn't allowed. Thus, I created Calculator tool that solves those problems. Moreover, it doesn't take much space and always stays on top of other windows.

Quite often, I have a lot of windows open. Sometimes, I need to jump between them. Resizing some windows to a quarter of window, for example, doesn't help (or doesn't work for some applications). Thus, I came up with an idea of topmost windows (unfortunately, Windows doesn't provide this option for an ordinary user, like Linux does). And, again, installing third-party applications on a corporate PC isn't allowed. Therefore, I wrote a piece of code that solves this problem too.

### How to make this code work:
1. Create a macro-enabled Excel workbook (`.xlsb`, `.xlsm`, or `.xls`).
1. Open the VBA editor.
1. Import the project files.
1. In order to run the code more easily, create one or two button controls (whichever tool you need) in a Worksheet. Link one button to `modMain.Calc_Click` macro and the other to `modMain.Top_Click`.
1. Save the workbook.

### How to use Calculator:
- Type or paste an expression (a numerical part only) directly into the input field (refer to Calculator supported features).
- Format the result as required. You can:
  - Show/hide leading plus sign.
  - Show/hide trailing zeros.
  - Control number of digits after decimal comma.
    **Note:** The result in the title bar remains unformatted.
- If needed, you may copy the formatted result from the output field.
- Topmost tool button is embedded into Calculator window also (refer to usage instructions below).

### Calculator supported features:
- You can use available Excel functions, also.
- You can copy-paste an MS Word Equation expression (refer to Supported MS Word Equation functions).

### Supported MS Word Equation functions:
- Fractions.
- Raise to power (superscript).
- Radicals (square and cubic roots, radical with degree).
- A pair of round `()`, square `[]`, and figure `{}` brackets are all converted to round brackets.
- A pair of single vertical pipes `|` are converted to Excel `ABS()` function.
- A pair of floor `⌊⌋` and ceiling `⌈⌉` brackets are converted to Excel `ROUNDDOWN()` and `ROUNDUP()` functions and rounds the value to a whole number.
- `sin`, `cos`, `tan`, `cot`, `sec`, `csc`, `sinh`, `cosh`, `tanh`, `coth`, `sech`, and `csch` functions can be used in conjunction with the degree sign `°` and/or a pair of round `()`, square `[]`, and/or figure `{}` brackets.
- `min` and `max` functions can be used in conjunction with the semicolon `;` (used as values separator) and/or a pair of round `()`, square `[]`, and/or figure `{}` brackets.
**Notes:**
  - The semicolon `;` is used as the separator for a function arguments.
  - The comma `,` is used as the thousands separator and is removed.
  - The degree sign `°` tells Calculator to convert the input value from degrees to radians.
  - Breaking and non-breaking spaces are removed.
  - `π` is converted to Excel `PI()` function.
  - The square root `√` is converted to Excel `SQRT()` function.
  - The multiplication sign `×` is converted to the asterisk sign `*`.
  - The fraction slash `⁄` is converted to the forward slash `/`.
  - The hyphen `‐`, the figure dash `‒`, the en dash `–`, the em dash `—`, and the horizontal bar `―` are all converted to the hyphen-minus sign `-`.
  - The Unicode function application symbol is removed.

### How to use Topmost tool:
- Click a topmost run button and within 3 seconds activate the window you want to make topmost/non-topmost.
  **Note:** If a window is topmost, it will become non-topmost and vice versa.

### Known issues:
- Topmost tool may not work with some remote application windows (e.g., it may not work with Citrix applications).
- When activating Calculator, the Excel window gets completely transparent and, sometimes, the window's shadow can be visible, so, a user may be clicking in that transparent window.

### Examples (highlighted portion is what should be copy-pasted into Calculator):
![examples](https://user-images.githubusercontent.com/18612775/209526584-48f653a0-0f0c-44ec-b8ac-740bbad392c3.png)
