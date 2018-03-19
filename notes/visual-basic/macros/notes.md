# Visual Basic Programming

## Macros

### Recording Macros

  1. Click "Developer" > "Record Macro", then assign it a unique name and optional description.
  2. Manually perform one or more actions via the front-end MS Excel interface.
  3. Click "Developer" > "Stop Recording".

![A screenshot of the "Code" section of the "Developer" menu, where the macro-related buttons exist.](recording-a-macro.png)

### Executing Macros

Execute a macro manually by selecting it from the "Developer" > "Macros" menu and pressing "Run". 

![A screenshot of the "Macros" window.](playing-a-macro.png)

Or execite it programmatically in VBA by using the keyword `Call` followed by the name of the macro. For example:

```vba
Call Macro1
```