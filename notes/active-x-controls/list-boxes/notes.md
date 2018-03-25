# MS Excel ActiveX Controls

## The `ListBox` Control

A picker menu which allows the user to choose one option from a provided list.

Reference: [documentation](https://msdn.microsoft.com/en-us/VBA/Language-Reference-VBA/articles/listbox-control).

### Initialization

"Developer" > "Insert" > "ActiveX Controls" > "List Box"

![a screenshot of a user selecting an option from a list-style menu](list-box.png)

### Properties

name | description
--- | ---
`ListFillRange` | The address of a range of cells to populate the control's list of selectable options.
`Value` | The name of the currently-selected list item.
`LinkedCell` | The address of a specified cell which is bidirectionally associated with control's value.

### Events

name | description
--- | ---
`Click` (default) | Triggers when an option is selected from the from the list.
`Change` | Triggers when an the control's value is changed. Triggers before the `Click` event in the control's event lifecycle.
