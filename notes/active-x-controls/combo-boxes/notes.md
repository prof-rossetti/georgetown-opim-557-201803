# MS Excel ActiveX Controls

## The `ComboBox` Control

A drop-down menu which allows the user to choose one option from a provided list.

Reference: [documentation](https://msdn.microsoft.com/en-us/VBA/Language-Reference-VBA/articles/combobox-control).

### Initialization

"Developer" > "Insert" > "ActiveX Controls" > "Combo Box"

![a screenshot of a user selecting an option from a drop-down menu](combo-box-1.png)

### Properties

name | description
--- | ---
`ListFillRange` | The address of a range of cells to populate the control's list of selectable options.
`Value` | The name of the currently-selected list item.
`LinkedCell` | The address of a specified cell which is bidirectionally associated with control's value.

### Events

name | description
--- | ---
`Change` (default) | Triggers when an option is selected from the drop-down.
