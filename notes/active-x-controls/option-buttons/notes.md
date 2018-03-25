# MS Excel ActiveX Controls

## The `OptionButton` Control

A selectable circle (a.k.a. "radio button") belonging to a specified group from which only one may be selected at any given time.

Reference: [documentation](https://msdn.microsoft.com/en-us/VBA/Language-Reference-VBA/articles/optionbutton-control).

### Initialization

For each radio button: "Developer" > "Insert" > "ActiveX Controls" > "Option Button".

![a screenshot depicting one of four selected option buttons](option-button-1.png)

### Properties

name | description
--- | ---
`Caption` | A human-friendly name for the selectable option.
`GroupName` | Associates the control with a logical grouping of one or more controls (default: "Sheet1").
`Value` | The current state of the control (i.e. `True` if selected, otherwise `False`).
`LinkedCell` | The address of a specified cell which is bidirectionally associated with control's value.

### Events

name | description
--- | ---
`Click` (default) | Triggers when an option is selected.
`Change` | Triggers when an the control's value is changed. Triggers before the `Click` event in the control's event lifecycle.
