# MS Excel ActiveX Controls

## The `CheckBox` Control

A checkable box belonging to a specified group from which zero or more may be selected at any given time.

Reference: [documentation](https://msdn.microsoft.com/en-us/VBA/Language-Reference-VBA/articles/checkbox-control).

### Initialization

For each box: "Developer" > "Insert" > "ActiveX Controls" > "Check Box".

![a screenshot depicting two of four checked boxes](check-box.png)

### Properties

name | description
--- | ---
`Caption` | A human-friendly name for the checkable option.
`GroupName` | Associates the control with a logical grouping of one or more controls (default: "Sheet1").
`Value` | The current state of the control (i.e. `True` if checked, otherwise `False`).
`LinkedCell` | The address of a specified cell which is bidirectionally associated with control's value.

### Events

name | description
--- | ---
`Click` (default) | Triggers when an option is checked.
`Change` | Triggers when the control's value is changed. Triggers before the `Click` event in the control's event lifecycle.
