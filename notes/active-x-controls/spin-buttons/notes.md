# MS Excel ActiveX Controls

## The `SpinButton` Control

A set of arrows which allow the user to increment or decrement an integer value.

Reference: [documentation](https://msdn.microsoft.com/en-us/VBA/Language-Reference-VBA/articles/spinbutton-control).

### Initialization

"Developer" > "Insert" > "ActiveX Controls" > "Spin Button"

![a screenshot of a pair of buttons: a left arrow and a right arrow.](spin-button.png)

### Properties

name | description
--- | ---
`Orientation` | Specifies whether the control's pair of buttons should be arranged horizontally (i.e. left arrow and right arrow) or vertically (i.e. up arrow and down arrow) (default: automatically horizontal).
`Min` | The minimum allowable integer value, inclusive (default: 0).
`Max` | The maximum allowable integer value, inclusive (default: 100).
`SmallChange` | The absolute value numeric difference to be applied when the integer value is incremented or decremented via the control's buttons.
`Value` | The control's current integer value.
`LinkedCell` | The address of a specified cell which is bidirectionally associated with control's value.

### Events

name | description
--- | ---
`Change` (default) | Triggers when the control's value is changed.
