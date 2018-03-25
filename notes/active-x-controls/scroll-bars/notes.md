# MS Excel ActiveX Controls

## The `ScrollBar` Control

A fluid scale which allows the user to increment or decrement an integer value.

Reference: [documentation](https://msdn.microsoft.com/en-us/VBA/Language-Reference-VBA/articles/scrollbar-control).

### Initialization

"Developer" > "Insert" > "ActiveX Controls" > "Scroll Bar"

![a screenshot of a horizontally-aligned scroll-bar.](scroll-bar.png)

### Properties

name | description
--- | ---
`Orientation` | Specifies whether the control should be arranged horizontally or vertically (default: automatically horizontal).
`Min` | The minimum allowable integer value, inclusive (default: 0).
`Max` | The maximum allowable integer value, inclusive (default: 32767 :smiley:).
`SmallChange` | The absolute value numeric difference to be applied when the integer value is incremented or decremented by clicking the control's arrows.
`LargeChange` | The absolute value numeric difference to be applied when the integer value is incremented or decremented by clicking somewhere between the current value and the control's arrows.
`Value` | The control's current integer value.
`LinkedCell` | The address of a specified cell which is bidirectionally associated with control's value.

### Events

name | description
--- | ---
`Change` (default) | Triggers when the control's value is changed.
