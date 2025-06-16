The procedure to add the macro is the following:

1. Enable Developer Tab (if not already visible)
- In PowerPoint, go to File → Options.
- Click Customize Ribbon.
- On the right, check Developer, then click OK.

2. Open the VBA Editor
- Press Alt + F11 or click Developer → Visual Basic.

3. Insert the Progress Bar Macro
- In the VBA editor, go to Insert → Module.
- Paste the code below

4. Run the Macro
- In the VBA window, click anywhere inside the macro code.
- Press F5 or go to Run → Run Sub/UserForm.

How to change macro:

Where can I change the width of the bar:

Absolutely! Here's a **guide to the key parts of the macro** so you can easily change:

---

### 🎨 1. **Bar Width (Highlight Line Under Text)**

Look for this part of the code:

```vba
Set shp = sld.Shapes.AddShape(msoShapeRectangle, _
    Left:=labelLeft + 15, Top:=bannerTop + barYOffset, Width:=70, Height:=barHeight)
```

#### What to change:

* `Width:=70` → make the bar longer or shorter.
* `Left:=labelLeft + 15` → adjust horizontal position (e.g., center it by reducing or increasing the offset).

💡 **Tip:** To make the bar the same width as the label, you could do something like:

```vba
Width:=100, Left:=labelLeft
```

---

### 🎨 2. **Colors**

You define all colors at the top under `STYLE SETTINGS`:

```vba
Dim textColorActive As Long: textColorActive = RGB(0, 0, 0)        ' Black text
Dim textColorInactive As Long: textColorInactive = RGB(150, 150, 150)  ' Gray text
Dim barColorActive As Long: barColorActive = textColorActive       ' Black bar
Dim barColorInactive As Long: barColorInactive = textColorInactive ' Gray bar
```

#### Customize like this:

* 🔵 Blue: `RGB(0, 176, 240)`
* 🟢 Green: `RGB(0, 200, 0)`
* 🔴 Red: `RGB(200, 0, 0)`
* 🟡 Yellow: `RGB(255, 200, 0)`

You can also set different colors for the text and the underline bars if you'd like more contrast.

---

### 🅰️ 3. **Font (Police) Style of Text**

This section controls **font size, weight, color, and alignment**:

```vba
With shp
    .TextFrame2.TextRange.Text = sectionNames(i)
    .TextFrame2.TextRange.Font.Size = 14
    .TextFrame2.TextRange.Font.Bold = isActive
    .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = IIf(isActive, textColorActive, textColorInactive)
    .TextFrame2.TextRange.Font.Name = "Calibri" ' ← You can add this line to change font
    .TextFrame2.TextRange.ParagraphFormat.Alignment = ppAlignCenter
    .TextFrame2.VerticalAnchor = msoAnchorMiddle
End With
```

#### What to change:

* `.Font.Size = 14` → any number (e.g., `16`)
* `.Font.Bold = isActive` → you can change to `msoFalse` if you don’t want bold
* `.Font.Name = "Arial"` → set your preferred font name

---

### 🧪 Example: Use Arial 16pt and Blue Highlight

Change:

```vba
.TextFrame2.TextRange.Font.Size = 16
.TextFrame2.TextRange.Font.Name = "Arial"
```

And at the top:

```vba
Dim barColorActive As Long: barColorActive = RGB(0, 176, 240) ' Blue bar
```

---

Let me know if you'd like:

* Bars to **fade in or animate**.
* Labels to be **centered globally** rather than equally spaced.
* Fonts or bars to **match the slide master**.

Happy to help!

