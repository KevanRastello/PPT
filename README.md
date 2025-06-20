Tired of wasting time on your PowerPoint? You probably need a macro!

# Implement the macro

#### 🧰 1. **Enable Developer Tab (if not already visible)**

1. In PowerPoint, go to **File → Options**.
2. Click **Customize Ribbon**.
3. On the right, check **Developer**, then click OK.

---

#### 💻 2. **Open the VBA Editor**

1. Press `Alt + F11` or click **Developer → Visual Basic**.

---

#### 📦 3. **Insert the Progress Bar Macro**

1. In the VBA editor, go to **Insert → Module**.
2. Paste this code:

---

#### ▶️ 4. **Run the Macro**

1. In the VBA window, click anywhere inside the macro code.
2. Press `F5` or go to **Run → Run Sub/UserForm**.

---

# How to Customize the Dynamic Banner Macro

| Change                   | Variable                            | Example                       |
| ------------------------ | ----------------------------------- | ----------------------------- |
| Font size                | `fontSizeText`                      | `16`                          |
| Font name                | `fontName`                          | `"Calibri"`                   |
| Text vertical position   | `bannerTop`                         | `5`                           |
| Space between text & bar | `barYOffset`                        | `5`                           |
| Bar thickness            | `barHeight`                         | `2`                           |
| Text active color        | `textColorActive`                   | `RGB(0,0,0)`                  |
| Text inactive color      | `textColorInactive`                 | `DarkerRGB(255,255,255,0.15)` |
| Bar active color         | `barColorActive`                    | same as active text           |
| Bar inactive color       | `barColorInactive`                  | same as inactive text         |
| Color inversion          | `invertedSlides`                    | `T` or `F` and specify slide numbers |
| Line vs. rectangle bar   | Use `AddLine` instead of `AddShape` |                               |

# How to Customize the Dynamic Numbering Macro

| Feature                          | Description                                              |
| -------------------------------- | -------------------------------------------------------- |
| `startSlide`, `endSlide`         | Define the slide range where slide numbers appear.       |
| `invertColorSlides = Array(...)` | Set which slide numbers should use white text.           |
| Text formatting                  | Font name, size, and alignment are applied consistently. |
| Positioning                      | Bottom-right placement with margins.                     |
| Text fit                         | Automatically sizes the box so numbers stay on one line. |
| Locked                           | Shape is locked to avoid accidental edits.               |

---

I tried to make the macro more modular and easier to edit by centralizing user-defined!

