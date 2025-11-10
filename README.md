# Word LaTeX Equation Converter

A VBA macro for Microsoft Word to convert LaTeX equation syntax (e.g., `$$...$$` and `$..$`) into native Word equations. This is particularly useful for pasting content from LLM chats or other sources that use LaTeX math notation.

## Macro Code

The code for the macro is in the `ConvertLatexToWordEquations.vba` file. You can also copy it from here:

```vb
Sub ConvertLatexToWordEquations()
    Dim rng As Range
    Dim latexEq As String
    
    ' First: Convert $$...$$ (display equations)
    Set rng = ActiveDocument.Content
    With rng.Find
        .ClearFormatting
        .Text = "\$\$(*)\$\$"
        .MatchWildcards = True
        
        Do While .Execute
            latexEq = Mid(rng.Text, 3, Len(rng.Text) - 4)
            rng.Text = Trim(latexEq)
            rng.OMaths.Add(rng).OMaths(1).BuildUp
            rng.OMaths(1).Type = wdOMathDisplay
            
            rng.Collapse wdCollapseEnd
            rng.End = ActiveDocument.Content.End
        Loop
    End With
    
    ' Second: Convert $...$ (inline equations)
    Set rng = ActiveDocument.Content
    With rng.Find
        .ClearFormatting
        .Text = "\$([!\$]@)\$"
        .MatchWildcards = True
        
        Do While .Execute
            latexEq = Mid(rng.Text, 2, Len(rng.Text) - 2)
            rng.Text = Trim(latexEq)
            rng.OMaths.Add(rng).OMaths(1).BuildUp
            rng.OMaths(1).Type = wdOMathInline
            
            rng.Collapse wdCollapseEnd
            rng.End = ActiveDocument.Content.End
        Loop
    End With
    
    MsgBox "Conversion complete!", vbInformation
End Sub
```

## How to Use

1.  **Open Microsoft Word.**
2.  Press `Alt + F11` to open the **Visual Basic for Applications (VBA) editor**.
3.  In the VBA editor, look for the **Project Explorer** pane (usually on the left). If you don't see it, go to `View > Project Explorer` or press `Ctrl + R`.
4.  In the Project Explorer, find and select `Normal` (or the project for your current document). By adding the macro to `Normal.dotm`, it will be available in all your documents.
5.  Right-click on `Normal`, then select `Insert > Module`. A new module (e.g., `Module1`) will appear under a `Modules` folder.
6.  Double-click the new module to open the code window on the right.
7.  **Copy and paste the VBA code** from above into this code window.
8.  Save the macro by pressing `Ctrl + S` or going to `File > Save Normal`.
9.  Close the VBA editor by pressing `Alt + Q` or clicking the 'x'.

### Running the Macro

1.  Go to the `View` tab in the Word ribbon.
2.  Click the `Macros` button on the far right.
3.  Select `ConvertLatexToWordEquations` from the list.
4.  Click `Run`.

### (Optional) Add to Quick Access Toolbar for Easy Access

1.  Right-click the **Quick Access Toolbar** (the small icons at the very top-left of the Word window) and select `Customize Quick Access Toolbar...`.
2.  In the "Choose commands from:" dropdown, select `Macros`.
3.  Find and select `Normal.NewMacros.ConvertLatexToWordEquations` (the name might vary slightly).
4.  Click the `Add >>` button.
5.  (Optional) Click the `Modify...` button to choose a different icon for the macro.
6.  Click `OK`. The macro will now be available as a single-click button in your Quick Access Toolbar.

## To create a GitHub repository

1.  [Create a new repository](https://github.com/new) on GitHub.
2.  Initialize a Git repository in the `word-latex-equation-converter` directory:
    ```bash
    git init
    git add .
    git commit -m "Initial commit"
    git branch -M main
    git remote add origin https://github.com/YOUR_USERNAME/word-latex-equation-converter.git
    git push -u origin main
    ```
    (Replace `YOUR_USERNAME` with your GitHub username).
