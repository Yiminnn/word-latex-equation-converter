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
