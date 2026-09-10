Option Explicit

' StandardizeTitles
' Loops through every slide in the active presentation and applies
' consistent formatting to the title placeholder.
' Written for a consulting team that needs decks to follow one visual standard.

Sub StandardizeTitles()

    Dim sld As Slide
    Dim shp As Shape
    Dim titleCount As Integer
    Dim skippedCount As Integer

    titleCount = 0
    skippedCount = 0

    ' Guard: make sure a presentation is actually open
    If Presentations.Count = 0 Then
        MsgBox "No presentation is open.", vbExclamation
        Exit Sub
    End If

    For Each sld In ActivePresentation.Slides

        ' Not every slide has a title placeholder, so check first
        If sld.Shapes.HasTitle Then

            Set shp = sld.Shapes.Title

            With shp.TextFrame.TextRange.Font
                .Name = "Arial"
                .Size = 28
                .Bold = msoTrue
                .Color.RGB = RGB(0, 60, 120)   ' corporate dark blue
            End With

            ' Align titles consistently across the deck
            shp.TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignLeft
            shp.Left = 50
            shp.Top = 30

            titleCount = titleCount + 1

        Else
            skippedCount = skippedCount + 1
        End If

    Next sld

    MsgBox "Formatted " & titleCount & " titles." & vbCrLf & _
           "Skipped " & skippedCount & " slides with no title.", vbInformation

End Sub
