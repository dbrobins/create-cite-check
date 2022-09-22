' This code goes in the "ThisDocument" module under "Microsoft Word Objects."
' It handles content control exit event (best we can do for selection change)
' and populates a following rich edit control according to the dropdown selection,
' which identifies a bookmark.

Option Explicit


Private Sub OnExit_fCorrect(ByVal cc As ContentControl)
    Dim rngSav As Range
    Set rngSav = Selection.Range
    
    ' find the range of the (logical) line (i.e., paragraph)
    cc.Range.Select
    Selection.Collapse wdCollapseEnd
    Selection.MoveStart wdCharacter, 1 ' otherwise colors the checkbox (sdt close?)
    Selection.MoveEnd wdParagraph      ' wdLine is the DL
    Selection.MoveEnd wdCharacter, -1  ' back off paragraph mark
    
    ' set color
    Dim ci As WdColorIndex
    Dim chSym As String
    If cc.Checked Then
        ci = wdAuto
        chSym = "*"
    Else
        ci = wdRed
        chSym = "-"
    End If
    Selection.Range.Font.ColorIndex = ci
    
    ' find symbol (allow for space)
    Dim rngSym As Range
    Set rngSym = ActiveDocument.Range(cc.Range.End + 1, cc.Range.End + 2) ' skip the sdt close
    If rngSym.Text = " " Then
        rngSym.SetRange rngSym.Start + 1, rngSym.End + 1
    End If
    
    ' update symbol
    If rngSym.Text = "*" Or rngSym.Text = "-" Then
        rngSym.Text = chSym
    End If ' if we don't find an acceptable symbol, don't do anything
    
    rngSav.Select
End Sub


Private Sub Document_ContentControlOnExit(ByVal cc As ContentControl, cancel As Boolean)
    'Exit Sub ' uncomment to turn off for submission
    
    If cc.Tag = "fCorrect" Then
        OnExit_fCorrect cc
        Exit Sub
    End If
    
    If Left(cc.Tag, 1) = "!" Then
        cancel = Not OnExit_Delta(cc)
        Exit Sub
    End If
End Sub

