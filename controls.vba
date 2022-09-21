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


' It's very [unfortunate] that Word doesn't provide the value
Function ValueFromCc(cc As ContentControl) As String
    If cc.Type = wdContentControlDropdownList Then
        Dim i As Long
        With cc
            For i = 1 To .DropdownListEntries.Count
                If .DropdownListEntries(i).Text = .Range.Text Then _
                    ValueFromCc = .DropdownListEntries(i).Value
            Next i
        End With
    ElseIf cc.Type = wdContentControlCheckBox Then
        If cc.Checked Then
            ValueFromCc = "yes"
        Else
            ValueFromCc = "no"
        End If
    Else
        MsgBox "Unexpected content control type: " & Str(cc.Type)
    End If
End Function


Private Function FBkmkTextMatch(doc As Document, sBkmk As String, sText As String) As Boolean
    'MsgBox "Compare" & vbCrLf & "[" & sText & "] with" & vbCrLf & "[" & doc.Bookmarks(sBkmk).Range.Text & "]"
    If Not doc.Bookmarks.Exists(sBkmk) Then
        FBkmkTextMatch = False
    Else
        FBkmkTextMatch = doc.Bookmarks(sBkmk).Range.Text = sText
    End If
End Function


Private Function OnExit_Delta(ByVal cc As ContentControl) As Boolean
    OnExit_Delta = True ' default success, else the user may get stuck
    Dim doc As Document
    Set doc = ActiveDocument
    Dim sPrefix As String
    sPrefix = Mid(cc.Tag, 1 + 1) & "_"
    
    ' find the next content control in the document
    ' sadly, Range.ContentControls includes controls outside the range!
    Dim ccText As ContentControl
    For Each ccText In doc.ContentControls
        If ccText.Range.Start > cc.Range.End Then
            Exit For
        End If
    Next
    ' sanity check
    If ccText.BuildingBlockType <> wdContentControlRichText Then
        MsgBox "Expected rich text control to follow dropdown!"
        Exit Function
    End If

    ' don't know the old value, so need to check against all for change...
    Dim sVal, sText As String, fChanged As Boolean
    sVal = ValueFromCc(cc)
    sText = ccText.Range.Text
    Dim entry As ContentControlListEntry
    fChanged = Not ccText.ShowingPlaceholderText
    If fChanged Then
        If cc.Type = wdContentControlDropdownList Then
            For Each entry In cc.DropdownListEntries
                If FBkmkTextMatch(doc, sPrefix & entry.Value, sText) Then
                    fChanged = False
                    Exit For
                End If
            Next
        ElseIf cc.Type = wdContentControlCheckBox Then
            If FBkmkTextMatch(doc, sPrefix & "yes", sText) Then
                fChanged = False
            ElseIf FBkmkTextMatch(doc, sPrefix & "no", sText) Then
                fChanged = False
            End If
        End If
    End If

    ' find replacement (bookmark content)
    Dim sName As String
    sName = sPrefix & sVal
    
    'TODO: another useful effect of saving the old value would be being able to do nothing if the sel (not the content) hadn't changed
    ' this also wouldn't blow away the copy buffer...
    ' might be better to data-bind it, but that would probably be annoying re: having to somehow make unique paths - no free copy
    
    If fChanged Then
        If MsgBox("Text has been edited; replace?", vbYesNo) <> vbYes Then
            ' TODO: to be able to select the old value, would have to stash it (hash by CC ID?)
            ' see, e.g., https://stackoverflow.com/questions/1309689/hash-table-associative-array-in-vba
            ' cancel the exit event? [then we can't exit sub... haha]
            ' entry.Select
            'OnExit_Delta = False ' problem here is that even going back to the original, at present, doesn't allow exiting!
            Exit Function
        End If
    End If
    
    If Not doc.Bookmarks.Exists(sName) Then
        If sVal = "" Or sVal = "no" Then
            ccText.Range.Text = ""
        Else
            ' only flag in debug; otherwise may be the bookmarks were removed
            ' after citechecking, which ought to be harmless when passing through it
            'ccText.Range.Text = """" & sName & """" & " bookmark not found."
            'ccText.Range.Font.ColorIndex = wdRed
        End If
    Else
        Dim bkmk As Bookmark
        Set bkmk = doc.Bookmarks(sName)

        Dim rngSav As Range
        Set rngSav = Selection.Range
        
        ccText.Range.Select
        bkmk.Range.Copy
        Selection.Paste
        
        rngSav.Select
    End If
End Function


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
