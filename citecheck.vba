' David B. Robins for <law journal>, 20220815
'
' I don't enjoy writing VBA, but it's what's here. How popular would this be if users had to install Python first?
' Granted, a cloud solution might be doable, but that seems more difficult on the whole (only Word can parse Word...).

Private Sub PasteAfterColon(rng As Range)
    ichText = rng.Start + InStr(rng.Text, ": ") + 1
    Selection.SetRange ichText, ichText
    Selection.Paste
End Sub

Sub CreateCiteCheck()
    ' find the article document: use the first with a selection with footnotes
    Dim docText As Document
    For Each docOpen In Documents
        If docOpen.ActiveWindow.Selection.Type = wdSelectionNormal Then
            If docOpen.ActiveWindow.Selection.Footnotes.Count > 0 Then
                Set docText = docOpen
                GoTo LFound
            End If
        End If
    Next
LFound:
    If docText Is Nothing Then
        MsgBox "Please select text (with at least one footnote flag) in your article."
        End
    End If
    docText.Activate
    
    ' find the cite-checking report template (open document with "CC" and "Report" in the name)
    Set docRpt = Nothing
    For Each docOpen In Documents
        If InStr(docOpen.Name, "CC") > 0 And InStr(docOpen.Name, "Report") > 0 Then
            If docOpen = docText Then
                MsgBox "Please select the article text before running (do not select in the CC Report document)."
                End
            End If
            Set docRpt = docOpen
        End If
    Next
    If docRpt Is Nothing Then
        MsgBox "Please open or save a document with 'CC Report' in the file name."
        End
    End If
    
    ' find the footnote offset (skip non-numbered footnotes)
    If Selection.Footnotes.Count = 0 Then
        MsgBox "Degenerate selection (no footnotes), aborting.", vbMsgExclamation Or vbOKOnly
        End
    End If
    cftnSkip = 0
    For Each ftn In docText.Footnotes
        ' note this is not a space, it's a special character used for the footnote number
        If ftn.Range.Characters(1).Text <> "" Then
            'MsgBox "Skipping: " & ftn.Range.Text
            cftnSkip = cftnSkip + 1
        Else
            Exit For
        End If
    Next

    ' save relevant selection information before moving (we know it's a linear sel)
    ichFirst = Selection.Start
    ichMac = Selection.End
    cftnSel = Selection.Footnotes.Count
    iftnFirst = Selection.Footnotes(1).Index
    iftnLast = Selection.Footnotes(cftnSel).Index

    ' find 2-row table in the report doc (for cloning and editing)
    If docRpt.Tables.Count <> 1 Then
        MsgBox "Expected single table in CC Report."
        End
    End If
    Set tblRpt = docRpt.Tables(1)
    If tblRpt.Rows.Count <> 2 Then
        MsgBox "Expected single 2-row table in CC Report."
        End
    End If
    docRpt.Activate

    ' loop over the footnotes in the range
    For iftn = iftnFirst To iftnLast
        Set ftn = docText.Footnotes(iftn)
        ' find the range of text to copy
        Set rngText = docText.Range(ichFirst, ftn.Reference.Start)
        ichFirst = ftn.Reference.End
        
        ' find the footnote text to copy
        ichFtnFirst = ftn.Range.Start
        ichFtnMac = ftn.Range.End
        ich = 1
        While ichFtnFirst < ichFtnMac
            Set ch = ftn.Range.Characters(ich)
            ich = ich + 1
            If Asc(ch) > 32 And ch <> "." Then
                GoTo LBreak
            End If
            ichFtnFirst = ichFtnFirst + 1
        Wend
LBreak:
        Set rngFtn = ftn.Range.Duplicate
        rngFtn.Start = ichFtnFirst
        rngFtn.End = ichFtnMac
        
        ' clone the last 2 table rows
        crow = tblRpt.Rows.Count
        docRpt.Range(tblRpt.Rows(crow - 1).Range.Start, tblRpt.Rows(crow).Range.End).Copy
        ichTableEnd = tblRpt.Range.End
        Selection.SetRange ichTableEnd, ichTableEnd
        Selection.Paste
        
        ' insert the footnote number (2x)
        tblRpt.Rows(crow - 1).Cells(1).Range.Text = Str(iftn - cftnSkip)
        tblRpt.Rows(crow).Cells(1).Range.Text = Str(iftn - cftnSkip)
        
        ' copy body text after "TEXT: " in odd row
        rngText.Copy
        PasteAfterColon tblRpt.Rows(crow - 1).Cells(2).Range
                        
        ' copy footnote text after "ENTIRE ORIGINAL CITATION" in even row
        rngFtn.Copy
        PasteAfterColon tblRpt.Rows(crow).Cells(2).Range
        
        ' find range for subpart
        Dim rngSub As Range
        Set rngSub = tblRpt.Rows(crow).Cells(2).Range.Duplicate
        Set fndSub = rngSub.Find
        fndSub.Execute "SUBPART 1: "
        If Not fndSub.Found Then
            MsgBox "Missing SUBPART 1:, aborting.", vbMsgExclamation Or vbOKOnly
            End
        End If
        rngSub.End = tblRpt.Rows(crow).Cells(2).Range.End - 1 ' cell mark
        'MsgBox "[" & rngSub.Text & "]"

        ' break string cites (at semicolon, "dumb" for now; may want to consider periods too?)
        ichStart = rngFtn.Start
        Do
            Dim rngSplit As Range
            Set rngSplit = rngFtn.Duplicate ' because find result clobbers
            rngSplit.Start = ichStart
            Dim fnd As Find
            Set fnd = rngSplit.Find
            
            fnd.Execute "; "
            If fnd.Found Then
                ichNext = rngSplit.End
                rngSplit.Start = ichStart
                rngSplit.End = ichNext - 2
                ichStart = ichNext
            End If
            
            Dim rngNext As Range
            If fnd.Found Then
'                ' paste in another subpart
                Selection.SetRange rngSub.End, rngSub.End
                Selection.TypeText vbCrLf & vbCrLf
                rngSub.Copy
                Selection.Paste

                Set rngNext = rngSub.Duplicate
                rngNext.Start = rngSub.End + 2 ' newlines
                rngNext.End = tblRpt.Rows(crow).Cells(2).Range.End - 1 ' cell mark

                ' bump the count [not dealing with double digits; be insane on your own time]
                Dim rngI As Range
                Set rngI = rngNext.Duplicate
                rngI.Start = rngI.Start + 8 ' subpart number
                rngI.End = rngI.Start + 1
                rngI.Text = CStr(CInt(rngI.Text) + 1)
            End If
            
            rngSplit.Copy
            'MsgBox "[" & rngSplit.Text & "]"
            
            PasteAfterColon rngSub
            
            Set rngSub = rngNext
            If Not fnd.Found Then
                Exit Do
            End If
        Loop
    Next

    ' always delete the last table row (there's no final footnote, just a possibility of text)
    crow = tblRpt.Rows.Count
    tblRpt.Rows(crow).Delete
    
    ' if there's extra text, add it to the last text table row, else delete it
    If ichFirst < ichMac Then
        Set rngText = docText.Range(ichFirst, ichMac)
        rngText.Copy
        PasteAfterColon tblRpt.Rows(crow - 1).Cells(2).Range
    Else
        tblRpt.Rows(crow - 1).Delete
    End If
End Sub
