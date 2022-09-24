' David B. Robins for use with ILJ, 20220815
'
' Released under GPL 3; see https://github.com/dbrobins/create-cite-check

Option Explicit


Private Sub PasteAfterColon(rng As Range)
    Dim ichText As Long
    ichText = rng.Start + InStr(rng.Text, ": ") + 1
    Selection.SetRange ichText, ichText
    Selection.Paste
End Sub


Private Function FHasQuotation(s As String) As Boolean
    FHasQuotation = s Like "*“*”*" ' very limited; can't exclude, say, single words in quotes
End Function


' Returns best guess at source type (using dropdown values), empty string if no guess.
Private Function SourceType(rng As Range) As String
    Dim s As String
    s = rng.Text
    
    ' low-hanging fruit...
    If StrComp(Left(s, 3), "Id.", vbTextCompare) = 0 Then
        SourceType = "id"
    ElseIf InStr(s, ", supra ") > 0 Then
        SourceType = "supra"
    ElseIf InStr(s, ", infra ") > 0 Then ' internal cross-ref, need to add (not in cheat sheet)
        SourceType = "infra"
    End If
    ' leaving more complex types for later
End Function


Sub CreateCiteCheck()
    ' find the article document: use the first with a selection with footnotes
    Dim docText, docOpen As Document
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
    Dim docRpt As Document
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
        MsgBox "Degenerate selection (no footnotes), aborting.", vbExclamation Or vbOKOnly
        End
    End If
    Dim cftnSkip As Long
    Dim ftn As Footnote
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
    Dim ichFirst, ichMac, cftnSel, iftnFirst, iftnLast As Long
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
    Dim tblRpt As Table
    Set tblRpt = docRpt.Tables(1)
    If tblRpt.Rows.Count <> 2 Then
        MsgBox "Expected single 2-row table in CC Report."
        End
    End If
    docRpt.Activate

    Application.ScreenUpdating = False

    ' loop over the footnotes in the range
    Dim iftn As Long
    For iftn = iftnFirst To iftnLast
        Set ftn = docText.Footnotes(iftn)
        ' find the range of text to copy
        Dim rngText As Range
        Set rngText = docText.Range(ichFirst, ftn.Reference.Start)
        ichFirst = ftn.Reference.End
        
        ' find the footnote text to copy
        Dim ichFtnFirst, ichFtnMac, ich As Long
        ichFtnFirst = ftn.Range.Start
        ichFtnMac = ftn.Range.End
        ich = 1
        While ichFtnFirst < ichFtnMac
            Dim ch As Range
            Set ch = ftn.Range.Characters(ich)
            ich = ich + 1
            If Asc(ch) > 32 And ch <> "." Then
                GoTo LBreak
            End If
            ichFtnFirst = ichFtnFirst + 1
        Wend
LBreak:
        Dim rngFtn As Range
        Set rngFtn = ftn.Range.Duplicate
        rngFtn.SetRange ichFtnFirst, ichFtnMac
        
        ' clone the last 2 table rows
        Dim crow, ichTableEnd As Long
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
        
        ' [[ check text content
        
        Dim sText As String
        sText = rngText.Text
        If FHasQuotation(sText) Then
            Dim ccQuote As ContentControl
            Set ccQuote = CcNext(docRpt, tblRpt.Rows(crow - 1).Cells(2).Range.Start + Len(sText), "!quote")
            If Not ccQuote Is Nothing Then
                ccQuote.Checked = True
                Call OnExit_Delta(ccQuote)
            End If
        End If
        
        ' ]] check text content
        
        ' find range for subpart
        Dim rngSub As Range
        Set rngSub = tblRpt.Rows(crow).Cells(2).Range.Duplicate
        Dim fndSub As Find
        Set fndSub = rngSub.Find
        fndSub.Execute "SUBPART 1: "
        If Not fndSub.Found Then
            MsgBox "Missing SUBPART 1:, aborting.", vbExclamation Or vbOKOnly
            End
        End If
        rngSub.End = tblRpt.Rows(crow).Cells(2).Range.End - 1 ' cell mark
        'MsgBox "[" & rngSub.Text & "]"

        ' break string cites (at semicolon, "dumb" for now; may want to consider periods too?)
        Dim ichStart, ichNext As Long
        ichStart = rngFtn.Start
        Do
            Dim rngSplit As Range
            Set rngSplit = rngFtn.Duplicate ' because find result clobbers
            rngSplit.Start = ichStart
            Dim fnd As Find
            Set fnd = rngSplit.Find
            
            fnd.Execute "; "
            Dim fFound As Boolean
            fFound = fnd.Found
            If fFound Then
                ichNext = rngSplit.End
                rngSplit.Start = ichStart
                rngSplit.End = ichNext - 2 ' "; "
                ichStart = ichNext
                
                ' prevent sorcerer's apprentice on trailing ";"
                If ichNext >= rngFtn.End Then fFound = False
            End If
            
            Dim rngNext As Range
            If fFound Then
                ' paste in another subpart
                ' doing Selection.SetRange and .TypeText vbCrLf & vbCrLf used to work, but not just after a CC
                ' add replacing the .TypeText with two .InsertParagraphAfters to list of what doesn't work
                ' storing off the range object also doesn't work ("object or with not set"); crazily, this does
                docRpt.Range(rngSub.End, rngSub.End).InsertParagraphAfter
                docRpt.Range(rngSub.End, rngSub.End).InsertParagraphAfter
                
                rngSub.Copy
                Selection.SetRange rngSub.End + 2, rngSub.End + 2 ' newlines
                Selection.Paste

                Set rngNext = rngSub.Duplicate
                rngNext.Start = rngSub.End + 2 ' newlines
                rngNext.End = tblRpt.Rows(crow).Cells(2).Range.End - 1 ' cell mark

                ' bump the count
                Dim rngI As Range
                Set rngI = rngNext.Duplicate
                rngI.Start = rngI.Start + 8 ' subpart number
                rngI.End = rngI.Start + 1
                rngI.Text = CStr(CInt(rngI.Text) + 1)
            End If
            
            rngSplit.Copy
            'MsgBox "[" & rngSplit.Text & "]"
            
            PasteAfterColon rngSub
            
            ' [[ check footnote subpart content
            
            ' try to determine the signal and set the dropdown
            Dim ccSig As ContentControl
            Set ccSig = CcNext(docRpt, rngSub.Start, "!sig")
            Dim cchSig As Long
            If Not ccSig Is Nothing Then
                cchSig = Len(FindAndSelectSignal(ccSig, rngSplit.Text))
                Call OnExit_Delta(ccSig)
                If cchSig > 0 Then cchSig = cchSig + 1 ' space
            End If
            
            ' flag quotation if it has one
            If FHasQuotation(rngSplit.Text) Then
                Set ccQuote = CcNext(docRpt, rngSub.Start, "!quote")
                If Not ccQuote Is Nothing Then
                    ccQuote.Checked = True
                    Call OnExit_Delta(ccQuote)
                End If
            End If
            
            ' try to determine source type
            Dim sSrc As String
            Dim rngSrc As Range
            Set rngSrc = rngSplit.Duplicate
            rngSrc.Start = rngSrc.Start + cchSig
            sSrc = SourceType(rngSrc)
            If sSrc <> "" Then
                Dim ccSrc As ContentControl
                Set ccSrc = CcNext(docRpt, rngSub.Start, "!source")
                Dim entry As ContentControlListEntry
                For Each entry In ccSrc.DropdownListEntries
                    If entry.Value = sSrc Then
                        entry.Select
                        Call OnExit_Delta(ccSrc)
                        Exit For
                    End If
                Next
            End If
            
            ' ]] check footnote subpart content
            
            Set rngSub = rngNext
            If Not fFound Then
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

    Application.ScreenUpdating = True
End Sub


' Given citation (text), make content control selection based on signal found.
Public Function FindAndSelectSignal(cc As ContentControl, sCite As String) As String
    Dim entry As ContentControlListEntry
    Dim cchLongest As Integer
    For Each entry In cc.DropdownListEntries
        Dim sEntry As String
        sEntry = entry.Text
        If entry.Value = "" Then
            sEntry = ""
        End If
        Dim cch As Long
        cch = Len(sEntry)
        If cch > cchLongest And Mid(sCite, 1, cch) = sEntry Then
            entry.Select
            FindAndSelectSignal = sEntry
            ' no early exit because want to keep longest
        End If
    Next
End Function


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



' Next content control >= cp.
Public Function CcNext(doc As Document, cp As Long, Optional sTag As String = "") As ContentControl
    Dim rng As Range
    Set rng = doc.Range(cp, doc.Range.End)
    ' buggy in tables - collection goes back to start of cell (so O(n) scan should be reasonable)
    Dim cc As ContentControl
    For Each cc In rng.ContentControls
        If sTag = "" Or sTag = cc.Tag Then
            If cc.Range.Start >= cp Then
                Set CcNext = cc
                Exit For
            End If
        End If
    Next
End Function


Public Function OnExit_Delta(ByVal cc As ContentControl) As Boolean
    OnExit_Delta = True ' default success, else the user may get stuck
    
    Dim doc As Document
    Set doc = ActiveDocument
    Dim sPrefix As String
    sPrefix = Mid(cc.Tag, 1 + 1) & "_"
    
    Dim ccText As ContentControl
    Set ccText = CcNext(doc, cc.Range.End)
    
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
                If FBkmkTextMatch(doc, "ALL_" & entry.Value, sText) Then
                    fChanged = False
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
    
    ' allow an "ALL" control name backup for generic options
    Dim fExists As Boolean
    fExists = doc.Bookmarks.Exists(sName)
    If Not fExists Then
        If doc.Bookmarks.Exists("ALL_" & sVal) Then
            sName = "ALL_" & sVal  ' don't prospectively change name for sake of the error message if missing
            fExists = True
        End If
    End If
    
    If Not fExists Then
        If sVal = "" Or sVal = "no" Then
            ccText.Range.Text = ""
        Else
            ' only flag in debug; otherwise may be the bookmarks were removed
            ' after citechecking, which ought to be harmless when passing through it
            ccText.Range.Text = """" & sName & """" & " bookmark not found."
            ccText.Range.Font.ColorIndex = wdRed
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
