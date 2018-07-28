Sub PandocFormatEndnotes()
'
' PandocFormatEndnotes Macro
'
' Converts footnotes that start with <Endnote /> to endnotes. Adjusts formatting of endnote separator and continuation separator.
'
' Last updated: 07-27-2018

    
    Dim MyPaper As Document
    Set MyPaper = ActiveDocument


    ' Convert footnotes that start with <Endnote /> to endnotes
    
    Dim FootnoteObject As Footnote
    
    For Each FootnoteObject In MyPaper.Footnotes
        If InStr(1, FootnoteObject.Range.Text, "<Endnote />") Then
            FootnoteObject.Reference.Footnotes.Convert
        End If
    Next FootnoteObject
    
    
    ' If endnotes created...
    If MyPaper.Endnotes.Count >= 1 Then
        
        Dim PaperEndnotesText As Range
        Set PaperEndnotesText = ActiveDocument.StoryRanges(wdEndnotesStory)
        
        ' Remove <Endnote /> markup amongst endnotes
        With PaperEndnotesText.Find
            .ClearFormatting
            .Text = "<Endnote /> "
            With .Replacement
                .ClearFormatting
                .Text = ""
            End With
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchWildcards = False
            .MatchWholeWord = True
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            .Execute Replace:=wdReplaceAll
        End With
        
        ' Remove endnotes separators
        With MyPaper.Endnotes.Separator
            .Text = ""
            .Style = ActiveDocument.Styles("No Spacing")
        End With
        MyPaper.Endnotes.ContinuationSeparator.Text = ""

    End If
    
    
End Sub
