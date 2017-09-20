Public Sub FindReplaceAnywhere(pFindTxt As String, pReplaceTxt As String)
    Dim rngStory As Word.Range
    Dim lngJunk As Long
    Dim oShp As Shape

    'Fix the skipped blank Header/Footer problem
    lngJunk = ActiveDocument.Sections(1).Headers(1).Range.StoryType

    'Fix I don't know --- sh
    For Each oShp In ActiveDocument.Sections(1).Headers(wdHeaderFooterPrimary).Shapes
        If oShp.Type = msoTextBox Then
            SearchAndReplaceInStory oShp.TextFrame.TextRange, pFindTxt, pReplaceTxt
        End If
    Next oShp

    'Iterate through all story types in the current document
    For Each rngStory In ActiveDocument.StoryRanges

        'Iterate through all linked stories
        Do
            SearchAndReplaceInStory rngStory, pFindTxt, pReplaceTxt
            On Error Resume Next
            Select Case rngStory.StoryType
                Case WdStoryType.wdEvenPagesHeaderStory, _
                     WdStoryType.wdPrimaryHeaderStory, _
                     WdStoryType.wdEvenPagesFooterStory, _
                     WdStoryType.wdPrimaryFooterStory, _
                     WdStoryType.wdFirstPageHeaderStory, _
                     WdStoryType.wdFirstPageFooterStory



                    If rngStory.ShapeRange.Count > 0 Then
                        For Each oShp In rngStory.ShapeRange

                            If oShp.TextFrame.HasText Then
                                ' oShp.TextFrame.TextRange.LanguageID = wdSwissGerman
                                SearchAndReplaceInStory oShp.TextFrame.TextRange, pFindTxt, pReplaceTxt
                            End If
                        Next
                    End If
                Case Else
                    'Do Nothing
                End Select
                On Error GoTo 0

                'Get next linked story (if any)
                Set rngStory = rngStory.NextStoryRange
            Loop Until rngStory Is Nothing
        Next
End Sub

Public Sub SearchAndReplaceInStory(ByVal rngStory As Word.Range, ByVal strSearch As String, ByVal strReplace As String)
    With rngStory.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = strSearch
        .Replacement.Text = strReplace
        .Wrap = wdFindContinue
        .Execute replace:=wdReplaceAll
    End With
End Sub


' vbCr = Char(13) = CR (Carriage-return character). Used by Mac OS and Apply II family
' vbLf = Char(10) = LF (line-feed character). Used by Linux and Mac OS X
' vbCrLf = Char(13) + Char(10). CR LF (carriage-return followed by line-feed) Used by Windows
' vbNewLine = the same as vbCrLf

Public Sub SaveReplaced()

    Dim csvString As String
    Dim strLines() As String
    Dim strHeader() As String
    Dim strValue() As String
    Dim i, n As Long
    Dim line As String
    Dim field As String
    Dim sSaveAsPath As String


    csvString = InputBox("Enter TSV to process:", "Datainput", "Vorname" & vbTab & "Name" & vbLf & "Cordula" & vbTab & "Simmer")
    csvString = RepText(csvString, vbCr, vbLf)
    csvString = RepText(csvString, vbCrLf, vbLf)

    strLines = Split(csvString, vbLf)

    strHeader = Split(strLines(0), vbTab)


    For i = 1 To UBound(strLines)
        line = strLines(i)
        
        strValue = Split(line, vbTab)

        ' set new save path - filename is first field
        sSaveAsPath = ActiveDocument.Path & "/" & strValue(0) & ".dotm"

        
        'Save changes to original document
        ActiveDocument.Save

        ' add new document
        ' Dim oNewDoc As Document
        ' Set oNewDoc = Documents.Add(ActiveDocument.FullName)

        'copies the active document
        Application.Documents.Add ActiveDocument.FullName

        'saves the copy
        ' ActiveDocument.SaveAs sSaveAsPath
        ' https://msdn.microsoft.com/en-us/library/microsoft.office.interop.word.wdsaveformat(v=office.14).aspx
        ActiveDocument.SaveAs fileName:=sSaveAsPath, FileFormat:=wdFormatXMLTemplateMacroEnabled


        ' now replace all the {{fieldname}} texts
        For n = LBound(strHeader) To UBound(strHeader)
            FindReplaceAnywhere "{{" & strHeader(n) & "}}", strValue(n)
            Debug.Print strHeader(n) & ": " & strValue(n)
        Next n

        'Save changes to new document
        ActiveDocument.Save

        'closes the copy leaving then the original document
        ActiveDocument.Close

    Next i

End Sub


Function RepText(sIn As String, sFind As String, sRep As String) As String
    Dim x As Integer

    x = InStr(sIn, sFind)
    While x > 0
        sIn = Left(sIn, x - 1) & sRep & Mid(sIn, x + Len(sFind))
        x = InStr(sIn, sFind)
    Wend
    RepText = sIn
End Function
