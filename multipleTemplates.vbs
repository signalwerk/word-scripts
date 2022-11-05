Public Sub FindReplaceAnywhere(pFindTxt As String, pReplaceTxt As String)
    Dim rngStory    As Word.Range
    Dim lngJunk     As Long
    Dim oShp        As shape

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
        .Execute Replace:=wdReplaceAll
    End With
End Sub

' vbCr = Char(13) = CR (Carriage-return character). Used by Mac OS and Apply II family
' vbLf = Char(10) = LF (line-feed character). Used by Linux and Mac OS X
' vbCrLf = Char(13) + Char(10). CR LF (carriage-return followed by line-feed) Used by Windows
' vbNewLine = the same as vbCrLf

Public Sub SaveReplaced()

    Dim csvString   As String
    Dim strLines()  As String
    Dim strHeader() As String
    Dim strValue()  As String
    Dim i, n        As Long
    Dim line        As String
    Dim field       As String
    Dim sSaveAsPath As String
    Dim clipboard   As MSForms.DataObject

    Set clipboard = New MSForms.DataObject
    clipboard.GetFromClipboard
    csvString = clipboard.GetText

    'Get User Input
    vUserInput = MsgBox("Do you want To run With the clipboard content?" & vbNewLine & vbNewLine & "Preview:" & vbNewLine & Left(csvString, 250), vbYesNo)

    'Process User Input Yes,No,Cancel
    Select Case vUserInput
        Case vbYes
            'Code if User Input is Yes
        Case vbNo
            'Code if User Input is No
            Exit Sub
    End Select

    csvString = RepText(csvString, vbCrLf, vbLf)
    csvString = RepText(csvString, vbCr, vbLf)
    csvString = RepText(csvString, vbLf & vbLf, vbLf)
    ' Debug.Print "csvString"
    ' Debug.Print csvString

    strLines = Split(csvString, vbLf)

    strHeader = Split(strLines(0), vbTab)

    For i = 1 To UBound(strLines)
        line = strLines(i)
        
        strValue = Split(line, vbTab)
        ' Debug.Print "line"
        ' Debug.Print "'" & line & "'"

        ' set new save path - filename is first field
        sSaveAsPath = ActiveDocument.Path & "/" & strValue(0) & ".dotm"

        'Save changes to original document
        ActiveDocument.Save

        'copies the active document
        Application.Documents.Add ActiveDocument.FullName

        'saves the copy
        ' ActiveDocument.SaveAs sSaveAsPath
        ' https://msdn.microsoft.com/en-us/library/microsoft.office.interop.word.wdsaveformat(v=office.14).aspx
        ActiveDocument.SaveAs FileName:=sSaveAsPath, FileFormat:=wdFormatXMLTemplateMacroEnabled

        ' now replace all the {{fieldname}} texts
        For n = LBound(strHeader) To UBound(strHeader)

            ' Test if string begins with DEL_PIC and if so delete based on alt-text
            If InStr(1, strHeader(n), "DEL_PIC") = 1 Then
                Dim altTexts() As String
                Dim altText As String
                altTexts = Split(strValue(n), "|")
                For iAltText = 0 To UBound(altTexts)
                    altText = altTexts(iAltText)
                    Debug.Print "Delete by ID: " & "'" & altText & "'"
                    deleteAllPicturesByAltText (altText)
                Next iAltText
            Else
                FindReplaceAnywhere "{{" & strHeader(n) & "}}", strValue(n)
                Debug.Print "search For {{" & strHeader(n) & "}} replace by '" & strValue(n) & "'"
            End If

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

' Sub DeletePicByAltText()
'     Dim myPic As shape
'     Set myPic = getPictureByAltText("novopress")
'     If myPic Is Nothing Then
'         MsgBox "Your Picture was Not found. Check the 'Alt Text' is correct and try again."
'     End If
'     On Error Resume Next
'     myPic.Delete
' End Sub

Function getPictureByAltText(altText As String) As shape
    Dim shape As Variant

    ' Debug.Print "count content"
    ' Debug.Print ActiveDocument.Shapes.Count

    ' Debug.Print "count header/footer"
    ' Debug.Print ActiveDocument.Sections(1).Headers(wdHeaderFooterPrimary).Shapes.Count

    ' loop content
    For Each shape In ActiveDocument.Shapes
        If shape.AlternativeText = altText Then
            Set getPictureByAltText = shape
            Exit Function
        End If
    Next

    ' loop header and footer
    For Each shape In ActiveDocument.Sections(1).Headers(wdHeaderFooterPrimary).Shapes
        If shape.AlternativeText = altText Then
            Set getPictureByAltText = shape
            Exit Function
        End If
    Next

End Function

Sub deleteAllPicturesByAltText(altText As String)
    
    On Error GoTo skip
    Do
        Dim myPic   As shape
        Set myPic = getPictureByAltText(altText)
        myPic.Delete
    Loop
skip:
    
End Sub