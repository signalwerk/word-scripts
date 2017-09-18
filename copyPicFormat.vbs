Sub pasteXY()
Dim pLeft As Double
Dim pTop As Double

sPrompt = "Please insert XY Properties"
sTitle = "xy"
sDefault = ""
param = Trim(InputBox(sPrompt, sTitle, sDefault))
If param <> "" Then
param = Split(param, "|")
pLeft = CDbl(param(0))
pTop = CDbl(param(1))
  
Selection.ShapeRange.left = pLeft
Selection.ShapeRange.top = pTop
  End If



End Sub


Sub copyXY()


Dim pLeft As Double
Dim pTop As Double
pLeft = Selection.ShapeRange.left
pTop = Selection.ShapeRange.top

sPrompt = "Here the XY to Copy"
sTitle = "xy"
sDefault = CStr(pLeft) + "|" + CStr(pTop)
param = Trim(InputBox(sPrompt, sTitle, sDefault))


End Sub






Sub insertPicByParam()



' MsgBox (JSONlib.toString(Array("a", "b", Array(1, b, "3"))))
sPrompt = "Please insert pic Properties"
sTitle = "Pic Data"
sDefault = ""
param = Trim(InputBox(sPrompt, sTitle, sDefault))
If param <> "" Then
param = Split(param, "|")



Dim pLeft As Double
Dim pTop As Double
Dim pWidth As Double
Dim pHeight As Double
Dim pPath As String

' MsgBox (param(0))
pLeft = CDbl(param(0))
pTop = CDbl(param(1))
pWidth = CDbl(param(2))
pHeight = CDbl(param(3))
pPath = param(4)


Call importByParams(pPath, pLeft, pTop, pWidth)
 
  GoTo Done
  End If
' MsgBox (pLeft)



Application.ScreenUpdating = False
Dim Rng As Range, shp As Shape





With Application.Dialogs(wdDialogInsertPicture)
  .Display
  If .Name <> "" Then
    importByFile (.Name)
  GoTo Done
  End If
End With
Done:
Set Rng = Nothing: Set shp = Nothing
Application.ScreenUpdating = True
End Sub




Sub importByFile(strPath As String)

    
    Call importByParamsA(strPath, 20, 10, 120)

End Sub


Sub importByParams(strPath As String, pLeft As Double, pTop As Double, pWidth As Double)


    
    Set Rng = Selection.Range
    Rng.Collapse
    Set pic = ActiveDocument.InlineShapes.AddPicture(filename:=strPath, SaveWithDocument:=True, Range:=Rng)

    Set shp = pic.ConvertToShape
    
    With shp
       
      ' https://msdn.microsoft.com/en-us/library/bb214041(office.12).aspx
      ' 5 = wdWrapBehind
      .WrapFormat.Type = 5
      .LockAspectRatio = True
      .RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
      .RelativeVerticalPosition = wdRelativeVerticalPositionPage
    
      .left = CentimetersToPoints(pLeft / 10)
      .top = CentimetersToPoints(pTop / 10)
      .width = CentimetersToPoints(pWidth / 10)
    End With


End Sub



