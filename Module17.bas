Attribute VB_Name = "Module17"
Sub Углыитангенсы_Кнопка8_Щелчок()
Dim ACad As Object
Dim ADoc As Object
Dim MSpace As Object
Dim TStyle As Object
Dim begp(0 To 2) As Double
Dim endp(0 To 2) As Double
Dim Heightt As Double
Dim buf As Object
Dim s, s1 As String
Dim i As Long
Dim poly As AcadLWPolyline
Dim i1 As Integer
Dim n As Double
Dim line As AcadLine
Dim DownPoint, UpPoint As Long
Dim arc As AcadArc
Dim startAngleInDegree As Double
Dim endAngleInDegree As Double



Set ACad = GetObject(, "AutoCAD.Application")
Set ADoc = ACad.ActiveDocument
Set MSpace = ADoc.ModelSpace

Dim txt As String
Dim ssetObj As AcadSelectionSet
Dim mode As Integer
Dim gpCode(0) As Integer
Dim dataValue(0) As Variant
Dim coord As Variant



Randomize

i1 = Int((10000 * Rnd) + 1)

s = "SSET" & CStr(i1)

If ssetObj Is Nothing Then
   Set ssetObj = ADoc.SelectionSets.Add(s)
  Else
    ssetObj.Clear
End If


ssetObj.SelectOnScreen


Sheets("Лист разработчика").Activate
Cells.Select
Selection.ClearContents

Sheets("1. Углы и тангенсы").Activate

On Error GoTo ErrorHandler
 
For Each poly In ssetObj
    For i = 0 To 2147483645
        coord = poly.Coordinate(i)
        Sheets("Лист разработчика").Cells(1 + i, 1).Value = coord(0)
        Sheets("Лист разработчика").Cells(1 + i, 2).Value = coord(1)

    Next

Next

ErrorHandler:

DownPoint = Application.Min(Sheets("Лист разработчика").Range("B:B"))
UpPoint = Application.Max(Sheets("Лист разработчика").Range("B:B"))

For i = 1 To Sheets("Лист разработчика").UsedRange.Rows.Count - 1
    
'рисуем табличку уклонов
    begp(0) = Sheets("Лист разработчика").Cells(i, 1).Value
    begp(1) = DownPoint - 100
    endp(0) = Sheets("Лист разработчика").Cells(i, 1).Value
    endp(1) = DownPoint - 120
    Set line = MSpace.AddLine(begp, endp)
    
    begp(0) = Sheets("Лист разработчика").Cells(i, 1).Value
    begp(1) = DownPoint - 100
    endp(0) = Sheets("Лист разработчика").Cells(i + 1, 1).Value
    endp(1) = DownPoint - 100
    Set line = MSpace.AddLine(begp, endp)
    
    begp(0) = Sheets("Лист разработчика").Cells(i, 1).Value
    begp(1) = DownPoint - 120
    endp(0) = Sheets("Лист разработчика").Cells(i + 1, 1).Value
    endp(1) = DownPoint - 120
    Set line = MSpace.AddLine(begp, endp)
    
    If Sheets("1. Углы и тангенсы").Cells(i + 1, 26).Value = 0 Then
        begp(0) = Sheets("Лист разработчика").Cells(i, 1).Value
        begp(1) = DownPoint - 110
        endp(0) = Sheets("Лист разработчика").Cells(i + 1, 1).Value
        endp(1) = DownPoint - 110
        Set line = MSpace.AddLine(begp, endp)
        
        begp(0) = (Sheets("Лист разработчика").Cells(i + 1, 1).Value + Sheets("Лист разработчика").Cells(i, 1).Value) / 2
        begp(1) = DownPoint - 105
        n = WorksheetFunction.RoundDown(Sheets("1. Углы и тангенсы").Cells(i + 1, 26).Value, 3)
        s1 = n
        Set buf = MSpace.AddMText(begp, 2, s1)
        buf.AttachmentPoint = acAttachmentPointMiddleCenter
        buf.Height = 2.5
        
        begp(0) = (Sheets("Лист разработчика").Cells(i + 1, 1).Value + Sheets("Лист разработчика").Cells(i, 1).Value) / 2
        begp(1) = DownPoint - 115
        n = WorksheetFunction.RoundDown(Sheets("1. Углы и тангенсы").Cells(i + 1, 25).Value, 2)
        s1 = n
        Set buf = MSpace.AddMText(begp, 2, s1)
        buf.AttachmentPoint = acAttachmentPointMiddleCenter
        buf.Height = 2.5
    
    End If
    
    If Sheets("1. Углы и тангенсы").Cells(i + 1, 26).Value > 0 Then
        begp(0) = Sheets("Лист разработчика").Cells(i, 1).Value
        begp(1) = DownPoint - 100
        endp(0) = Sheets("Лист разработчика").Cells(i + 1, 1).Value
        endp(1) = DownPoint - 120
        Set line = MSpace.AddLine(begp, endp)
        
        'уклон
        begp(0) = Sheets("Лист разработчика").Cells(i + 1, 1).Value - 5
        begp(1) = DownPoint - 105
        n = Abs(WorksheetFunction.RoundDown(Sheets("1. Углы и тангенсы").Cells(i + 1, 26).Value, 3))
        s1 = n
        Set buf = MSpace.AddMText(begp, 2, s1)
        buf.AttachmentPoint = acAttachmentPointMiddleRight
        buf.Height = 2.5
        
        'длина участка
        begp(0) = Sheets("Лист разработчика").Cells(i, 1).Value + 2
        begp(1) = DownPoint - 115
        n = Abs(WorksheetFunction.RoundDown(Sheets("1. Углы и тангенсы").Cells(i + 1, 25).Value, 2))
        s1 = n
        Set buf = MSpace.AddMText(begp, 2, s1)
        buf.AttachmentPoint = acAttachmentPointMiddleLeft
        buf.Height = 2.5
    End If

    If Sheets("1. Углы и тангенсы").Cells(i + 1, 26).Value < 0 Then
        begp(0) = Sheets("Лист разработчика").Cells(i, 1).Value
        begp(1) = DownPoint - 120
        endp(0) = Sheets("Лист разработчика").Cells(i + 1, 1).Value
        endp(1) = DownPoint - 100
        Set line = MSpace.AddLine(begp, endp)
        
        'уклон
        begp(0) = Sheets("Лист разработчика").Cells(i, 1).Value
        begp(1) = DownPoint - 105
        n = Abs(WorksheetFunction.RoundDown(Sheets("1. Углы и тангенсы").Cells(i + 1, 26).Value, 3))
        s1 = n
        Set buf = MSpace.AddMText(begp, 2, s1)
        buf.AttachmentPoint = acAttachmentPointMiddleLeft
        buf.Height = 2.5
        
        'длина участка
        begp(0) = Sheets("Лист разработчика").Cells(i + 1, 1).Value - 5
        begp(1) = DownPoint - 115
        n = Abs(WorksheetFunction.RoundDown(Sheets("1. Углы и тангенсы").Cells(i + 1, 25).Value, 2))
        s1 = n
        Set buf = MSpace.AddMText(begp, 2, s1)
        buf.AttachmentPoint = acAttachmentPointMiddleRight
        buf.Height = 2.5
    End If
Next

    begp(0) = Sheets("Лист разработчика").Cells(Sheets("Лист разработчика").UsedRange.Rows.Count, 1).Value
    begp(1) = DownPoint - 100
    endp(0) = Sheets("Лист разработчика").Cells(Sheets("Лист разработчика").UsedRange.Rows.Count, 1).Value
    endp(1) = DownPoint - 120
    Set line = MSpace.AddLine(begp, endp)



'добавляем надписи над углами с полукругом
For i = 1 To Sheets("Лист разработчика").UsedRange.Rows.Count - 2
'линия до надписи
    begp(0) = Sheets("Лист разработчика").Cells(i + 1, 1).Value
    begp(1) = Sheets("Лист разработчика").Cells(i + 1, 2).Value
    endp(0) = Sheets("Лист разработчика").Cells(i + 1, 1).Value
    endp(1) = UpPoint + 50
    Set line = MSpace.AddLine(begp, endp)
'надпись
    begp(0) = Sheets("Лист разработчика").Cells(i + 1, 1).Value + 2
    begp(1) = UpPoint + 50
    s1 = Sheets("1. Углы и тангенсы").Cells(i + 6, 12).Value
    Set buf = MSpace.AddMText(begp, 15, s1)
    buf.AttachmentPoint = acAttachmentPointMiddleLeft
    buf.Height = 2.5
    
    begp(0) = Sheets("Лист разработчика").Cells(i + 1, 1).Value - 2
    begp(1) = UpPoint + 15
    s1 = Sheets("1. Углы и тангенсы").Cells(i + 6, 13).Value
    Set buf = MSpace.AddMText(begp, 2, s1)
    buf.AttachmentPoint = acAttachmentPointMiddleLeft
    buf.Height = 2.5
    buf.Rotation = 1.5708
    
'добавляем полукруг
    If Sheets("1. Углы и тангенсы").Cells(i + 1, 28).Value = "+" Then
        begp(0) = Sheets("Лист разработчика").Cells(i + 1, 1).Value
        begp(1) = UpPoint + 50
        startAngleInDegree = 0
        endAngleInDegree = 3.142
        Set arc = MSpace.AddArc(begp, 1, startAngleInDegree, endAngleInDegree)
    Else
        begp(0) = Sheets("Лист разработчика").Cells(i + 1, 1).Value
        begp(1) = UpPoint + 50
        startAngleInDegree = 3.142
        endAngleInDegree = 0
        Set arc = MSpace.AddArc(begp, 1, startAngleInDegree, endAngleInDegree)
    End If
    
Next
    
    
    

Range(Cells(8, 2), Cells(8, 30)).Select
Selection.AutoFill Destination:=Range(Cells(8, 2), Cells(Sheets("Лист разработчика").UsedRange.Rows.Count + 8, 30))
Range(Cells(Sheets("Лист разработчика").UsedRange.Rows.Count + 5, 1), Cells(ActiveSheet.UsedRange.Rows.Count + 5, 30)).Select
Selection.Clear


ADoc.Regen (True)
'MsgBox ("Все готово!")
End Sub
