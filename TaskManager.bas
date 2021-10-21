Attribute VB_Name = "TaskManager"
Option Base 1
Option Explicit

Public EngraveSR As New ShapeRange
Public CutSR As New ShapeRange
Public UnSortedSR As New ShapeRange

Public Sub CloneSR()
Dim s As Shape
Dim ss As Shape
Dim tempsr As New ShapeRange
Set tempsr = ActiveSelectionRange.Clone

For Each s In tempsr.Shapes
   If s.Type = cdrGroupShape Then
    s.UngroupAllEx
   End If
Next s
For Each s In tempsr.Shapes
   Select Case s.Type
        Case cdrGroupShape: s.UngroupAllEx
        Case cdrTextShape:  s.ConvertToCurves
        Case cdrBlend:      s.BreakApartEx
   End Select
Next s
   
GetUniqueCol tempsr
tempsr.Delete
End Sub
Public Sub DefineColors(sr As ShapeRange)   ' получаем цвета контуров фигур в выделении
Dim s As Shape
Dim r, g, b As Byte
Dim n, k As Integer
Dim ColName As Long
Dim tempsr As New ShapeRange
Dim NoMatches As Boolean ' Есть ли совпадения по цвету
Dim ListOfColors() As Long


ReDim Preserve ListOfColors(1)
For Each s In sr.Shapes
If s.Outline.Type = cdrNoOutline Then GoTo line
         s.Outline.Color.ConvertToRGB
         r = s.Outline.Color.RGBRed: g = s.Outline.Color.RGBGreen: b = s.Outline.Color.RGBBlue
         ColName = RGB(r, g, b)
         If n = 0 Then ListOfColors(1) = ColName
        'Debug.Print ColName
        For k = 1 To UBound(ListOfColors)
            If ColName = ListOfColors(k) Then NoMatches = False: Exit For Else NoMatches = True
        Next k
    If NoMatches = True Then
        ReDim Preserve ListOfColors(UBound(ListOfColors) + 1)
        ListOfColors(UBound(ListOfColors)) = ColName
    End If
n = n + 1
line:
Next s
'_______________
For k = 1 To UBound(ListOfColors)
   
Next k
ReDim TASK(UBound(ListOfColors))
For n = 1 To UBound(ListOfColors)
    TASK(n).OutlineColor = ListOfColors(n)
Next n
End Sub
Public Sub CopySRtoTempLayer(sr As ShapeRange)
Dim templayer As New Layer
Dim s As Shape
Dim i As Long
Dim r, g, b As Byte
Dim n, k As Integer
Dim ColName As Long
Dim tempsr As New ShapeRange
Dim NoMatches As Boolean ' Есть ли совпадения по цвету
Dim ListOfColors() As Long

If sr.count = 0 Then MsgBox ("Выберите объекты!"): Exit Sub
Set templayer = ActivePage.CreateLayer("TempL")
Set UnSortedSR = sr.CopyToLayer(templayer)
Set UnSortedSR = templayer.Shapes.All
For Each s In UnSortedSR
    s.Separate
    s.UngroupAllEx
Next s
Set UnSortedSR = templayer.Shapes.All

TaskSort UnSortedSR




If EngraveSR.count > 0 Then
    ReDim TASK(EngraveSR.count)
    For i = 1 To EngraveSR.count
    With TASK(i)
        .sr.Add EngraveSR.Item(i)
        .IsGrav = True
        .Resolution = 200
        .IsUSE = True
     .PosX = EngraveSR.Item(i).PositionX
        .PosY = EngraveSR.Item(i).PositionY
    End With
    'debug.Print EngraveSR.Item(i).PositionX
    'Debug.Print EngraveSR.Item(i).PositionY
    Next i
End If

    ReDim ListOfColors(1)
    For i = 1 To CutSR.count
    If CutSR(i).Outline.Type = cdrNoOutline Or CutSR.count = 0 Then GoTo line1
         CutSR(i).Outline.Color.ConvertToRGB
         r = CutSR(i).Outline.Color.RGBRed: g = CutSR(i).Outline.Color.RGBGreen: _
         b = CutSR(i).Outline.Color.RGBBlue
         ColName = RGB(r, g, b)
         CutSR(i).Name = ColName
          If i = 1 Then ListOfColors(1) = ColName
        For k = 1 To UBound(ListOfColors)
            If ColName = ListOfColors(k) Then NoMatches = False: Exit For Else NoMatches = True
        Next k
    If NoMatches = True Then
        ReDim Preserve ListOfColors(UBound(ListOfColors) + 1)
        ListOfColors(UBound(ListOfColors)) = ColName
    End If
line1:
    Next i

MainDialog.Label1.Caption = "Unsorted=" & UnSortedSR.count & " Grav=" & _
EngraveSR.count & " Cut=" & CutSR.count & "  Colors=" & UBound(ListOfColors)

If CutSR.count > 0 Then
ReDim Preserve TASK(EngraveSR.count + UBound(ListOfColors))
For i = 1 To UBound(ListOfColors)

With TASK(EngraveSR.count + i)
    .Resolution = 50
    .OutlineColor = ListOfColors(i)
    .IsGrav = False
    .IsUSE = True
    .Repeat = 1
    .sr.AddRange CutSR.Shapes.FindShapes(Name:=CStr(ListOfColors(i)))
    .PosX = .sr.PositionX
    .PosY = .sr.PositionY
End With
'Debug.Print TASK(i + EngraveSR.count).sr.count; TASK(i + EngraveSR.count).PosX * 25.4; TASK(i + EngraveSR.count).PosY * 25.4
Next i
End If
End Sub
Sub DelTemplayer()
Dim lyr As Layer
For Each lyr In ActivePage.Layers
If lyr.Name = "TempL" Then lyr.Delete
Next lyr
End Sub
Sub TaskSort(sr As ShapeRange)
Dim s As Shape
Dim black As New Color
black.RGBAssign 0, 0, 0
For Each s In sr
If s.Type = cdrBitmapShape Or ((s.Fill.UniformColor.GetColorDistanceFrom(black) = 0 And s.Type <> cdrBitmapShape)) Then
EngraveSR.Add s
End If
If s.Fill.UniformColor.GetColorDistanceFrom(black) <> 0 And s.Type <> cdrBitmapShape Then CutSR.Add s
Next s

End Sub
