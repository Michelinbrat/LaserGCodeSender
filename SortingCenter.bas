Attribute VB_Name = "SortingCenter"
Option Explicit

Sub smartBreakApart()

Dim s As Shape, sr As ShapeRange, shs As Shape
Dim sr2 As New ShapeRange, sr3 As ShapeRange
Dim x As Double, y As Double
Dim nodecount As Long, tempDis As Double

   On Error GoTo smartBreakApart_Error

If ActiveSelection.Shapes.count = 0 Then Exit Sub

Optimization = True
EventsEnabled = False
ActiveDocument.BeginCommandGroup "smart break"

ActiveSelection.UngroupAll
Set sr2 = ActiveSelection.BreakApartEx
Set sr = OrderBySize(sr2)

tempDis = 0.005

For Each s In sr
    s.OrderToFront
    s.Fill.ApplyUniformFill CreateRGBColor(64, 32, 32)
    nodecount = 1
    s.Curve.Nodes(nodecount).GetPosition x, y
    
If 1 = 2 Then
1001:
    If nodecount <= s.Curve.Nodes.count Then
        s.Curve.Nodes(nodecount).GetPosition x, y
    Else
        GoTo 1002:
    End If
End If
    
    If s.IsOnShape(x + tempDis, y) = cdrInsideShape And s.IsOnShape(x + tempDis, y) <> cdrOnMarginOfShape Then
        x = x + tempDis
        
    ElseIf s.IsOnShape(x - tempDis, y) = cdrInsideShape And s.IsOnShape(x - tempDis, y) <> cdrOnMarginOfShape Then
        x = x - tempDis
        
    ElseIf s.IsOnShape(x, y + tempDis) = cdrInsideShape And s.IsOnShape(x, y + tempDis) <> cdrOnMarginOfShape Then
        y = y + tempDis
        
    ElseIf s.IsOnShape(x, y - tempDis) = cdrInsideShape And s.IsOnShape(x, y - tempDis) <> cdrOnMarginOfShape Then
        y = y - tempDis
        
    ElseIf s.IsOnShape(x - tempDis, y - tempDis) = cdrInsideShape And s.IsOnShape(x - tempDis, y - tempDis) <> cdrOnMarginOfShape Then
        y = y - tempDis: x = x - tempDis

    ElseIf s.IsOnShape(x + tempDis, y + tempDis) = cdrInsideShape And s.IsOnShape(x + tempDis, y + tempDis) <> cdrOnMarginOfShape Then
        y = y + tempDis: x = x + tempDis
        
    ElseIf s.IsOnShape(x - tempDis, y + tempDis) = cdrInsideShape And s.IsOnShape(x - tempDis, y + tempDis) <> cdrInsideShape Then
        y = y + tempDis: x = x - tempDis
        
    ElseIf s.IsOnShape(x + tempDis, y - tempDis) = cdrInsideShape And s.IsOnShape(x + tempDis, y - tempDis) <> cdrOnMarginOfShape Then
        y = y - tempDis:  x = x + tempDis
        
    Else
        nodecount = nodecount + 1
        's.Fill.ApplyUniformFill CreateRGBColor(255, 0, 0) 'RED - testing
        GoTo 1001:
    End If
1002:
    Set shs = ActivePage.SelectShapesAtPoint(x, y, False, tempDis / 2) 'notice!!! tempdis /2
    If Not IsOdd(shs.Shapes.count) Then s.Fill.ApplyUniformFill CreateRGBColor(255, 255, 121)
    sr2.Add s
Next s

ActiveDocument.EndCommandGroup
Optimization = False
EventsEnabled = True
ActiveWindow.Refresh

   On Error GoTo 0
   Exit Sub

smartBreakApart_Error:
    ActiveDocument.EndCommandGroup
    Optimization = False
    EventsEnabled = True
    ActiveWindow.Refresh
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure smartBreakApart of Module newSmartBreakApart"
    
End Sub

Private Function IsOdd(i As Long) As Boolean
    IsOdd = (i Mod 2) <> 0
End Function

Private Function OrderBySize(sr As ShapeRange) As ShapeRange
    Dim srSorted As New ShapeRange
    Dim s As Shape, i As Integer
    Dim t As Variant, j As Integer, y As Integer
    Dim iUpper As Integer, Condition1 As Boolean
    ReDim ShapesAndSizes(sr.count - 1, 1) As Double 'Create an Array to hold area and staticID
    
    'Add shape data to array
    For i = 1 To sr.count
        ShapesAndSizes(i - 1, 0) = Round(sr(i).SizeWidth * sr(i).SizeHeight, 3) 'Area of the shape
        ShapesAndSizes(i - 1, 1) = sr(i).StaticID 'Static ID of current shape
    Next i
    
    'A very simple sort
    For i = LBound(ShapesAndSizes, 1) To UBound(ShapesAndSizes, 1) - 1
        For j = LBound(ShapesAndSizes, 1) To UBound(ShapesAndSizes, 1) - 1
            Condition1 = ShapesAndSizes(j, 0) <= ShapesAndSizes(j + 1, 0)
            If Condition1 Then
                For y = LBound(ShapesAndSizes, 2) To UBound(ShapesAndSizes, 2)
                    t = ShapesAndSizes(j, y)
                    ShapesAndSizes(j, y) = ShapesAndSizes(j + 1, y)
                    ShapesAndSizes(j + 1, y) = t
                Next y
            End If
        Next
    Next
    
    'Create a ShapeRange from the sorted array
    For i = 0 To sr.count - 1
        srSorted.Add ActivePage.FindShape(StaticID:=ShapesAndSizes(i, 1))
    Next i

    Set OrderBySize = srSorted 'Return the new sorted shaperange
End Function
Public Sub small(sr As ShapeRange)
Dim s As Shape
Dim n As Long
Dim Arr() As Variant
Dim Tempname, tempID As Long
Dim count As Long: count = sr.count
Dim index As Long
Dim flag As Boolean: flag = False
ReDim Arr(count, 2)
For Each s In sr
    s.Name = Round(s.SizeHeight * s.SizeWidth, 2)
    Arr(n, 1) = s.Name: Arr(n, 2) = s.StaticID
    n = n + 1
Next s

again:
For index = 1 To count - 1
If Arr(index + 1, 1) < Arr(index, 1) Then
    flag = True
    Tempname = Arr(index, 1): tempID = Arr(index, 2)
    Arr(index, 1) = Arr(index + 1, 1): Arr(index, 2) = Arr(index + 1, 2)
    Arr(index + 1, 1) = Tempname: Arr(index + 1, 2) = tempID
End If
Next index
If flag = True Then flag = False: GoTo again
For index = 1 To count

                                                'Debug.Print Arr(index, 1); Arr(index, 2)
Next index
End Sub
