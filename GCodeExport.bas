Attribute VB_Name = "GCodeExport"
Option Explicit

Public OutFileName As String
Public M92Ratio, M92StepsPerMM As Double

Public Sub StartExport()
    Open OutFileName For Output As #5
    Print #5, ";***********Marlin firmware oriented GCode export with G5 support*************"
    Print #5, "G21; All units in mm"
    Print #5, "M80; Turn on Optional Peripherals Board at LMN"
    Print #5, "M5;turn the laser off"
        If Mainform.cbUseRotaryAxis = True Then
            Print #5, ";********************* Change Rotary Axis ratio**************************** "
            Print #5, "M92 Y" & M92StepsPerMM * 0.9 * (94 / Mainform.tbRotaryDia)
        End If
    Close #5
End Sub


Public Sub EndExport(ByVal TRFR As Single)
Open OutFileName For Append As #5
    Print #5, "M5; Turn the laser OFF"
    Print #5, "G0 X0 Y0 F" & TRFR & " ; Go Home"
    Print #5, "M400"
    If Mainform.cbUseRotaryAxis = True Then
    Print #5, ";********************* Return Rotary Axis ratio**************************** "
    Print #5, "M92 Y177.18"
    End If
    Print #5, "M300 S300 P300"
    
Close #5
End Sub



Function DigToStr(ByVal Number As Double) As String
Dim number2 As String
number2 = CStr(Round(Number, 3))
DigToStr = Replace(number2, ",", ".", , , vbTextCompare)
End Function
Public Sub ExportCurveShapeIJKL(s2 As Shape, Power As Single, Feed As Double)

    Dim LaserSet As String
    Dim x As Double
    Dim y As Double
    Dim BegX As Double
    Dim BegY As Double
    Dim EndX As Double
    Dim EndY As Double
    Dim seg As Segment
    Dim Sp As SubPath
    Dim i, j, k, l, Xend, Yend As Double
    LaserSet = " S" & Power & " F" & Feed & " P" & PulsesPerMM
    If Not s Is Nothing Then
   
Print #5, ";**********Start shape*************************"

        For Each Sp In s2.Curve.SubPaths
        
            For Each seg In Sp.Segments
    
                'cdrLineSegment

                If seg.Type = cdrLineSegment Then
'Print #5, ";**********************Line segment*************************"
                    BegX = seg.StartNode.PositionX
                    BegY = seg.StartNode.PositionY
                    EndX = seg.EndNode.PositionX
                    EndY = seg.EndNode.PositionY
                If seg.index = 1 Then
                    'Print #5, "M5"
                    Print #5, "G00 X" & DigToStr(BegX) & " Y" & DigToStr(BegY) & _
                    " F" & TraversalFR
                    'Print #5, "G01 F" & Feed
                    Print #5, "M3"
                    'Print #5, "G01 X" & DigToStr(EndX) & " Y" & DigToStr(EndY) '
                                     
                End If
                    Print #5, "G01 X" & DigToStr(EndX) & " Y" & DigToStr(EndY); LaserSet
                    
                End If

                If seg.Type = cdrCurveSegment Then 'G5 Bezier conversation
'Print #5, ";**********************Curve segment*************************"
' G5 parameters are of the form: G5 I0.0 J0.0 K0.0 L0.0 X0.0 Y0.0
'the curve will start from current position
'I/J are the X/Y co-ords of first control point
'K/L are the co-ords of the second control point
'X/Y are the co-ords of the end of the curve
                    i = seg.StartingControlPointX
                    j = seg.StartingControlPointY
                    k = seg.EndingControlPointX
                     l = seg.EndingControlPointY
                    Xend = seg.EndNode.PositionX
                    Yend = seg.EndNode.PositionY
                    Print #5, ";Index of Segment " & seg.index
        If seg.index = 1 Then
                    BegX = seg.StartNode.PositionX
                    BegY = seg.StartNode.PositionY
                    Print #5, "M5"
                    Print #5, "G00 X" & DigToStr(BegX) & " Y" & DigToStr(BegY) & " F" & TraversalFR
                    Print #5, "G1 " & LaserSet
                   ' Print #5, "M3"
                   ' Print #5, "G5 " & "I" & DigToStr(I) & " J" & DigToStr(J) _
                   ' & " K" & DigToStr(K) & " L" & DigToStr(L) _
                    '& " X" & DigToStr(Xend) & " Y" & DigToStr(Yend) '
                                       
        End If
                    Print #5, "M3"
                    'Print #5, "G01 F" & Feed
                    Print #5, "G5 " & "I" & DigToStr(i) & " J" & DigToStr(j) _
                    & " K" & DigToStr(k) & " L" & DigToStr(l) _
                    & " X" & DigToStr(Xend) & " Y" & DigToStr(Yend) & " F" & Feed
                    

                End If

            Next seg
Print #5, "M5 ; Turn the laser off"
        Next Sp

Print #5, ";*******End of shape*************************"
    End If

End Sub

Public Sub ExportCurveShapeIJPQ(shp As Shape, pwr As Byte, FDRT As Single, ByVal TrFeed As Single, ByVal PPM As Single)
Open OutFileName For Append As #5
    Dim LaserSet As String
    Dim x As Double
    Dim y As Double
    Dim BegX As Double
    Dim BegY As Double
    Dim EndX As Double
    Dim EndY As Double
    Dim seg As Segment
    Dim Sp As SubPath
    Dim i, j, p, Q, Xend, Yend As Double
    
    LaserSet = " S" & pwr & " F" & FDRT & " P" & PPM & " L60000 B1 D0"
    If Not shp Is Nothing Then
   
Print #5, ";**********Start shape*************************"

        For Each Sp In shp.DisplayCurve.SubPaths
        
            For Each seg In Sp.Segments
    
            If seg.Type = cdrLineSegment Then
                
'Print #5, ";**********************Line segment*************************"
                    BegX = seg.StartNode.PositionX
                    BegY = seg.StartNode.PositionY
                    EndX = seg.EndNode.PositionX
                    EndY = seg.EndNode.PositionY
                If seg.index = 1 Then
                    
                    Print #5, "G00 X" & DigToStr(BegX) & " Y" & DigToStr(BegY) & _
                    " F" & TrFeed
                    Print #5, "G01 X" & DigToStr(EndX) & " Y" & DigToStr(EndY); LaserSet
              
                    Print #5, "M3"
                End If
                
                    Print #5, "G01 X" & DigToStr(EndX) & " Y" & DigToStr(EndY)
                    
            End If

If seg.Type = cdrCurveSegment Then 'G5 Bezier conversation
'Print #5, ";**********************Curve segment*************************"
' G5 parameters are of the form: G5 I0.0 J0.0 P0.0 Q0.0 X0.0 Y0.0
'the curve will start from current position

                    BegX = seg.StartNode.PositionX
                    BegY = seg.StartNode.PositionY
                    EndX = seg.EndNode.PositionX
                    EndY = seg.EndNode.PositionY
                    
                    i = seg.StartingControlPointX - BegX
                    j = seg.StartingControlPointY - BegY
                    p = seg.EndingControlPointX - EndX
                    Q = seg.EndingControlPointY - EndY
                    'Print #5, ";Index of Segment " & seg.Index
            If seg.index = 1 Then
                    Print #5, "G00 X" & DigToStr(BegX) & " Y" & DigToStr(BegY) & " F" & TrFeed
                    Print #5, "G01 " & LaserSet
                    Print #5, "M3"
                     '
            End If
                    Print #5, "M3"
                    Print #5, "G5 " & "I" & DigToStr(i) & " J" & DigToStr(j) _
                    & " P" & DigToStr(p) & " Q" & DigToStr(Q) _
                    & " X" & DigToStr(EndX) & " Y" & DigToStr(EndY)
End If

            Next seg
Print #5, "M5 ; Turn the laser off"
        Next Sp

Print #5, ";*******End of shape*************************"
    End If
Close #5
End Sub





