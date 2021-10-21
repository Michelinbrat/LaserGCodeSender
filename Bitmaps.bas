Attribute VB_Name = "Bitmaps"
Option Explicit
'***********************************************************************************************
'********************** ������� � ������������ ���� ********************************************
'***********************************************************************************************
Public Type ScanLine
StartPXL As Double
EndPXL As Double
Length As Double
Dir As Boolean
isZeroLine As Boolean
End Type
Public PXLArray() As Byte
Public ProcessedLineInfo() As ScanLine ' ���������� � ������ ������, ������ � ����. ������� � �������
'Option Base 1 ' ��������� ������� ���������� � 1

Sub GetPixelArray(ByVal FileName As String) '��������� ������ ������� �� BMP �����
StartImagePointer = 1078
Dim LinePointer, PixelPointer As Double '����� ������, ����� �������
Dim PXLBuffer As Byte ' �������� �������
Open FileName For Binary As #2
ReDim PXLArray(HeightPXL, WidthPXL)
Mainform.ProgressBar.Max = HeightPXL

For LinePointer = 0 To HeightPXL - 1
Mainform.ProgressBar.Max = Mainform.ProgressBar.Max + 1
    For PixelPointer = 0 To WidthPXL - 1
        Get #2, StartImagePointer + PixelPointer + LinePointer * ScanLength, PXLBuffer
            If Mainform.cbInvert = False Then
                PXLBuffer = 255 - PXLBuffer ' ����������� �������� �������
            End If
        If PixelPointer = 0 Then PXLBuffer = 0
        PXLArray(LinePointer, PixelPointer) = PXLBuffer
        
       ' Debug.Print PXLBuffer
      '  Debug.Print PixelPointer
        ' Debug.Print " ____________"
    Next

Next
Close #2
Mainform.ProgressBar.Max = 0
End Sub
Sub ProcessPixelArray() ' ������������ ������ ��������
Dim LinePointer, PixelPointer As Long '����� ������, ����� �������
Dim PXL  As Byte

ReDim ProcessedLineInfo(HeightPXL)
Mainform.ProgressBar.Max = HeightPXL

For LinePointer = 0 To HeightPXL - 1
'If LinePointer = 0 Then GoTo label1
'ProcessedLineInfo(LinePointer).StartPXL = 0: ProcessedLineInfo(LinePointer).EndPXL = WidthPXL - 1:
Mainform.ProgressBar.Max = Mainform.ProgressBar.Max + 1
' ���� ������ �������, ���������� ������
        For PixelPointer = 0 To WidthPXL - 1
            PXL = PXLArray(LinePointer, PixelPointer)
            ProcessedLineInfo(LinePointer).StartPXL = PixelPointer
            If PXL > 0 Or PixelPointer = WidthPXL - 1 Then Exit For
        Next
       
       
      
       ' ��������� ������ � �������� �������
       If PixelPointer = WidthPXL - 1 Then ' ���� � ������ ��� ��������� ������
            With ProcessedLineInfo(LinePointer)
            .isZeroLine = True
            '.Length = 11
            .StartPXL = 0
            .EndPXL = 20
        End With
        GoTo Label1
        End If

' ���� ��������� �������, ���������� ������
        For PixelPointer = WidthPXL - 1 To ProcessedLineInfo(LinePointer).StartPXL Step -1
            PXL = PXLArray(LinePointer, PixelPointer)
            ProcessedLineInfo(LinePointer).EndPXL = PixelPointer
            If PXL > 0 Then Exit For
        Next
        
        ' ���������� ����������� ������������
        
Label1:
If LinePointer Mod 2 = 0 Then ProcessedLineInfo(LinePointer).Dir = True Else ProcessedLineInfo(LinePointer).Dir = False

'Debug.Print LinePointer
'Debug.Print ProcessedLineInfo(LinePointer).StartPXL
'Debug.Print ProcessedLineInfo(LinePointer).EndPXL
'Debug.Print "---"
Next

End Sub
Sub SplitLines() ' �������� ������ � ��������� ��������� ������ � ����� ������
' !!!!!!!!!!!!!!!!!!!! ������ �� �������������� !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim LinePointer As Double
Dim Acc As Long
Dim StartDiff, EndDiff As Double
Acc = 10
For LinePointer = 0 To HeightPXL - 1
    If LinePointer = 0 Then
    ProcessedLineInfo(0).StartPXL = 0
    
    EndDiff = ProcessedLineInfo(1).EndPXL - ProcessedLineInfo(0).EndPXL  ' �������� ����� ����������� �����

    If EndDiff >= 0 Then
        ProcessedLineInfo(1).EndPXL = ProcessedLineInfo(1).EndPXL + Acc
        ProcessedLineInfo(0).EndPXL = ProcessedLineInfo(1).EndPXL ' �������� ������ �� ��������
    Else:
        ProcessedLineInfo(0).EndPXL = ProcessedLineInfo(0).EndPXL + Acc
        ProcessedLineInfo(1).EndPXL = ProcessedLineInfo(0).EndPXL
    End If
    If ProcessedLineInfo(0).EndPXL > WidthPXL - 1 Then ProcessedLineInfo(0).EndPXL = WidthPXL - 1
    If ProcessedLineInfo(1).EndPXL > WidthPXL - 1 Then ProcessedLineInfo(1).EndPXL = WidthPXL - 1
    GoTo LengthCalc
    End If
    ' ************************ ���� ������ ��� ������  ********************************************
    If ProcessedLineInfo(LinePointer).isZeroLine = True Then
    
       ' ProcessedLineInfo(LinePointer).EndPXL = ProcessedLineInfo(LinePointer - 1).EndPXL
       ' ProcessedLineInfo(LinePointer).StartPXL = ProcessedLineInfo(LinePointer).EndPXL - 5
       ' If ProcessedLineInfo(LinePointer).StartPXL < 0 Then ProcessedLineInfo(LinePointer).StartPXL = 0: _
       ' ProcessedLineInfo(LinePointer).EndPXL = 5
       ' If ProcessedLineInfo(LinePointer).EndPXL > WidthPXL - 1 Then ProcessedLineInfo(LinePointer).EndPXL = WidthPXL - 1: _
       ' ProcessedLineInfo(LinePointer).StartPXL = WidthPXL - 5
        'GoTo LengthCalc
    End If
 If ProcessedLineInfo(LinePointer).Dir = True Then
    ' --------------------�������� ����� ������ ��������� ������ � ������ �������------------------------------------
    EndDiff = ProcessedLineInfo(LinePointer + 1).EndPXL - ProcessedLineInfo(LinePointer).EndPXL ' �������� ����� ����������� �����

    If EndDiff >= 0 Then
        ProcessedLineInfo(LinePointer + 1).EndPXL = ProcessedLineInfo(LinePointer + 1).EndPXL + Acc
        ProcessedLineInfo(LinePointer).EndPXL = ProcessedLineInfo(LinePointer + 1).EndPXL ' �������� ������ �� ��������
    Else:
        ProcessedLineInfo(LinePointer).EndPXL = ProcessedLineInfo(LinePointer).EndPXL + Acc
        ProcessedLineInfo(LinePointer + 1).EndPXL = ProcessedLineInfo(LinePointer).EndPXL
    End If
    
       
    Else '******************************************���� ������ �������� ***********************
    StartDiff = ProcessedLineInfo(LinePointer + 1).StartPXL - ProcessedLineInfo(LinePointer).StartPXL

    If StartDiff >= 0 Then
        ProcessedLineInfo(LinePointer).StartPXL = ProcessedLineInfo(LinePointer).StartPXL - Acc ' �������� ������ �� ��������
        ProcessedLineInfo(LinePointer + 1).StartPXL = ProcessedLineInfo(LinePointer).StartPXL
    Else:
        ProcessedLineInfo(LinePointer + 1).StartPXL = ProcessedLineInfo(LinePointer + 1).StartPXL - Acc
        ProcessedLineInfo(LinePointer).StartPXL = ProcessedLineInfo(LinePointer + 1).StartPXL
    End If
End If

If ProcessedLineInfo(LinePointer).EndPXL > WidthPXL - 1 Then ProcessedLineInfo(LinePointer).EndPXL = WidthPXL - 1
If ProcessedLineInfo(LinePointer + 1).EndPXL > WidthPXL - 1 Then ProcessedLineInfo(LinePointer + 1).EndPXL = WidthPXL - 1
If ProcessedLineInfo(LinePointer).StartPXL < 0 Then ProcessedLineInfo(LinePointer).StartPXL = 0
If ProcessedLineInfo(LinePointer + 1).StartPXL < 0 Then ProcessedLineInfo(LinePointer + 1).StartPXL = 0

LengthCalc:

 

'If LinePointer = 0 Then
ProcessedLineInfo(0).StartPXL = 0
ProcessedLineInfo(LinePointer).Length = ProcessedLineInfo(LinePointer).EndPXL - ProcessedLineInfo(LinePointer).StartPXL + 1 '***���������
'Debug.Print LinePointer; "S "; ProcessedLineInfo(LinePointer).StartPXL; "E "; ProcessedLineInfo(LinePointer).EndPXL
'Debug.Print ProcessedLineInfo(LinePointer).EndPXL
'Debug.Print "---"; ProcessedLineInfo(LinePointer).Length
Next

End Sub
Sub ExportBMPShape2(ByVal PosX, ByVal PosY As Double, ByVal pwr As Byte, ByVal FR As Long)
Dim TestGcodeFile As String
Dim LinePointer, CNT, CNt51, ChunkCNT  As Double
Dim WholeString() As Byte
Dim String51() As Byte
Dim flgDir As Integer
Dim prnDir As String
Dim Chunks, LastChunk, LastChunkToPrn, TotalChunks, ChunkLength As Long
Open OutFileName For Append As #5
Print #5, ";**************************��������� ������*****************************"
Print #5, ";****             ������� ����������� "; WidthPXL & "X"; HeightPXL
Print #5, ";*****************************************************************************"
Print #5, "M649 S" & pwr & " B2 D0 R" & DigToStr(BeamDia)
Print #5, "G0 X" & DigToStr(PosX) & " Y" & DigToStr(PosY) & " F" & FR
Mainform.ProgressBar.Max = HeightPXL
For LinePointer = 0 To HeightPXL - 1
'Debug.Print LinePointer; "S "; ProcessedLineInfo(LinePointer).StartPXL; "E "; ProcessedLineInfo(LinePointer).EndPXL
'Debug.Print ProcessedLineInfo(LinePointer).Length
    '************* ���������� ��������� ������ - ����� ������� � �� *************************

    Chunks = ProcessedLineInfo(LinePointer).Length \ 51 ' ���������� ����� ������� �� 51 �����
    LastChunk = (ProcessedLineInfo(LinePointer).Length - Chunks * 51)
    If LastChunk = 0 And Chunks = 0 Then LastChunk = 2 ' �����  ��������� �������
    
    Select Case LastChunk Mod 3
    Case 0: LastChunkToPrn = (LastChunk * 4) \ 3
    Case 1: LastChunkToPrn = (LastChunk * 4) \ 3 + 3
    Case 2: LastChunkToPrn = (LastChunk * 4) \ 3 + 2
    End Select
       
    ChunkCNT = 0 ' �������� ������� �������
    CNt51 = 0    ' � ������� ������� 51
 '********************************************************************************************
    If ProcessedLineInfo(LinePointer).Dir = True Then  ' ������������ ������ ������
    prnDir = "$1"
    If Chunks = 0 Then GoTo Label1
        For ChunkCNT = 0 To Chunks - 1 ' ��������� ������ ��������� �� 51 ����
        ReDim String51(50)
        If ChunkCNT > 0 Then prnDir = ""
            For CNT = 0 To 50
                String51(CNT) = PXLArray(LinePointer, ProcessedLineInfo(LinePointer).StartPXL + CNT + ChunkCNT * 51)
            Next
            
            Print #5, "G7 "; prnDir; " L68 "; " D" & ReplaceStr(String51)
        ' ���� ��������� �������
        Next
Label1: If LastChunk > 0 Then
             If Chunks > 0 Then prnDir = "" Else prnDir = "$1"
            ReDim String51(LastChunk - 1)
            For CNT = 0 To LastChunk - 1
                String51(CNT) = PXLArray(LinePointer, ProcessedLineInfo(LinePointer).StartPXL + CNT + ChunkCNT * 51)
            Next
           
            Print #5, "G7"; prnDir; "  L" & LastChunkToPrn; " D" & ReplaceStr(String51)
        End If
       
       
    Else ' ������������ �������� ������
    prnDir = "$0"
    If Chunks = 0 Then GoTo Label2
    For ChunkCNT = 0 To Chunks - 1 ' ��������� ������ ��������� �� 51 ����
    
            ReDim String51(50)
            If ChunkCNT <> 0 Then prnDir = ""
            For CNT = 0 To 50
                String51(CNT) = PXLArray(LinePointer, ProcessedLineInfo(LinePointer).EndPXL - CNT - ChunkCNT * 51)
            Next
            
           Print #5, "G7 "; prnDir; " L68 "; " D" & ReplaceStr(String51)
    Next
Label2:        If LastChunk > 0 Then
               If Chunks <> 0 Then prnDir = "" Else prnDir = "$0"
                ReDim String51(LastChunk - 1)
                        For CNT = 0 To LastChunk - 1
                        String51(CNT) = PXLArray(LinePointer, ProcessedLineInfo(LinePointer).EndPXL - CNT - ChunkCNT * 51)
                        Next
            If ChunkCNT > 0 Then prnDir = ""
            
            Print #5, "G7"; prnDir; "  L" & LastChunkToPrn; " D" & ReplaceStr(String51)
        End If
   
    End If
 Mainform.ProgressBar = Mainform.ProgressBar + 1
Next
   
Close #5
Mainform.ProgressBar = 0
End Sub

Function ReplaceStr(InString() As Byte) As String
ReplaceStr = Base64Encode(InString)
ReplaceStr = Replace(ReplaceStr, "/", "9")
ReplaceStr = Replace(ReplaceStr, "+", "9")
End Function


