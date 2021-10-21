Attribute VB_Name = "BMPConvert"
Option Explicit
Type BITMAPFILEHEADER
bfType As Integer
bfSize As Long
bfReserved1 As Integer
bfReserved2 As Integer
bfOffBits As Long
End Type


Type BITMAPINFOHEADER
biSize As Long
biWidth As Long
biHeight As Long
biPlanes As Integer '?????????? ???????? ?????????? - ?????? 1
biBitCount As Integer '?????? ?????????? ??? ?? ??????:

biCompression As Long '??? ?????? (???? BI_RGB, ?? ?????? ?? ????????????)
biSizeImage As Long '?????? ??????????? ? ??????
biXPelsPerMeter As Long '?????????? ???????? ?? ???? ??-???????????, ??? DIB
biYPelsPerMeter As Long '?????????? ???????? ?? ???? ??-?????????, ??? DIB
biClrUsed As Long '?????????? ?????????? ???????????? ????????? ? ???????? ??????? DIP
biClrImportant As Long '?????????? ???????? ????????? ? ??????? DIP
End Type

Dim BMPTempFilename As String
Dim FileName As String   'имя открываемого файла
Dim TypeOfFile As String 'тип файл (BM)
Public StartImagePointer As Double
Dim b As Byte
Dim RazmW As String 'ширина изображения
Dim RazmH As String 'высота изображения
Dim RazmF As String 'размер файла в байтах
Dim x As Long
Public WidthPXL, HeightPXL As Double
Dim PixelValue As Byte
Dim LastChain, LastChainToPrn As Long
Dim WholeChains As Long '  Число полных цепочек по 51 пикселю
Public ScanLength As Double ' Длина строки сканирования
Public GCodeFileNum, BMPFileNum As Byte
' Параметры изображения
Public EngraveResol As Long
Public BeamDia As Double
Public Sub BMPToDisk(ByVal s As ShapeRange)
Dim expflt As ExportFilter
Dim Xpxl, Ypxl As Long
    s.CreateSelection
    Dim pal As StructPaletteOptions
    Set pal = New StructPaletteOptions
    With pal
        .PaletteType = cdrPaletteGrayscale
        .DitherType = cdrDitherOrdered
        .DitherIntensity = 100
    End With
    Xpxl = Round(s.SizeWidth * EngraveResol / 25.4, 0)
    Ypxl = Round(s.SizeHeight * EngraveResol / 25.4, 0)
    Set expflt = ActiveDocument.ExportBitmap("C:\temp\tmpbitmap.bmp", cdrBMP, cdrSelection, cdrPalettedImage, _
    Xpxl, Ypxl, EngraveResol, EngraveResol, cdrNormalAntiAliasing, False, False, False, False, cdrCompressionNone, pal)
    expflt.Finish
    'MsgBox "РИСУНОК успешно экспортирован"
    Close
End Sub

Public Sub GetInfoFromBmp(ByVal fname As String)
Dim bmpInfo As BITMAPINFOHEADER
Dim bmpType As BITMAPFILEHEADER
'FileName = Fname
Open fname For Binary As #2
Get #2, , bmpType
Get #2, , bmpInfo
Close #2
StartImagePointer = bmpType.bfOffBits
WidthPXL = bmpInfo.biWidth
HeightPXL = bmpInfo.biHeight
'MsgBox WidthPXL & "-x-" & HeightPXL
ScanLength = WidthPXL

If WidthPXL Mod 4 > 0 Then ScanLength = (WidthPXL \ 4 + 1) * 4
  
 WholeChains = ScanLength \ 51
 LastChain = ScanLength - WholeChains * 51
 'MsgBox "LastChain=" & LastChain & " Wholechains = " & WholeChains & "ScanLength =" & ScanLength
 Select Case LastChain Mod 3
    Case 0: LastChainToPrn = (LastChain * 4) \ 3
    Case 1: LastChainToPrn = (LastChain * 4) \ 3 + 3
    Case 2: LastChainToPrn = (LastChain * 4) \ 3 + 2
    End Select
End Sub

Function GetString51(ByVal StartPos As Double, ByVal StrLen As Byte, ByVal Direction As Boolean, ByVal LowerLevel As Byte) As String
Dim CNT As Long
Dim Line51() As Byte
ReDim Line51(StrLen)
Dim Lambda As Double
Lambda = (255 - LowerLevel) / 255 'Замена на / 6 дек 2017
CNT = 0
If Direction = True Then
    For CNT = 0 To StrLen
        Get #2, StartPos + CNT, Line51(CNT)
        Line51(CNT) = LowerLevel + Round(Lambda * Line51(CNT), 0)
        Line51(CNT) = 255 - Line51(CNT)
    Next
Else
    For CNT = StrLen To 0 Step -1
        Get #2, StartPos - CNT, Line51(CNT)
        Line51(CNT) = LowerLevel + Round(Lambda * Line51(CNT), 0)
        Line51(CNT) = 255 - Line51(CNT)
    Next
End If
    GetString51 = Base64Encode(Line51)
    GetString51 = Replace(GetString51, "/", "9", , , vbTextCompare)
    GetString51 = Replace(GetString51, "+", "9", , , vbTextCompare)
End Function

Sub GetBMPLine(ByVal LineNum As Long, Dir As Boolean)
Dim ChainCnt As Single
Dim prnDir As String
Dim Pointer As Long
Dim LCCalc As Long
Open OutFileName For Append As #5
        For ChainCnt = 1 To WholeChains
            If ChainCnt > 1 Then
                prnDir = ""
            ElseIf Dir = True Then
                prnDir = " $1"
            Else
                prnDir = " $0"
            End If
            If Dir = True Then
                Pointer = StartImagePointer + (LineNum - 1) * ScanLength + (ChainCnt - 1) * 51
                'Print #3, Pointer
            Else
                Pointer = StartImagePointer + (LineNum - 1) * ScanLength + ScanLength - (ChainCnt - 1) * 51
            End If
            Print #5, "G7" & prnDir & " L68 D"; GetString51(Pointer, 50, Dir, 2)
        Next
        If LastChain > 0 Then
            ChainCnt = ChainCnt + 1
            Pointer = StartImagePointer + (LineNum - 1) * ScanLength + WholeChains * 51  '*******
            'Print #3, Pointer
            Print #5, "G7" & prnDir & " L" & LastChainToPrn & " D"; GetString51(Pointer, LastChain - 1, Dir, 2)
        End If
    Print #5, ""
    Close #5
End Sub

Public Sub ExportBMPShape(ByVal PosX, ByVal PosY As Double, ByVal pwr As Byte, ByVal FR As Long)
Dim TempBmpFILE As String
Dim StringCNT As Double
Dim Dir As Boolean
Dim Offset As Byte

TempBmpFILE = "c:\temp\tmpbitmap.bmp"
GetInfoFromBmp TempBmpFILE
On Error Resume Next
'OutFileName = "D:\test2N2.g"
'Open TempBmpFILE For Binary As #2
Open OutFileName For Append As #5
Print #5, ";**************************Обработка Растра*****************************"
Print #5, ";****             Размеры изображения "; WidthPXL & "X"; HeightPXL
Print #5, ";*****************************************************************************"
Print #5, "M649 S" & pwr & " B2 D0 R" & BeamDia
Print #5, "G0 X" & DigToStr(PosX) & " Y" & DigToStr(PosY) & " F" & FR
Close #5

Mainform.ProgressBar.Max = HeightPXL
For StringCNT = 1 To HeightPXL
    If StringCNT Mod 2 <> 0 Then Dir = True Else Dir = False
    GetBMPLine StringCNT, Dir
    Mainform.ProgressBar = Mainform.ProgressBar + 1
Next StringCNT


Mainform.ProgressBar = 0
'MsgBox "Нормалды!"
End Sub
