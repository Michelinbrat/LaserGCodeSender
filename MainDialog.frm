VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainDialog 
   Caption         =   "UserForm1"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11085
   OleObjectBlob   =   "MainDialog.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MainDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Base 1

Public Sub MakePreview(sr As ShapeRange)
Dim expflt As ExportFilter
Dim Xpxl, Ypxl As Long
Dim Resolution As Integer: Resolution = 96
Dim pal As StructPaletteOptions
Dim w As Double, h As Double
Dim Scale1 As Integer: Scale1 = 50

sr.GetSize w, h
    Set pal = New StructPaletteOptions
    With pal
        .PaletteType = cdrPaletteOptimized
        .DitherType = cdrDitherOrdered
        .DitherIntensity = 300
    End With
   
    Set expflt = ActiveDocument.ExportBitmap("C:\temp\taskscreen.bmp", cdrBMP, cdrSelection, cdrPalettedImage, _
    w * Scale1, h * Scale1, Resolution, Resolution, cdrNormalAntiAliasing, False, False, False, False, cdrCompressionNone, pal)
    expflt.Finish
    Image2.PictureAlignment = fmPictureAlignmentCenter
Image2.PictureSizeMode = fmPictureSizeModeZoom
Image2.Picture = LoadPicture("c:\Temp\taskscreen.bmp")
'Close
End Sub
Sub CreateTaskList()
Dim k As Byte
Dim Offset As Integer: Offset = 16
Dim Top As Integer: Top = 40
Dim Left As Integer: Left = 300
Dim BTColor As Long
Dim Capt As String

For k = 1 To UBound(TASK)
If TASK(k).IsGrav = True Then
    With TASK(k)
        .Flip = False
        .Invert = False
        .Repeat = 1
        .LaserMode = lmPULSE
        .OutlineColor = 16777215
    End With
    Capt = "ENGR"
    Else
    Capt = "CUT"
End If

BTColor = TASK(k).OutlineColor

'1 ------- Использовать--------------
Set newCBuse = Me.Controls.Add("Forms.checkbox.1")
With newCBuse
    .Name = "Use" & k
    .Caption = ""
    .Top = Top + (k * Offset)
    .Left = Left
    .Width = 14
    .Height = 16
    .Value = True
    .Font.Size = 8
    .Font.Bold = False
    .Font.Name = "Tahoma"
    .BackColor = BTColor
End With

'2-------Порядковый номер обработки-----------
Set newTBorder = Me.Controls.Add("forms.textbox.1")
With newTBorder
    .Name = "Order" & k
    .Top = Top + (k * Offset)
    .Left = Left + 18
    .Width = 22
    .Height = 16
    .Value = k
    .Font.Size = 8
    .Font.Name = "Tahoma"
End With
'3-------- Режим обработки -------------------
Set newLblbMode = Me.Controls.Add("Forms.label.1")
With newLblbMode
    .Name = "Mode" & k
    .Caption = Capt
    .Top = Top + (k * Offset)
    .Left = Left + 42
    .Width = 28
    .Height = 16
    .BackColor = TASK(k).OutlineColor
    .TextAlign = 2
    .Font.Name = "Tahoma"
    .Font.Size = 10
    .Font.Bold = True
    If TASK(k).IsGrav = False Then .ForeColor = 13158600
End With
'4-------- Мощность -------------------

Set newTBpwr = Me.Controls.Add("forms.textbox.1")
With newTBpwr
    .Name = "PWR" & k
    .Top = Top + (k * Offset)
    .Left = Left + 72
    .Width = 30
    .Height = 16
    .Value = TASK(k).Power
    .Font.Size = 8
    .Font.Name = "Tahoma"
End With
'5-------- Скорость -------------------
Set newTBfeed = Me.Controls.Add("forms.textbox.1")
With newTBfeed
    .Name = "Feed" & k
    .Top = Top + (k * Offset)
    .Left = Left + 102
    .Width = 30
    .Height = 16
    .Value = TASK(k).Feed
    .Font.Size = 8
    .Font.Bold = False
    .Font.Name = "Tahoma"
End With
'6-------- Разрешение -------------------
Set newTBres = Me.Controls.Add("forms.textbox.1")
With newTBres
    .Name = "Res" & k
    .Top = Top + (k * Offset)
    .Left = Left + 132
    .Width = 30
    .Height = 16
    .Value = TASK(k).Resolution
    .Font.Size = 8
    .Font.Bold = False
    .Font.Name = "Tahoma"
End With
'7-------- Число проходов -------------------
Set newTBRepeat = Me.Controls.Add("forms.textbox.1")
With newTBRepeat
    .Name = "Repeat" & k
    .Top = Top + (k * Offset)
    .Left = Left + 162
    .Width = 24
    .Height = 16
    .Value = TASK(k).Repeat
    .Font.Size = 8
    .Font.Bold = False
    .Font.Name = "tahoma"
End With
'8 ------- Инвертировать--------------
Set newCBInvert = Me.Controls.Add("Forms.checkbox.1")
With newCBInvert
    .Name = "Invert" & k
    .Caption = ""
    .Top = Top + (k * Offset)
    .Left = Left + 192
    .Width = 14
    .Height = 12
    .Visible = TASK(k).IsGrav
    .Value = False
    .Font.Size = 8
    .Font.Bold = False
    .Font.Name = "tahoma"
End With
'9 ------- Отзеркалить--------------
Set newCBFlip = Me.Controls.Add("Forms.checkbox.1")
With newCBFlip
    .Name = "Flip" & k
    .Caption = ""
    .Top = Top + (k * Offset)
    .Left = Left + 210
    .Width = 14
    .Height = 12
    .Visible = TASK(k).IsGrav
    .Value = False
    .Font.Size = 8
    .Font.Bold = False
    .Font.Name = "tahoma"
End With
Next k
End Sub



Private Sub CheckBox6_Click()

End Sub

Private Sub CommandButton3_Click()
DelTemplayer
MakePreview ActiveSelectionRange.All

CopySRtoTempLayer ActiveSelectionRange
CreateTaskList
End Sub



Private Sub CommandButton4_Click()

End Sub

Private Sub lblMarkCol_Click()

End Sub

Private Sub UserForm_Initialize()
'ActiveDocument.Unit = cdrMillimeter
ActiveDocument.ReferencePoint = cdrBottomLeft
'Application.Optimization = True

End Sub

Private Sub UserForm_Terminate()
'Application.Optimization = False
EngraveSR.Delete: CutSR.Delete
DelTemplayer
End Sub
