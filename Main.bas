Attribute VB_Name = "Main"
Option Base 1
Option Explicit

Public Type Lasertask
Order As Byte
sr As New ShapeRange
Repeat As Byte
Power As Byte
Feed As Integer
Resolution As Integer
Flip As Boolean
Invert As Boolean
IsUSE As Boolean
IsGrav As Boolean
OutlineColor As Long
PosX As Double
PosY As Double
LaserMode As LMode
End Type

Const RESFILEPATH = "C:\Temp"
Public TASK() As Lasertask

Public Enum LMode
lmCONTINUOUS = 0
lmPULSE = 1
End Enum
