Attribute VB_Name = "MdlSample"
Option Explicit
Private Declare Function QueryPerformanceCounter Lib "kernel32.dll" (ByRef lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32.dll" (ByRef lpFrequency As Currency) As Long


Sub Main()
    FrmGUIevaluator.Show
    'Form1.Show
    'Stop
    'MsgBox Join((QSort(Array(3, 2, 1, 2, 3), 0, 4)), ",")
    
End Sub

Public Function Timer() As Double
Static secFreq As Currency, secStart As Currency
If secFreq = 0 Then QueryPerformanceFrequency secFreq
QueryPerformanceCounter secStart
If secFreq Then Timer = secStart / secFreq Else Timer = 0
'if no high resolution timer
End Function

