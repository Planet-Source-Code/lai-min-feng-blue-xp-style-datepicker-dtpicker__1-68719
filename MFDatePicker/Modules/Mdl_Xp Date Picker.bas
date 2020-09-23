Attribute VB_Name = "Mdl_XPDatePicker"
Option Explicit

Public ResultDate As Long
Public ParentHwnd As Long
Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

Public Function GetDateSys() As String
    GetDateSys = DateSys(Date)
End Function

Public Function DateSys(StrDate As String) As String
 On Error GoTo errsystems
        DateSys = Format(StrDate, "yyyy") & "/" & Format(StrDate, "mm") & "/" & Format(StrDate, "dd")
errsystems:
End Function

Sub main()
    Load MyForm
End Sub
