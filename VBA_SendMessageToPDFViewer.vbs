Private Declare PtrSafe Function FindWindow Lib "USER32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare PtrSafe Function SendMessage Lib "USER32" Alias "SendMessageA" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As LongPtr
Option Explicit
Public Type CopyDataStruct
    dwData As LongPtr
    cbData As Long
    lpData As String
End Type
Private Const WM_COPYDATA = &H4A

Sub PreviewPDF(pdf_path)
    Dim cds As CopyDataStruct, result As LongPtr
    Static hwndTargetWin As Long
    Dim wParam As Long
    wParam = 0
    hwndTargetWin = FindWindow(vbNullString, "JustPDFViewer")
    If (hwndTargetWin > 0) Then
        cds.dwData = 1
        cds.cbData = Len(pdf_path) * 2 + 2 'The size, in bytes
        cds.lpData = pdf_path
        'Debug.Print pdf_path
        result = SendMessage(hwndTargetWin, WM_COPYDATA, wParam, cds)

    End If

End Sub


