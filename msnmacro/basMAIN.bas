Attribute VB_Name = "basMAIN"
Option Explicit
Private Declare Sub InitCommonControls Lib "comctl32" ()
Public Sub Main()
    On Error GoTo ErrorMain
    InitCommonControls
    frmMAIN.Show
    Exit Sub
ErrorMain:
    MsgBox Err & ":Error in Main.  Error Message: " _
    & Err.Description, vbCritical, "Warning"
    Exit Sub
End Sub
