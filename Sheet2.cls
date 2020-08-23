VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

Private Sub Worksheet_Activate()
    Dim strMsg As String
    Dim strMsgTitle As String

    Worksheets("memory").Visible = False
    strMsgTitle = "NO CHEATING"
    strMsg = "This sheet is restricted, for game purposes"
    MsgBox strMsg, vbCritical, strMsgTitle
End Sub

