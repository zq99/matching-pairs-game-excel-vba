VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit
Private Const cintMinRow As Integer = 4
Private Const cintMaxRow As Integer = 11
Private Const cintMinColumn As Integer = 2
Private Const cintMaxColumn As Integer = 9
Private Const cintUncovered As Integer = 6

Private Sub Worksheet_BeforeDoubleClick(ByVal rngCell As Range, Cancel As Boolean)
    Call Verifyclick(rngCell)
    Call GameMove(rngCell, True)
End Sub

Private Sub Verifyclick(ByVal rngCell As Range)
    Dim strMsg As String
    Dim strMsgTitle As String

    strMsg = strMsg & "Your selection must be within the game board."
    strMsgTitle = "Invalid"
    If rngCell.row < cintMinRow Or rngCell.row > cintMaxRow Then
        MsgBox strMsg, vbOKOnly + vbExclamation, strMsgTitle
        Range("a1").Select
        End
    End If
    
    If rngCell.Column < cintMinColumn Or rngCell.Column > cintMaxColumn Then
        MsgBox strMsg, vbOKOnly + vbExclamation, strMsgTitle
        Range("a1").Select
        End
    End If
    
    With rngCell
        If .Interior.ColorIndex = cintUncovered Then
            strMsg = "You can't select that uncovered tile."
            MsgBox strMsg, vbOKOnly + vbExclamation, strMsgTitle
            Range("a1").Select
            End
        End If
    End With
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim intRow As Integer
    Dim intCol As Integer
    
    intRow = Target.row
    intCol = Target.Column
    If intRow = 2 Then
        If intCol = 4 Or intCol = 8 Then
            MsgBox "You can't fix the Scores you Cheat!", vbCritical, "No Cheating!"
            Range("a1").Select
        End If
    End If
End Sub
