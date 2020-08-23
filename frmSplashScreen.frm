VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSplashScreen 
   Caption         =   "Match Up"
   ClientHeight    =   1510
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4380
   OleObjectBlob   =   "frmSplashScreen.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSplashScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdcancel_Click()
    Hide
End Sub

Private Sub cmdstart_Click()
    Dim strMsg As String
    Dim strMsgTitle As String
    Dim strUser As String
    
    Call Reset
    Call InitializeGame
    Hide
    strUser = UCase(fnGetUserLogin)
    strMsg = "Please make the first move."
    strMsg = strMsg & Chr(10) & Chr(10) & "To make a selection double click on a square."
    strMsg = strMsg & Chr(10) & Chr(10) & "You need to select TWO squares for each turn."
    strMsgTitle = "Match Up"
    MsgBox strMsg, vbOKOnly, strMsgTitle
End Sub


