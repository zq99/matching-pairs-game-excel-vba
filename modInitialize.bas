Attribute VB_Name = "modInitialize"
'******************************
'Title:   Matchup Game
'Author:  Zaid Qureshi
'******************************

Option Explicit
Option Base 1

Private Const cintCovered As Integer = 45
Private Const cintCoveredIcon As Integer = 127
Private Const cintCoveredIconSize As Integer = 28
Private Const cstrEmpty As String = ""
Private Const cintUncovered As Integer = 6
Private Const cintNone As Integer = 0
Private Const cintMinRow As Integer = 4
Private Const cintMaxRow As Integer = 11
Private Const cintMinColumn As Integer = 2
Private Const cintMaxColumn As Integer = 9
Private Const cintNumberOfSquares As Integer = 64
Private Const cintFirstSquare As Integer = 1
Private Const cintFlash As Integer = 2
Private Const cintPlayerScore As Integer = 4
Private Const cintComputerScore As Integer = 8
Private Const cintScoreRow As Integer = 2

Public Type udtGamePosition
    ID As Integer
    row As Integer
    col As Integer
    icon As Integer
End Type

Public Sub StartGame()
    If fnGameInProgress Then
        Call ResetGame
        frmSplashScreen.Show
    Else
        Call Reset
        frmSplashScreen.Show
    End If
End Sub

Public Sub ResetGame()
    Dim strMsg As String
    Dim strMsgTitle As String
    Dim strPrompt As String
    
    If fnGameInProgress Then
        strMsg = "Are you sure you want to reset?"
        strMsgTitle = "Reset"
        strPrompt = MsgBox(strMsg, vbExclamation + vbYesNo, strMsgTitle)
        Select Case strPrompt
        Case vbYes
            Call Reset
        Case vbNo
            End
        End Select
    Else
        Call Reset
    End If
End Sub

Public Sub Reset()
On Error GoTo errHandler:
    
    Dim strFileName As String
    Dim udtGamePosition As udtGamePosition
    Dim intCount As Integer
    Dim intFile As Integer
    Dim blnScreen As Boolean
    blnScreen = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    Call YourMoveMessageRequired(Worksheets("Match Board"), False)
    Call ClearBoard
    strFileName = ThisWorkbook.Path & "\GamePos.txt"
    intCount = cintNone
    udtGamePosition.row = cintMinRow
    udtGamePosition.col = cintMinColumn
    intFile = FreeFile
    Open strFileName For Random As intFile Len = Len(udtGamePosition)
        Do While intCount <= cintNumberOfSquares
            intCount = intCount + 1
            udtGamePosition.icon = cintCoveredIcon
            Put intFile, intCount, udtGamePosition
            If udtGamePosition.col >= cintMaxColumn Then
                udtGamePosition.col = cintMinColumn
                udtGamePosition.row = udtGamePosition.row + 1
            Else
                udtGamePosition.col = udtGamePosition.col + 1
            End If
        Loop
    Close intFile
exitHere:
    Application.ScreenUpdating = blnScreen
    Exit Sub
errHandler:
    GoTo exitHere
End Sub

Public Sub InitializeGame()
    
    Dim vntIconArray As Variant
    Dim arrIcon(cintNumberOfSquares) As Variant
    Dim intRandom As Integer
    Dim arrCount As Integer
    Dim intRow As Integer
    Dim intCol As Integer
    Dim strFileName As String
    Dim udtGamePosition As udtGamePosition
    Dim intNumberOfRecords As Integer
    Dim intFile As Integer
    
    vntIconArray = Array(33, 33, 34, 34, 36, 36, 38, 38, 37, 37, 40, 40, 42, 42, 46, 46, 54, 54, 55, _
                         55, 57, 57, 58, 58, 64, 64, 65, 65, 74, 74, 76, 76, 77, 77, 78, 78, 81, 81, _
                         82, 82, 91, 91, 165, 165, 171, 171, 174, 174, 177, 177, 181, 181, 190, 190, _
                         185, 185, 216, 216, 219, 219, 253, 253, 255, 255)
    
    On Error Resume Next
    Call ClearBoard
    On Error GoTo 0
    
    udtGamePosition.row = cintMinRow
    udtGamePosition.col = cintMinColumn
    arrCount = cintFirstSquare
    strFileName = ThisWorkbook.Path & "\GamePos.txt"
    Do While arrCount <= cintNumberOfSquares
        intRandom = Int((cintNumberOfSquares - cintFirstSquare + 1) * Rnd + 1)
        If fnSearchArray(intRandom, arrIcon) = False Then
            udtGamePosition.icon = vntIconArray(intRandom)
            udtGamePosition.ID = arrCount
            intFile = FreeFile
            Open strFileName For Random As intFile Len = Len(udtGamePosition)
                Put intFile, udtGamePosition.ID, udtGamePosition
            Close intFile
            arrIcon(arrCount) = intRandom
            arrCount = arrCount + 1
            If udtGamePosition.col >= cintMaxColumn Then
                udtGamePosition.col = cintMinColumn
                udtGamePosition.row = udtGamePosition.row + 1
            Else
                udtGamePosition.col = udtGamePosition.col + 1
            End If
        End If
    Loop
    Call YourMoveMessageRequired(Worksheets("Match Board"), True)
End Sub
                     
Public Function fnSearchArray(ByVal criteria As Variant, ByVal ArraytoSearch As Variant) As Boolean
    Dim intMin As Integer, intMax As Integer, intIndex As Integer
    
    intMin = LBound(ArraytoSearch)
    intMax = UBound(ArraytoSearch)
    For intIndex = intMin To intMax
        If ArraytoSearch(intIndex) = criteria Then
            fnSearchArray = True
            Exit Function
        End If
    Next intIndex
    fnSearchArray = False
    
End Function

Public Sub ClearBoard()
    Dim shtGame As Worksheet
    Dim shtMemory As Worksheet
    Dim rngGame As Range
    Dim rngCell As Range
    
    Set shtGame = Worksheets("Match Board")
    Set shtMemory = Worksheets("memory")
    Set rngGame = shtGame.Range("board")
    With shtGame
        .Range("Score").Value = cintNone
        .Range("Score").Interior.ColorIndex = cintUncovered
        .Range("CompScore").Value = cintNone
        .Range("CompScore").Interior.ColorIndex = cintUncovered
        For Each rngCell In rngGame
            With rngCell
              .Value = Chr(cintCoveredIcon)
              .Interior.ColorIndex = cintCovered
              .Interior.Pattern = xlSolid
              .Font.Size = cintCoveredIconSize
              .Font.Name = "Arial"
            End With
         Next rngCell
        .Range("last_col").Value = cstrEmpty
        .Range("last_row").Value = cstrEmpty
        .Range("last_icon").Value = cstrEmpty
    End With
    shtMemory.Cells.ClearContents
    Set shtGame = Nothing
    Set shtMemory = Nothing
    Set rngGame = Nothing
    Set rngCell = Nothing
End Sub

Private Function fnGameInProgress() As Boolean
    Dim shtGame As Worksheet
    Dim rngGame As Range
    Dim rngCell As Range
    
    Set shtGame = Worksheets("Match Board")
    Set rngGame = shtGame.Range("board")
    fnGameInProgress = False
    With shtGame
        For Each rngCell In rngGame
            If rngCell.Interior.ColorIndex = cintUncovered Then
                fnGameInProgress = True
                Exit Function
            End If
        Next rngCell
    End With
    Set shtGame = Nothing
    Set rngGame = Nothing
    Set rngCell = Nothing
End Function

Public Sub AboutGame()
    Dim strMsg As String
    Dim strMsgTitle As String
    
    strMsgTitle = "About"
    strMsg = "Match UP - Programmed by Zaid Qureshi"
    MsgBox strMsg, vbInformation, strMsgTitle
End Sub





