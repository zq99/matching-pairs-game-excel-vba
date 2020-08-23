Attribute VB_Name = "modGamemoves"
'******************************
'Title:   Matchup Game
'Author:  Zaid Qureshi
'******************************

Option Explicit

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
Private Const csngInterval As Single = 0.5
Private Const cintFlash As Integer = 2
Private Const cintPlayerScore As Integer = 4
Private Const cintComputerScore As Integer = 8
Private Const cintScoreRow As Integer = 2
Private Const cintComputerGuessFromViewedProbability As Integer = 7

Private Sub HideGameTile(ByVal rngtile As Range)
    With rngtile
        .Value = Chr(cintCoveredIcon)
        .Interior.ColorIndex = cintCovered
        .Interior.Pattern = xlSolid
        .Font.Size = cintCoveredIconSize
        .Font.Name = "Arial"
    End With
    Range("a1").Select
End Sub

Public Sub GameMove(ByVal rngtile As Range, ByVal player As Boolean)
    Dim shtGame As Worksheet
    Dim rngBoard As Range
    
    On Error Resume Next
    Set shtGame = Worksheets("Match Board")
    With shtGame
        If Trim(.Cells(1, 14).Value) = cstrEmpty Then
           Call ShowTile(rngtile)
          .Cells(1, 16).Value = rngtile.Column
          .Cells(1, 15).Value = rngtile.row
          .Cells(1, 14).Value = rngtile.Value
        Else
            Call ShowTile(rngtile)
            If Trim(.Cells(1, 14).Value) = Trim(rngtile.Value) Then
                If player Then
                    .Range("Score").Value = .Range("Score").Value + 1
                    Call FlashingScore(True)
                    Call CheckScore
                Else
                    .Range("CompScore").Value = .Range("CompScore").Value + 1
                    Call FlashingScore(False)
                    Call CheckScore
                End If
            Else
                Call AddShortDelay
                Call HideGameTile(rngtile)
                Call HideGameTile(.Cells(.Cells(1, 15).Value, .Cells(1, 16).Value))
            End If
            .Cells(1, 16).Value = cstrEmpty
            .Cells(1, 15).Value = cstrEmpty
            .Cells(1, 14).Value = cstrEmpty
            If player Then
                Call ComputerMove
            End If
        End If
    End With
    Set shtGame = Nothing
    Set rngBoard = Nothing
End Sub

Public Sub ShowTile(ByVal rngtile As Range)
    Dim udtGamePosition As udtGamePosition
    Dim intCount As Integer
    Dim intIcon As Integer
    Dim strFileName As String
    Dim strMsg As String
    Dim strMsgTitle As String
    Dim intFile As Integer
    
    intCount = cintNone
    strFileName = ThisWorkbook.Path & "\GamePos.txt"
    intFile = FreeFile
    Open strFileName For Random As intFile Len = Len(udtGamePosition)
        Do While (Not EOF(intFile))
            intCount = intCount + 1
            Get intFile, intCount, udtGamePosition
            If (CInt(Trim(udtGamePosition.row)) = rngtile.row) And (CInt(Trim(udtGamePosition.col)) = rngtile.Column) Then
                intIcon = CInt(Trim(udtGamePosition.icon))
                Close intFile
                Exit Do
            End If
        Loop
    Close intFile
    
    If intIcon <> cintCoveredIcon Then
        With rngtile
            .Value = Chr(intIcon)
            .Interior.ColorIndex = cintUncovered
            .Interior.Pattern = xlSolid
            .Font.Size = cintCoveredIconSize
            .Font.Name = "Wingdings"
        End With
    End If
    
    If intIcon = cintCoveredIcon Then
        strMsg = "Press Start to initalize the game"
        strMsgTitle = "Press Start"
        MsgBox strMsg, vbOKOnly + vbExclamation, strMsgTitle
        Range("a1").Select
        End
    End If
    Range("a1").Select
    Call Memory(rngtile)
End Sub

Private Sub AddShortDelay(Optional ByVal DelayTime)
    Dim sngPauseTime As Single
    Dim sngStart As Single
    Dim sngFinish As Single
    
    If IsMissing(DelayTime) Then
        sngPauseTime = csngInterval
    Else
        sngPauseTime = DelayTime
    End If
    sngStart = Timer
    Do While Timer < sngStart + sngPauseTime
    Loop
    sngFinish = Timer
End Sub

Private Sub RandPosition(ByRef intRandRow As Integer, ByRef intRandColumn As Integer)
    Dim shtGame As Worksheet
    Dim intCount As Integer
    Set shtGame = Worksheets("Match Board")
    intCount = 1
    Do
        intRandRow = Int((cintMaxRow - cintMinRow + 1) * Rnd + cintMinRow)
        intRandColumn = Int((cintMaxColumn - cintMinColumn + 1) * Rnd + cintMinColumn)
        intCount = intCount + 1
    Loop Until (shtGame.Cells(intRandRow, intRandColumn).Interior.ColorIndex <> cintUncovered) Or intCount > 64
    Set shtGame = Nothing
End Sub

Private Sub ComputerMove()
    Dim intCompRow As Integer
    Dim intCompCol As Integer
    Dim shtGame As Worksheet
    Dim shtMemory As Worksheet
    Dim intTotal As Integer
    Dim intRow As Integer
    Dim intCol As Integer
    Dim intRandom As Integer
    Dim intIndex As Integer
    
    Set shtGame = Worksheets("Match Board")
    Set shtMemory = Worksheets("memory")
    
    Call YourMoveMessageRequired(shtGame, False)
    intTotal = shtMemory.Cells(2, 1).CurrentRegion.Rows.Count
    Call RandPosition(intCompRow, intCompCol)
    Call GameMove(shtGame.Cells(intCompRow, intCompCol), False)
    
    intRandom = Int((10 - 1 + 1) * Rnd + 1)
    If intRandom <= cintComputerGuessFromViewedProbability Then
        For intIndex = 2 To intTotal
            If Trim(shtMemory.Cells(intIndex, 5)) <> Trim(intCompRow & intCompCol) Then
                If Trim(shtMemory.Cells(intIndex, 3).Value) = Trim(shtGame.Cells(intCompRow, intCompCol).Value) Then
                    intRow = shtMemory.Cells(intIndex, 1).Value
                    intCol = shtMemory.Cells(intIndex, 2).Value
                    If shtGame.Cells(intRow, intCol).Interior.ColorIndex <> cintUncovered Then
                        Call GameMove(shtGame.Cells(intRow, intCol), False)
                        Call YourMoveMessageRequired(shtGame, True)
                        Set shtGame = Nothing
                        Set shtMemory = Nothing
                        Exit Sub
                    End If
                End If
            End If
        Next
    End If
    Call RandPosition(intCompRow, intCompCol)
    Call GameMove(shtGame.Cells(intCompRow, intCompCol), False)
    Call YourMoveMessageRequired(shtGame, True)
    
    Set shtGame = Nothing
    Set shtMemory = Nothing
End Sub

Public Sub YourMoveMessageRequired(ByVal sht As Worksheet, ByVal blnShow As Boolean)
    Dim rng As Range
    With sht
        Set rng = .Range("B1")
        If blnShow Then
            With rng.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent3
                .TintAndShade = 0.799981688894314
                .PatternTintAndShade = 0
            End With
            rng.Value = "Your Move!"
        Else
            With rng.Interior
                .Pattern = xlNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            rng.Value = ""
        End If
    End With
    Set rng = Nothing
End Sub

Private Sub CheckScore()
    Dim shtGame As Worksheet
    Dim intScoreDiff As Integer
    Dim strMsg As String
    Dim strMsgTitle As String
    
    Set shtGame = Worksheets("Match Board")
    With shtGame
        intScoreDiff = Abs(.Range("Score").Value - .Range("CompScore").Value)
        If intScoreDiff > fnPointsLeft() Then
            If .Range("Score").Value > .Range("CompScore").Value Then
                frmWin.lblname.Caption = fnGetUserLogin
                frmWin.Show
                Call Reset
                End
            End If
            If .Range("CompScore").Value > .Range("Score").Value Then
                strMsgTitle = "GAME OVER!"
                strMsg = "I have beaten you!"
                MsgBox strMsg, vbOKOnly + vbExclamation, strMsgTitle
                Call Reset
                End
            End If
        End If
        If fnPointsLeft = cintNone Then
            If .Range("CompScore").Value = .Range("Score").Value Then
                strMsgTitle = "DRAW"
                strMsg = "The game is drawn, you are a worthy opponent!"
                MsgBox strMsg, vbOKOnly + vbExclamation, strMsgTitle
                Call Reset
                End
            End If
         End If
    End With
    Set shtGame = Nothing
End Sub

Private Function fnPointsLeft() As Single
    Dim shtGame As Worksheet
    Dim rngBoard As Range
    Dim intCount As Integer
    Dim rngCell As Range
    
    fnPointsLeft = 0
    Set shtGame = Worksheets("Match Board")
    Set rngBoard = shtGame.Range("board")
    intCount = cintNone
    With shtGame
        For Each rngCell In rngBoard
            If rngCell.Interior.ColorIndex = cintCovered Then
                intCount = intCount + 1
            End If
        Next rngCell
    End With
    fnPointsLeft = intCount / 2
    Set shtGame = Nothing
    Set rngBoard = Nothing
    Set rngCell = Nothing
End Function

Private Sub Memory(ByVal rng)
    Dim shtMemory As Worksheet
    Dim intNextRow As Integer
    Dim rngFindRef As Range
    
    Set shtMemory = Worksheets("memory")
    intNextRow = shtMemory.Cells(1, 1).CurrentRegion.Rows.Count + 1
    Set rngFindRef = shtMemory.Range(shtMemory.Cells(1, 5), shtMemory.Cells(intNextRow, 5)) _
                      .Find(what:=rng.row & rng.Column)
    
    If rngFindRef Is Nothing Then
        With shtMemory
          .Cells(intNextRow, 1).Value = rng.row
          .Cells(intNextRow, 2).Value = rng.Column
          .Cells(intNextRow, 3).Value = rng.Value
          .Cells(intNextRow, 4).Value = rng.Value
          .Cells(intNextRow, 5).Value = rng.row & rng.Column
        End With
    End If
    Set shtMemory = Nothing
    Set rngFindRef = Nothing
End Sub

Private Sub FlashingScore(ByVal bolplayer As Boolean)
    Dim shtGame As Worksheet
    Dim intCount As Integer
    Dim intCol As Integer
    
    Set shtGame = Worksheets("Match Board")
    If bolplayer Then intCol = cintPlayerScore _
                Else intCol = cintComputerScore
          
    With shtGame
        For intCount = 1 To 6
          .Cells(cintScoreRow, intCol).Interior.ColorIndex = cintFlash
           Call AddShortDelay(0.05)
          .Cells(cintScoreRow, intCol).Interior.ColorIndex = cintUncovered
        Next
    End With
    Set shtGame = Nothing
End Sub

Public Function fnGetUserLogin() As String
    fnGetUserLogin = Environ("username")
End Function
