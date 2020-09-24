Attribute VB_Name = "General"
'-------------------------------------------------------------------------------------------'
'    This Program and the code contained within is Copyright of Infostrategy Ltd. 2000
'    Any attempt to copy this program or any portion of it without first consulting the
'    Company will be met with deadly force and/or a lawsuit (particularly if copied with
'    Comercial gain in mind) -Viper
'-------------------------------------------------------------------------------------------'

Option Explicit

Public Function CheckMisalignment(ShowErrors As Boolean) As Long
Dim Player As Long, PieceN As Long, Indices As Long

For Player = 1 To 2
  For PieceN = 1 To 12
    If CurrentBoard.Pieces(Player, PieceN) <> 0 Then
      If PieceN <> (CurrentBoard.Fields((CurrentBoard.Pieces(Player, PieceN) And IndexMask)) And PieceNumMask) Then
        If ShowErrors Then MsgBox "Misalignment has occured in Player " & Player & " Piece " & PieceN, vbCritical
        CheckMisalignment = 1
      End If
    Else
      For Indices = 1 To 100
        If CurrentBoard.Fields(Indices) And PieceNumMask = PieceN Then
          If ShowErrors Then MsgBox "Misalignment has occured in Player " & Player & " Piece " & PieceN, vbCritical
          CheckMisalignment = 1
        End If
      Next Indices
    End If
  Next PieceN
Next Player

End Function

Public Function PieceImageNum(Field As Byte) As Long
If Field And P1Mask Then
  If Field And DoubleMask Then PieceImageNum = 2 Else PieceImageNum = 1
Else
  If Field And DoubleMask Then PieceImageNum = 4 Else PieceImageNum = 3
End If
End Function


Public Sub RefreshBoard(Board As Board)
Dim Lng1 As Long, Index As Long

For Lng1 = 1 To 100    'for every index
  If (Board.Fields(Lng1) And InvalidSquare) = False Then
    If frmMain.Shape1(IndexTranslation(Lng1)).Picture.Handle <> 0 And Board.Fields(Lng1) = 0 Then 'if there is a picture but shouldn't be
      frmMain.Shape1(IndexTranslation(Lng1)).Picture = Nothing
    ElseIf Board.Fields(Lng1) Then
      Index = IndexTranslation(Lng1)          'Set index to true index of the board
      If Board.Fields(Lng1) And P1Mask Then          'If there is a player one piece on that square
        If Board.Fields(Lng1) And DoubleMask Then    'If it is a double
          frmMain.Shape1(Index).Picture = frmMain.ImageList1.ListImages(2).Picture    'Set picture of square to that of the piece
        Else
          frmMain.Shape1(Index).Picture = frmMain.ImageList1.ListImages(1).Picture
        End If
      ElseIf Board.Fields(Lng1) And P2Mask Then
        If Board.Fields(Lng1) And DoubleMask Then
          frmMain.Shape1(Index).Picture = frmMain.ImageList1.ListImages(4).Picture
        Else
          frmMain.Shape1(Index).Picture = frmMain.ImageList1.ListImages(3).Picture
        End If
      End If
    
    End If
  End If
Next

End Sub

Public Sub ResetGame()
Dim Lng1 As Long, Lng2 As Long

GameStarted = False

With frmMain
  .lblMoves = "0"
  .lblCutoffs = "0"
  .lblP1Time = "0 Min 0 Sec"
  .lblP2Time = "0 Min 0 Sec"
  .lblTurns = "0"
  .cmdBack.Enabled = False
End With

VP1Time = 0
VP2Time = 0
VTurns = 0

With CurrentBoard
  For Lng1 = 12 To 89
      .Fields(Lng1) = 0
  Next
  For Lng1 = 1 To 10 'fill guard fields
      .Fields(Lng1) = InvalidSquare
      .Fields(Lng1 + 90) = InvalidSquare
      .Fields(Lng1 * 10 - 9) = InvalidSquare
      .Fields(Lng1 * 10) = InvalidSquare
  Next
  .Pieces(Player2, 1) = 62        'Place pieces on board and in piecelists
  .Fields(62) = P2Mask Or 1
  .Pieces(Player2, 2) = 64
  .Fields(64) = P2Mask Or 2
  .Pieces(Player2, 3) = 66
  .Fields(66) = P2Mask Or 3
  .Pieces(Player2, 4) = 68
  .Fields(68) = P2Mask Or 4
  .Pieces(Player2, 5) = 73
  .Fields(73) = P2Mask Or 5
  .Pieces(Player2, 6) = 75
  .Fields(75) = P2Mask Or 6
  .Pieces(Player2, 7) = 77
  .Fields(77) = P2Mask Or 7
  .Pieces(Player2, 8) = 79
  .Fields(79) = P2Mask Or 8
  .Pieces(Player2, 9) = 82
  .Fields(82) = P2Mask Or 9
  .Pieces(Player2, 10) = 84
  .Fields(84) = P2Mask Or 10
  .Pieces(Player2, 11) = 86
  .Fields(86) = P2Mask Or 11
  .Pieces(Player2, 12) = 88
  .Fields(88) = P2Mask Or 12
  
  .Pieces(Player1, 1) = 33
  .Fields(33) = P1Mask Or 1
  .Pieces(Player1, 2) = 35
  .Fields(35) = P1Mask Or 2
  .Pieces(Player1, 3) = 37
  .Fields(37) = P1Mask Or 3
  .Pieces(Player1, 4) = 39
  .Fields(39) = P1Mask Or 4
  .Pieces(Player1, 5) = 22
  .Fields(22) = P1Mask Or 5
  .Pieces(Player1, 6) = 24
  .Fields(24) = P1Mask Or 6
  .Pieces(Player1, 7) = 26
  .Fields(26) = P1Mask Or 7
  .Pieces(Player1, 8) = 28
  .Fields(28) = P1Mask Or 8
  .Pieces(Player1, 9) = 13
  .Fields(13) = P1Mask Or 9
  .Pieces(Player1, 10) = 15
  .Fields(15) = P1Mask Or 10
  .Pieces(Player1, 11) = 17
  .Fields(17) = P1Mask Or 11
  .Pieces(Player1, 12) = 19
  .Fields(19) = P1Mask Or 12
  .Score = 0
  .Turn = Player1
End With

MoveListChanged = True
ReDim BoardHistory(1 To 1)
BoardHistory(1) = CurrentBoard

Call RefreshDisplay
Call RefreshBoard(CurrentBoard)

GetSettings

Score(1) = 12
Score(2) = 12

End Sub

Public Sub CheckWin()
Dim Lng1 As Long, FieldNum As Long

MaxChainLength = 1
CurrentBoard.MovesListFrom = 1
ResetIx = 0
InPV = 0

For Lng1 = 1 To 12
  FieldNum = CurrentBoard.Pieces(CurrentBoard.Turn, Lng1) And IndexMask
  If FieldNum Then GenerateMoves CurrentBoard, FieldNum, 0, "", 0
Next

If MaxChainLength < 2 Or MoveListIx = 0 Then
  MsgBox Names(BothSides - CurrentBoard.Turn) & " wins!"
  Call ResetGame
End If

End Sub

Public Function MovePiece(MoveChain As String) As Long
Dim Lng1 As Long, FieldNum As Long

If MoveListIx = 0 Or MoveListChanged Then

MaxChainLength = 1
CurrentBoard.MovesListFrom = 1
ResetIx = 0
InPV = 0

For Lng1 = 1 To 12
  FieldNum = CurrentBoard.Pieces(CurrentBoard.Turn, Lng1) And IndexMask
  If FieldNum Then GenerateMoves CurrentBoard, FieldNum, 0, "", 0
Next

If MaxChainLength < 2 Or MoveListIx = 0 Then
  MsgBox Names(BothSides - CurrentBoard.Turn) & " wins!"
  Call ResetGame
End If

MoveListChanged = False

End If

For Lng1 = 1 To MoveListIx
  If MoveList(Lng1) = MoveChain Then
    MovePiece = MoveCompleted Or MoveCorrect
    MakeMove CurrentBoard, MoveChain, True
    CurrentBoard.Turn = BothSides - CurrentBoard.Turn
    ReDim Preserve BoardHistory(1 To UBound(BoardHistory) + 1)
    BoardHistory(UBound(BoardHistory)) = CurrentBoard
    frmMain.cmdBack.Enabled = True
    Exit Function
  ElseIf Mid$(MoveList(Lng1), 1, Len(MoveChain)) = MoveChain Then
    MovePiece = MoveCorrect
    Exit Function
  End If
Next

End Function

Public Sub MakeMove(Board As Board, ByVal MoveChain As String, ForReal As Boolean)
Dim FromField As Long, ToField As Long, Captured As Long

'The movechain is split up into groups of 3 characters. Where it came from, Where it's going to, and where the piece was which was taken was
'Every loop does another one of these groups (if the chain length is higher than 3)

Do
  FromField = Asc(Mid$(MoveChain, 1))
  ToField = Asc(Mid$(MoveChain, 2))
  If Len(MoveChain) < 4 And ForReal Then
    If ((Board.Turn = Player1 And ToField > 81) Or (Board.Turn = Player2 And ToField < 18)) And (Board.Fields(FromField) And DoubleMask) = 0 Then
      Board.Fields(FromField) = Board.Fields(FromField) Or DoubleMask
      Board.Score = Board.Score + ((SingleValue - Doublevalue) * ((Board.Turn * 2) - BothSides))
    End If
  End If
  If Len(MoveChain) > 2 Then
    Captured = Asc(Mid$(MoveChain, 3))
    If Board.Fields(Captured) And DoubleMask Then
      Board.Score = Board.Score + (Doublevalue * (BothSides - (Board.Turn * 2))) 'Effectively minusing the value if it's player1 turn (and visaversa)
    Else
      Board.Score = Board.Score + (SingleValue * (BothSides - (Board.Turn * 2)))
    End If
    Board.Pieces(BothSides - Board.Turn, Board.Fields(Captured) And PieceNumMask) = 0
    Board.Fields(Captured) = 0
  End If
  
  Board.Fields(ToField) = Board.Fields(FromField)
  Board.Fields(FromField) = 0
  Board.Pieces(Board.Turn, Board.Fields(ToField) And PieceNumMask) = ToField Or (DoubleMask And Board.Fields(ToField))
  
  MoveChain = Mid$(MoveChain, 4)
Loop While Len(MoveChain) >= 3

End Sub


Public Sub RefreshDisplay()
Dim PieceN As Long

  Score(1) = 0: Score(2) = 0

  If CurrentBoard.Turn = 1 Then
    frmMain.lblTurn = Names(1)
  Else
    frmMain.lblTurn = Names(2)
  End If
  
  For PieceN = 1 To 100
    If CurrentBoard.Fields(PieceN) And P1Mask Then
      Score(1) = Score(1) + 1
    ElseIf CurrentBoard.Fields(PieceN) And P2Mask Then
      Score(2) = Score(2) + 1
    End If
  Next
  
  frmMain.lblP1Points = Names(1) & " - " & 12 - Score(2)
  frmMain.lblP2Points = Names(2) & " - " & 12 - Score(1)
  
  If Score(1) = 0 Then
    If AutoDebug = False Then
      MsgBox Names(2) & " wins!", vbExclamation
    End If
    ResetGame
  ElseIf Score(2) = 0 Then
    If AutoDebug = False Then MsgBox Names(1) & " wins!", vbExclamation
    ResetGame
  End If
  
  frmMain.lblP1Time = IIf(((VP1Time Mod 60) / 60) >= 0.5, CLng(VP1Time / 60) - 1, CLng(VP1Time / 60)) & " Min " & VP1Time Mod 60 & " Sec"
  frmMain.lblP2Time = IIf(((VP2Time Mod 60) / 60) >= 0.5, CLng(VP2Time / 60) - 1, CLng(VP2Time / 60)) & " Min " & VP2Time Mod 60 & " Sec"
  
  frmMain.lblTurns = VTurns
  
End Sub

Public Sub XYConvert(Index As Long, ByRef X As Long, ByRef Y As Long)
  If Index > 63 Or Index < 0 Then X = 0: Y = 0: Exit Sub   'Ifindex is invalid then exit
  Y = (Index - (Index Mod 8)) / 8 + 1             'Get Y co-ordinate from index
  X = (Index Mod 8) + 1                            'Get X co-ordinate from index and y co-ordinate
End Sub

Public Function IndexTranslation(ByVal Index As Long, Optional Normal As Boolean, Optional Inverse As Boolean) As Long
Dim Y As Long, X As Long
  If Normal = False Then
    If Inverse = True Then 'Big to small index
      If Reversed = True Then                             'If the board is reserved the return the inverse index
        Index = 63 - Index
      Else
        Index = Index
      End If
      Y = (Index - (Index Mod 8)) / 8 + 1               'Get row number of 'board' type board (10 by 10)
      X = (Index Mod 8) + 1                              'Get column number of 'board' type board (10 by 10)
      If Y < 1 Or Y > 8 Or X < 1 Or X > 8 Then IndexTranslation = InvalidSquare: Exit Function  'Check if it will map to normal 8 by 8 board
      IndexTranslation = (Index + 11) + (2 * (Y - 1)) + 1     'Do the translation
      Exit Function
    Else
      Index = Index - 1
      Y = (Index - (Index Mod 10)) / 10 + 1              'Get row number of 'board' type board (10 by 10)
      X = (Index Mod 10) + 1                             'Get column number of 'board' type board (10 by 10)
      If Y < 2 Or Y > 9 Or X < 2 Or X > 9 Then IndexTranslation = InvalidSquare: Exit Function  'Check if it will map to normal 8 by 8 board
      Index = (Index - 11) - (2 * (Y - 2))     'Do the translation
    End If
  End If
  
  If Reversed = True Then                             'If the board is reserved the return the inverse index
    IndexTranslation = 63 - Index
  Else
    IndexTranslation = Index
  End If
  
End Function

Public Function SaveSettings()
  SetKeyValue HKEY_LOCAL_MACHINE, RegistryKey & App.Title, "Move Speed", MoveSpeed
  SetKeyValue HKEY_LOCAL_MACHINE, RegistryKey & App.Title, "GraphNum", ComputerGraphic
  SetKeyValue HKEY_LOCAL_MACHINE, RegistryKey & App.Title, "Gametype Mode", PlayType
  SetKeyValue HKEY_LOCAL_MACHINE, RegistryKey & App.Title, "Time Limit", TimeLimit
  SetKeyValue HKEY_LOCAL_MACHINE, RegistryKey & App.Title, "Cheat", CheatSwitch
  SetKeyValue HKEY_LOCAL_MACHINE, RegistryKey & App.Title, "Name 1", , Names(1)
  SetKeyValue HKEY_LOCAL_MACHINE, RegistryKey & App.Title, "Name 2", , Names(2)
  SetKeyValue HKEY_LOCAL_MACHINE, RegistryKey & App.Title, "Advanced", IsAdvanced
  SetKeyValue HKEY_LOCAL_MACHINE, RegistryKey & App.Title, "Autoswitch", AutoSwitch
End Function

Public Function GetSettings()
  GetKeyValue HKEY_LOCAL_MACHINE, RegistryKey & App.Title, "Move speed", MoveSpeed
  GetKeyValue HKEY_LOCAL_MACHINE, RegistryKey & App.Title, "GraphNum", ComputerGraphic
  GetKeyValue HKEY_LOCAL_MACHINE, RegistryKey & App.Title, "Gametype Mode", PlayType
  GetKeyValue HKEY_LOCAL_MACHINE, RegistryKey & App.Title, "Time Limit", TimeLimit
  GetKeyValue HKEY_LOCAL_MACHINE, RegistryKey & App.Title, "Cheat", CheatSwitch
  GetKeyValue HKEY_LOCAL_MACHINE, RegistryKey & App.Title, "Name 1", , Names(1)
  GetKeyValue HKEY_LOCAL_MACHINE, RegistryKey & App.Title, "Name 2", , Names(2)
  GetKeyValue HKEY_LOCAL_MACHINE, RegistryKey & App.Title, "Advanced", IsAdvanced
  GetKeyValue HKEY_LOCAL_MACHINE, RegistryKey & App.Title, "Autoswitch", AutoSwitch
End Function

Private Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, Optional ByRef KeyVal As Long, Optional ByRef KeyValStr As String) As Boolean
    Dim i As Long                                           ' Loop Counter
    Dim RC As Long                                          ' Return Code
    Dim hKey As Long                                        ' Handle To An Open Registry Key
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' Data Type Of A Registry Key
    Dim tmpvaral As String                                    ' Tempory Storage For A Registry Key Value
    Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
    '------------------------------------------------------------
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    RC = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
    
    If (RC <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...
    
    tmpvaral = String$(1024, 0)                             ' Allocate Variable Space
    KeyValSize = 1024                                       ' Mark Variable Size
    
    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
    RC = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpvaral, KeyValSize)    ' Get/Create Key Value
                        
    If (RC <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
    
    If (Asc(Mid(tmpvaral, KeyValSize, 1)) = 0) Then           ' Win95 Adds Null Terminated String...
        tmpvaral = Left(tmpvaral, KeyValSize - 1)               ' Null Found, Extract From String
    Else                                                    ' WinNT Does NOT Null Terminate String...
        tmpvaral = Left(tmpvaral, KeyValSize)                   ' Null Not Found, Extract String Only
    End If
    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Search Data Types...
    Case REG_SZ                                             ' String Registry Key Data Type
        KeyVal = Val(tmpvaral)                                ' Copy String Value
        KeyValStr = tmpvaral
    Case REG_DWORD                                          ' Double Word Registry Key Data Type
        For i = Len(tmpvaral) To 1 Step -1                    ' Convert Each Bit
            KeyVal = KeyVal + Hex(Asc(Mid(tmpvaral, i, 1)))   ' Build Value Char. By Char.
        Next
        KeyVal = Val(Format$("&h" + KeyVal))                ' Convert Double Word To String
        KeyValStr = CStr(Format$("&h" + KeyVal))
    End Select
    
    GetKeyValue = True                                      ' Return Success
    RC = RegCloseKey(hKey)                                  ' Close Registry Key
    Exit Function                                           ' Exit
    
GetKeyError:      ' Cleanup After An Error Has Occured...
    KeyVal = 0                                             ' Set Return Val To Empty String
    KeyValStr = ""
    GetKeyValue = False                                     ' Return Failure
    RC = RegCloseKey(hKey)                                  ' Close Registry Key
End Function

Private Function SetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, Optional ByRef KeyVal As Long, Optional ByRef KeyValStr As String) As Boolean
    Dim i As Long                                           ' Loop Counter
    Dim RC As Long                                          ' Return Code
    Dim hKey As Long                                        ' Handle To An Open Registry Key
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' Data Type Of A Registry Key
    Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
    '------------------------------------------------------------
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    RC = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
    
    If RC <> ERROR_SUCCESS Then
      If RC = ERROR_FILE_NOT_FOUND Then
        RC = RegCreateKey(KeyRoot, KeyName, hKey)
      Else
        GoTo GetKeyError          ' Handle Error...
      End If
    End If
    
    If KeyValStr = "" Then KeyValStr = CStr(KeyVal)
    KeyValSize = Len(KeyValStr) + 1
    KeyValType = REG_SZ
    
    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
    RC = RegSetValueEx(hKey, SubKeyRef, 0, KeyValType, ByVal KeyValStr, KeyValSize)
    
    If (RC <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
    
    SetKeyValue = True                                      ' Return Success
    RC = RegCloseKey(hKey)                                  ' Close Registry Key
    Exit Function                                           ' Exit
    
GetKeyError:      ' Cleanup After An Error Has Occured...
    SetKeyValue = False                                     ' Return Failure
    RC = RegCloseKey(hKey)                                  ' Close Registry Key
End Function


