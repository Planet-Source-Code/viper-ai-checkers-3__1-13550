Attribute VB_Name = "AI"
'-------------------------------------------------------------------------------------------'
'    This Program and the code contained within is Copyright of Infostrategy Ltd. 2000
'    Any attempt to copy this program or any portion of it without first consulting the
'    Company will be met with deadly force :P -Viper
'-------------------------------------------------------------------------------------------'

Option Explicit

Public Sub AIMove()
Dim StopTime As Long, IterDepth As Long, Result As Long, MoveChain As String
Dim FromField As Long, ToField As Long, Captured As Long, Temp As Long

StopTime = Timer + TimeLimit

PosnsVisited = 0
Cutoffs = 0
Forced = False

'---Cosmetic changes-----
With frmMain
  .MousePointer = vbHourglass
  .ComPicBox.Picture = .ImageList2.ListImages(ComputerGraphic).Picture
  .lblComStatus = "Thinking..."
End With

DoEvents

Do
  CurrentBoard.MovesListFrom = 1
  CurrentBoard.MovesListTo = 0
  IterDepth = IterDepth + 1
  Result = Search(CurrentBoard, 0, IterDepth, -Infinity, Infinity)
  InPV = 1
Loop While Timer < StopTime And IterDepth < DepthLimit And Not Forced And Abs(Result) <= Maxmaterial

'----More Cosmetic changes----
With frmMain
  .MousePointer = vbNormal
  .lblCutoffs = Cutoffs
  .lblMoves = PosnsVisited
  .ComPicBox.Picture = Nothing
  .lblComStatus = ""
End With

VTurns = VTurns + 1

Select Case Result
Case -Infinity
  frmMain.lblComStatus = "I win!!"
Case Is < -Maxmaterial
  MsgBox "I resign", vbExclamation
  ResetGame
Case Else
  If Result > Maxmaterial Then
    Temp = Infinity - Result
    frmMain.lblComStatus = "I win in " & Temp & " move" & IIf(Temp > 1, "s", "") & "!!"
  End If
  MoveChain = PV(0, 0)
  frmMain.Shape1(IndexTranslation(Asc(MoveChain), , False)).Picture = frmMain.ImageList1.ListImages(5).Picture
  MakeMove CurrentBoard, MoveChain, True
  DoEvents
  Sleep MoveSpeed
  
  Do
    FromField = IndexTranslation(Asc(MoveChain), , False)
    ToField = IndexTranslation(Asc(Mid$(MoveChain, 2)), , False)
    If Len(MoveChain) >= 3 Then Captured = IndexTranslation(Asc(Mid$(MoveChain, 3)), , False)
    frmMain.Shape1(FromField).Picture = Nothing
    If Len(MoveChain) > 3 Then
      frmMain.Shape1(ToField).Picture = frmMain.ImageList1.ListImages(5).Picture
    Else
      frmMain.Shape1(ToField).Picture = frmMain.ImageList1.ListImages(PieceImageNum(CurrentBoard.Fields(Asc(Mid$(MoveChain, 2))))).Picture
      If Captured And Len(PV(0, 0)) < 4 Then frmMain.Shape1(Captured).Picture = Nothing
    End If
  MoveChain = Mid$(MoveChain, 4)
  Sleep MoveSpeed
  Loop Until Len(MoveChain) < 3
  
'  MoveChain = PV(0, 0)
'  If Len(MoveChain) > 3 Then
'    Do
'      frmMain.Shape1(IndexTranslation(Asc(Mid$(MoveChain, 3)), , False)).Picture = frmMain.ImageList1.ListImages(5).Picture
'      Sleep 100
'      MoveChain = Mid$(MoveChain, 4)
'    Loop Until Len(MoveChain) < 3
'  End If

  MoveChain = PV(0, 0)
  If Len(MoveChain) > 3 Then
    Do
      frmMain.Shape1(IndexTranslation(Asc(Mid$(MoveChain, 3)), , False)).Picture = Nothing
      Sleep 180
      MoveChain = Mid$(MoveChain, 4)
    Loop Until Len(MoveChain) < 3
  End If
End Select

CheckWin

If CurrentBoard.Turn = 2 Then
  VP2Time = VP2Time + (Timer - (StopTime - TimeLimit))
Else
  VP1Time = VP1Time + (Timer - (StopTime - TimeLimit))
End If

CurrentBoard.Turn = BothSides - CurrentBoard.Turn
MoveListChanged = True

ReDim Preserve BoardHistory(1 To UBound(BoardHistory) + 1)
BoardHistory(UBound(BoardHistory)) = CurrentBoard
frmMain.cmdBack.Enabled = True

End Sub

Private Function Search(Board As Board, Depth As Long, MinDepth As Long, Alpha As Long, Beta As Long) As Long
Dim TempBoard As Board, Field As Long, BestResult As Long, CurrChain As String
Dim Lng1 As Long, Lng2 As Long, Result As Long

PosnsVisited = PosnsVisited + 1 'Record number of moves searched
If InPV Then 'If Principal Variation has something in...
  If Len(PV(0, Depth)) Then 'Another check if PV has something in
    MoveList(Board.MovesListFrom) = PV(0, Depth)
  Else
    InPV = 0 'Correct false information (cause PV has nothin init)
  End If
End If

MaxChainLength = 1
MoveListIx = Board.MovesListFrom - 1 + InPV
ResetIx = MoveListIx

For Lng1 = 1 To 12 'For every piece
  Field = Board.Pieces(Board.Turn, Lng1) And IndexMask 'Translate to an index of the board
  If Field Then GenerateMoves Board, Field, 0, "", 0
Next Lng1

PV(Depth, Depth) = ""
Board.MovesListTo = MoveListIx

If (MaxChainLength = 2 And Depth >= MinDepth) Or (Depth = DepthLimit) Then
  BestResult = Evaluate(Board)
Else
  BestResult = -Infinity + Depth
  For Lng1 = Board.MovesListFrom To Board.MovesListTo
    If BestResult >= Beta Then 'move is bad enough we don't want to know if there are any worse ones
      Cutoffs = Cutoffs + 1
      Exit For
    Else
      CurrChain = MoveList(Lng1)
      If BestResult > Alpha Then Alpha = BestResult     '-Alpha becomes beta in next depth
      TempBoard = Board
      MakeMove TempBoard, CurrChain, True
      If Board.MovesListTo = 1 Then
        Result = Evaluate(Board)
        Forced = True
      Else
        TempBoard.MovesListFrom = Board.MovesListTo + 1         'Establish new segment in move list
        TempBoard.Turn = BothSides - TempBoard.Turn             'Toggle side to move
        Result = -Search(TempBoard, Depth + 1, MinDepth, -Beta, -Alpha)       'Minimax recursion
        TempBoard.Turn = BothSides - TempBoard.Turn             'Back to this side to move
      End If
      If Result > BestResult Then
        BestResult = Result
        For Lng2 = Depth + 1 To DepthLimit - 1 'Create principal variation
          PV(Depth, Lng2) = PV(Depth + 1, Lng2) 'Copy all best moves after this 1
          If PV(Depth, Lng2) = "" Then Exit For
        Next Lng2
        PV(Depth, Depth) = CurrChain 'Enter this best move into principal variation
      End If
    End If
  Next Lng1
End If

Search = BestResult

End Function

Public Sub GenerateMoves(Board As Board, Field As Long, Depth As Long, MovesSoFar As String, ForbiddenDirection As Long)
Dim TempBoard As Board, Direction As Long, Adjacent As Long, Beyond As Long
Dim EnemyBit As Long, CurrMoveChain As String, Slide As Long

EnemyBit = Board.Turn * 16 '(p1mask has turn =1 this makes enemybit p2mask which is 16)

If Board.Fields(Field) And DoubleMask Then
'---------Double piece section----
For Direction = 1 To 4
  If Direction <> ForbiddenDirection Then
    Adjacent = IndexMoves(Direction) 'Record numberindex of a particular direction
    Beyond = Field + Adjacent 'set beyond to 1 square beyond current piece location
    Do Until Board.Fields(Beyond) And InvalidSquare 'do until the beyond square cannot be moved onto
      Slide = Beyond
      Beyond = Beyond + Adjacent
      If (Board.Fields(Slide) And EnemyBit) And Board.Fields(Beyond) = 0 Then 'if there is a piece on slide square and an empty space behind it
        TempBoard = Board
        CurrMoveChain = Chr$(Field) & Chr$(Beyond) & Chr$(Slide)
        MakeMove TempBoard, CurrMoveChain, False
        GenerateMoves TempBoard, Beyond, Depth + 1, MovesSoFar & CurrMoveChain, 5 - Direction 'this move is now branched off further
        Exit Do
      ElseIf Board.Fields(Slide) = 0 Then
        If Depth = 0 Then RecordMove Board, Chr$(Field) & Chr$(Slide) 'record all moves between field and the invalid/occupied square
      Else
        Exit Do
      End If
    Loop
  End If
Next

Else 'presumably piece is a single bit

For Direction = 1 To 4
  If Direction <> ForbiddenDirection Then
    Adjacent = Field + IndexMoves(Direction)
    Beyond = Adjacent + IndexMoves(Direction)
    If Board.Fields(Adjacent) And EnemyBit Then
      If Board.Fields(Beyond) = 0 Then
        CurrMoveChain = Chr$(Field) & Chr$(Beyond) & Chr$(Adjacent)
        TempBoard = Board
        MakeMove TempBoard, CurrMoveChain, False
        GenerateMoves TempBoard, Beyond, Depth + 1, MovesSoFar & CurrMoveChain, 5 - Direction
      End If
    Else
      If Board.Fields(Adjacent) = 0 And ((Adjacent < Field And Board.Turn = 2) Or (Adjacent > Field And Board.Turn = 1)) And Depth = 0 Then
        RecordMove Board, Chr$(Field) & Chr$(Adjacent)
      End If
    End If
  End If
Next

End If

If MovesSoFar <> "" Then RecordMove Board, MovesSoFar

End Sub

Private Sub RecordMove(Board As Board, MoveChain As String)
Dim K As Long

K = Len(MoveChain)
If K > MaxChainLength Then 'Change the requirement for chainlength (since rules say you must take as many pieces as possible)
  MaxChainLength = K
  MoveListIx = ResetIx
End If

If K = MaxChainLength Then 'If move meets length requirement then record it
  If MoveChain <> MoveList(Board.MovesListFrom) Or InPV = 0 Then
    MoveListIx = MoveListIx + 1
    MoveList(MoveListIx) = MoveChain
  End If
End If

End Sub

Private Function Evaluate(Board As Board) As Long
  If Board.Turn = 1 Then
    Evaluate = Board.Score
  Else
    Evaluate = -Board.Score
  End If
End Function
