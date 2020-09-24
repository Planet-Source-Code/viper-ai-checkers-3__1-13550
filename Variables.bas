Attribute VB_Name = "Variables"
'-------------------------------------------------------------------------------------------'
'    This Program and the code contained within is Copyright of Infostrategy Ltd. 2000
'    Any attempt to copy this program or any portion of it without first consulting the
'    Company will be met with deadly force and/or a lawsuit (particularly if copied with
'    Comercial gain in mind) -Viper
'-------------------------------------------------------------------------------------------'

Option Explicit

'--------------------------API Declarations------------------------
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpvaralueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpvaralueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function GetTickCount Lib "kernel32" () As Long 'used to time things

'--------------------------Constants-------------------------------


'The Fields are composed as follows     : (doublemask) (ForbiddenBit) (p2mask) (p1mask) (four Bits pointer into PieceLists) - zero if free
'The Piecelists are composed as follows : (doublemask) (seven Bits position on board) - zero if captured
'Bytes are used because the Board is copied into TempBoard very frequently and should therefore be as short as possible
'The board itself is stretched to one dimension because a lone index is faster than two indexes

Public Const DoubleMask     As Byte = 128
Public Const InvalidSquare  As Byte = 64
Public Const P1Mask         As Byte = 32
Public Const P2Mask         As Byte = 16
Public Const PlayerMask     As Byte = 48
Public Const IndexMask      As Byte = 127
Public Const PieceNumMask   As Byte = 15
Public Const DepthLimit     As Long = 48    'max ply searchable

'---Registry Constants---
Public Const RegistryKey As String = "Software\Infostrategy\"
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const ERROR_SUCCESS = 0
Public Const REG_SZ = 1                         ' Unicode nul terminated string
Public Const REG_DWORD = 4
Public Const ERROR_FILE_NOT_FOUND = 2&
Public Const READ_CONTROL = &H20000
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_CREATE_LINK = &H20
Public Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL

'--------------------------Self Defined Types-----------------------

Public Const Player1 As Long = 1
Public Const Player2 As Long = 2
Public Const BothSides As Long = 3

Public Const MoveCompleted As Long = 16
Public Const MoveCorrect As Long = 32

Public Const SingleValue As Long = 10
Public Const Doublevalue As Long = 30
Public Const Maxmaterial As Long = 12 * Doublevalue
Public Const Infinity As Long = 10000
Public Const ForwardJumpBonus As Long = SingleValue / 10

Public Const ERR_Misalignment As Long = 1
Public Const NumComImages As Long = 6


Public Type Board
  Score                    As Integer       'current board material balance
  Pieces(1 To 2, 1 To 12)  As Byte          'lists of black and white pieces
  Fields(1 To 100)         As Byte          '8 by 8 plus guard fields (ForbiddenBit) on all four sides
  MovesListFrom            As Long          'pointers into MoveList
  MovesListTo              As Long
  Turn                     As Byte
End Type

Public Type History
  Board As Board
End Type

'The Fields are composed as follows     : (doublemask) (ForbiddenBit) (p2mask) (p1mask) (four bits pointer into piecelist) - zero if free
'The Piecelists are composed as follows : (doublemask) (seven bits Position on board) - zero if captured
'Bytes are used because the board is copied into TempBoard very frequently and should therefore be as short as possible

'-----------Public Variables---------------

Public VP1Time As Long, VP2Time As Long, VTurns As Long 'Timing variables
Public GameStarted As Boolean
Public IndexMoves(1 To 4) As Long           'Index for moves (Direction)
Public Score(1 To 2) As Long                'Stores the amount of pieces each player has
Public AutoDebug As Boolean                 'Stores if the program is in auto debug mode
Public BoardHistory() As Board              'Stores board of moves made in the past
Public Names(1 To 2) As String              'Stores player names in string array
Public PlayType As Long                     'Stores play type (1=human vs computer 2=human vs human)
Public AutoSwitch As Long                   'Stores if board automatically switches orientation (only in 2 player mode)
Public CheatSwitch As Long                  'Stores whether cheat switch is toggled (1 = on 0 = off)
Public MoveSpeed As Long                    'Stores in milliseconds time to keep piece selected before moving (computer only)
Public IsAdvanced As Long                   'Stores if advanced panel is open (-1 = open 0 = closed)
Public Reversed As Boolean                  'Stores (session length) if the board orientation is reversed
Public ShowHelp As Boolean                  'Stores whether help msg boxes are shown
Public ComputerGraphic As Long              'Stores imagelist number for the computer picture
Public MoveListChanged As Boolean           'Stores whether current movelist is correct (if it needs changing)

Public Terminate            As Boolean       'True when unloading requested
Public TimeLimit            As Long
Public Cutoffs              As Long
Public PosnsVisited         As Long
Public Forced               As Boolean       'True when forced move encountered
Public MoveList(1 To DepthLimit * 156) As String 'Min size which will definitely not overflow
Public MoveListIx           As Long          'Index into MoveList
Public ResetIx              As Long          'Index is reset to this value if a longer MoveChain is found
Public MaxChainLength       As Long
Public PV(0 To DepthLimit, 0 To DepthLimit) As String        'Principal Variation
Public InPV                 As Long          '1 if in PV and 0 else
Public CurrentBoard         As Board         'The Board
