Attribute VB_Name = "mdlSuDoku"
Option Explicit
Option Base 1

    Public Cntr As Integer ' General Purpose Counter
    
    Enum PuzzleStatus
        [Solving Puzzle]
        [Puzzle Solved]
        [Puzzle Stalled]
        [Invalid Solution]
    End Enum
    Public SolveStatus As PuzzleStatus
    
    Public Type CP ' CellPosition
        GridX As Byte
        GridY As Byte
        BlockX As Byte
        BlockY As Byte
    End Type
    Public POS As CP
    Public ALT As CP
    
    Public Type CellData
        IDX As Integer ' Index of the Text Box
        VAL As Byte    ' Value of the Cell
        CAN As String  ' CANdidate values for the Cell
        BMK As Integer ' BitMasK of canditate values for the Cell
    End Type
    Public CD(3, 3, 3, 3) As CellData
    
    Public Type CellColours
        BackColour As OLE_COLOR
        TextColour As OLE_COLOR
    End Type
    Public Type AppSettings
        SEED As CellColours
        SOLVED As CellColours
    End Type
    Public SuDoku As AppSettings


'================================='
' Get settings from the Registry. '
'================================='
Public Sub GetValues()

    SuDoku.SEED.BackColour = GetSetting("SuDoku Solver", "Seed Cell", "Back Colour", vbWhite)
    SuDoku.SEED.TextColour = GetSetting("SuDoku Solver", "Seed Cell", "Text Colour", vbRed)
    
    SuDoku.SOLVED.BackColour = GetSetting("SuDoku Solver", "Solved Cell", "Back Colour", vbWhite)
    SuDoku.SOLVED.TextColour = GetSetting("SuDoku Solver", "Solved Cell", "Text Colour", vbBlue)

End Sub

'==================================='
' Write settings into the Registry. '
'==================================='
Public Sub PutValues()

    SaveSetting "SuDoku Solver", "Seed Cell", "Back Colour", SuDoku.SEED.BackColour
    SaveSetting "SuDoku Solver", "Seed Cell", "Text Colour", SuDoku.SEED.TextColour
    
    SaveSetting "SuDoku Solver", "Solved Cell", "Back Colour", SuDoku.SOLVED.BackColour
    SaveSetting "SuDoku Solver", "Solved Cell", "Text Colour", SuDoku.SOLVED.TextColour

End Sub

'====================================================='
' Return a Cell's Grid position when given it's Index '
'====================================================='
Public Function Get_POS(Index As Integer) As CP

    Get_POS.GridY = CByte(1 + Int((Index - 1) / 27))
    Get_POS.GridX = CByte(1 + Int((Index - 1) / 9) - 3 * (Get_POS.GridY - 1))
    Get_POS.BlockY = CByte(1 + Int((Index - 1) / 3) - 9 * (Get_POS.GridY - 1) - 3 * (Get_POS.GridX - 1))
    Get_POS.BlockX = CByte(Index - 27 * (Get_POS.GridY - 1) - 9 * (Get_POS.GridX - 1) - 3 * (Get_POS.BlockY - 1))

End Function


'==================================================='
' Returns the Base 2 logarithm of the given integer '
'==================================================='
Public Function LOG2(ByVal Num As Integer) As Double

    LOG2 = Log(Num) / Log(2#)

End Function

'============================================================'
' Returns TRUE or FALSE for whether a number is whole or not '
'============================================================'
Public Function WholeNumber(Num As Double) As Boolean

    WholeNumber = IIf(Fix(Num) = Num, True, False)

End Function

'========================================'
' Returns a string of Bits representing  '
' then Binary value of the Num parameter '
'========================================'
Public Function StrBitMask(Num As Integer) As String

    Dim Pntr As Integer
    
    StrBitMask = ""
    For Pntr = 1 To 9
        If (Num And (2 ^ (Pntr - 1))) > 0 Then
            StrBitMask = "1" + StrBitMask
        Else
            StrBitMask = "0" + StrBitMask
        End If
    Next Pntr

End Function

'========================================='
' Returns a integer containing the number '
' of Bits set to "1" in the Num parameter '
'========================================='
Public Function BitCount(Num As Integer) As Integer

    Dim Pntr As Integer
    
    BitCount = 0
    For Pntr = 1 To 9
        If (Num And (2 ^ (Pntr - 1))) = (2 ^ (Pntr - 1)) Then
            BitCount = BitCount + 1
        End If
    Next Pntr

End Function

'========================================================'
' Returns an integer containing the Value of the         '
' BitNumth bit, from lowest to highest, that is set to 1 '
'========================================================'
Public Function GetSetBitVAL(Num As Integer, _
                             Optional BitNum As Integer = 1) As Integer

    Dim Pntr As Integer
    Dim BitCount As Integer
    
    For Pntr = 1 To 9
        If (Num And (2 ^ (Pntr - 1))) = (2 ^ (Pntr - 1)) Then
            BitCount = BitCount + 1
        End If
        If BitCount = BitNum Then
            GetSetBitVAL = Pntr
            Exit For
        End If
    Next Pntr

End Function
