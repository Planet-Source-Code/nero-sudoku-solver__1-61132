VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSOLVER 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Su Doku Solver"
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5100
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   463
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   340
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdEXIT 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4050
      TabIndex        =   4
      Top             =   5790
      Width           =   750
   End
   Begin MSComctlLib.ProgressBar pgbSOLVED 
      Height          =   255
      Left            =   240
      Negotiate       =   -1  'True
      TabIndex        =   5
      Top             =   6480
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Max             =   81
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdCLEAR 
      Caption         =   "Clear"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1140
      TabIndex        =   3
      Top             =   5790
      Width           =   750
   End
   Begin VB.CommandButton cmdSOLVE 
      Caption         =   "Solve"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   5790
      Width           =   750
   End
   Begin VB.TextBox txtCELL 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   480
      Index           =   1
      Left            =   240
      MaxLength       =   1
      TabIndex        =   1
      ToolTipText     =   "111111111"
      Top             =   990
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   "Enter the seed numbers into the cells, then click on SOLVE to get the answer. Click CLEAR to restart for a new puzzle."
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "frmSOLVER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

    Dim CrtSolved As Byte
    Dim NewSolved As Byte

    Dim SP As String ' String to hold a Saved Puzzle

Private Sub cmdEXIT_Click()

    Unload Me

End Sub

Private Sub LoadPUZZLE(PzlStr As String)

    Call cmdCLEAR_Click
    For Cntr = 1 To 81
        txtCELL(Cntr).Text = Mid(PzlStr, Cntr, 1)
    Next Cntr

End Sub

Private Sub Form_Initialize()

    Dim Xleft As Integer, Ytop As Integer
    
    For Cntr = 1 To 81
        POS = Get_POS(Cntr)
        If Cntr > 1 Then
            Load txtCELL(Cntr)
            Xleft = txtCELL(1).Left + txtCELL(1).Width * (POS.BlockX - 1) _
                  + (3 * txtCELL(1).Width + 8) * (POS.GridX - 1)
            Ytop = txtCELL(1).Top + txtCELL(1).Height * (POS.BlockY - 1) _
                 + (3 * txtCELL(1).Height + 8) * (POS.GridY - 1)
            txtCELL(Cntr).Move Xleft, Ytop
            txtCELL(Cntr).Visible = True
            txtCELL(Cntr).TabIndex = Cntr
        End If
        CD(POS.GridX, POS.GridY, POS.BlockX, POS.BlockY).IDX = Cntr
    Next Cntr
    
    cmdSOLVE.TabIndex = txtCELL(81).TabIndex + 1
    cmdCLEAR.TabIndex = cmdSOLVE.TabIndex + 1
    pgbSOLVED.TabIndex = cmdCLEAR.TabIndex + 1
    
    Call GetValues
    Call cmdCLEAR_Click

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Const PS1 As String = "200006300456080000008100006008403009109000305400209600600005100000010562005300004"
    Const PS2 As String = "400080706002104080830002500100270030007500940050019006008300042090806700705090003"
    Const PS3 As String = "000906002009040510000015000605070800000000000008030706000740000067050400800902000"
    
    Select Case Button
        Case 1: Call LoadPUZZLE(PS2)
        Case 2: Call LoadPUZZLE(PS3)
    End Select

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call PutValues

End Sub

'====================================================='
' Highlight all text within a Cell when it gets Focus '
'====================================================='
Private Sub txtCELL_GotFocus(Index As Integer)

    If Len(txtCELL(Index).Text) > 0 Then
        txtCELL(Index).SelStart = 0
        txtCELL(Index).SelLength = Len(txtCELL(Index).Text)
    End If

End Sub

'========================================================'
' Enable the use of Arrow Keys to navigate between Cells '
'========================================================'
Private Sub txtCELL_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    POS = Get_POS(Index)
    Select Case KeyCode
        Case vbKeyLeft
            If POS.GridX > 1 Or POS.BlockX > 1 Then
                If POS.BlockX = 1 Then
                    txtCELL(CD(POS.GridX - 1, POS.GridY, 3, POS.BlockY).IDX).SetFocus
                Else
                    txtCELL(CD(POS.GridX, POS.GridY, POS.BlockX - 1, POS.BlockY).IDX).SetFocus
                End If
            Else
                txtCELL(CD(3, POS.GridY, 3, POS.BlockY).IDX).SetFocus
            End If
        Case vbKeyUp
            If POS.GridY > 1 Or POS.BlockY > 1 Then
                If POS.BlockY = 1 Then
                    txtCELL(CD(POS.GridX, POS.GridY - 1, POS.BlockX, 3).IDX).SetFocus
                Else
                    txtCELL(CD(POS.GridX, POS.GridY, POS.BlockX, POS.BlockY - 1).IDX).SetFocus
                End If
            Else
                txtCELL(CD(POS.GridX, 3, POS.BlockX, 3).IDX).SetFocus
            End If
        Case vbKeyRight
            If POS.GridX < 3 Or POS.BlockX < 3 Then
                If POS.BlockX = 3 Then
                    txtCELL(CD(POS.GridX + 1, POS.GridY, 1, POS.BlockY).IDX).SetFocus
                Else
                    txtCELL(CD(POS.GridX, POS.GridY, POS.BlockX + 1, POS.BlockY).IDX).SetFocus
                End If
            Else
                txtCELL(CD(1, POS.GridY, 1, POS.BlockY).IDX).SetFocus
            End If
        Case vbKeyDown
            If POS.GridY < 3 Or POS.BlockY < 3 Then
                If POS.BlockY = 3 Then
                    txtCELL(CD(POS.GridX, POS.GridY + 1, POS.BlockX, 1).IDX).SetFocus
                Else
                    txtCELL(CD(POS.GridX, POS.GridY, POS.BlockX, POS.BlockY + 1).IDX).SetFocus
                End If
            Else
                txtCELL(CD(POS.GridX, 1, POS.BlockX, 1).IDX).SetFocus
            End If
    End Select
    
End Sub

'============================'
' Enter a number into a Cell '
'============================'
Private Sub txtCELL_Change(Index As Integer)

    POS = Get_POS(Index)
    
    If Not IsNumeric(txtCELL(Index).Text) _
    Or txtCELL(Index).Text = "0" Then Call Reset_Cell(Index)
    
    If IsNumeric(txtCELL(Index).Text) Then
        If Number_Exists(Index, CByte(txtCELL(Index).Text)) Then
            MsgBox "Number already exixts in this" + vbCrLf + _
                   "3x3 Cell, Row or Column, please" + vbCrLf + _
                   "enter a different selection.", _
                   vbExclamation Or vbOKOnly, "SuDoku Solver"
            Call Reset_Cell(Index)
        Else
            Call SetCellVal(Index, CInt(txtCELL(Index).Text))
            If Index < 81 Then
                If cmdSOLVE.Enabled = True Then txtCELL(Index + 1).SetFocus
            Else
                cmdSOLVE.SetFocus
            End If
        End If
    End If

End Sub

'======================================'
' Check that entered numbers are valid '
'======================================'
Private Function Number_Exists(Index As Integer, Nmbr As Byte) As Boolean

    Dim NEX As CP
    
    Number_Exists = False
    NEX = Get_POS(Index)
    
    'Check that the entered number doesn't exist in the same 3X3 CELL
    For ALT.BlockY = 1 To 3: For ALT.BlockX = 1 To 3
        If CD(NEX.GridX, NEX.GridY, ALT.BlockX, ALT.BlockY).VAL = Nmbr _
        And (NEX.BlockX <> ALT.BlockX Or NEX.BlockY <> ALT.BlockY) Then
            Number_Exists = True: Exit Function
        End If
    Next ALT.BlockX: Next ALT.BlockY
    'Check that the entered number doesn't exist in the same ROW
    For ALT.GridX = 1 To 3: For ALT.BlockX = 1 To 3
        If CD(ALT.GridX, NEX.GridY, ALT.BlockX, NEX.BlockY).VAL = Nmbr _
        And NEX.GridX <> ALT.GridX Then
            Number_Exists = True: Exit Function
        End If
    Next ALT.BlockX: Next ALT.GridX
    'Check that the entered number doesn't exist in the same COLUMN
    For ALT.GridY = 1 To 3: For ALT.BlockY = 1 To 3
        If CD(NEX.GridX, ALT.GridY, NEX.BlockX, ALT.BlockY).VAL = Nmbr _
        And NEX.GridY <> ALT.GridY Then
            Number_Exists = True: Exit Function
        End If
    Next ALT.BlockY: Next ALT.GridY

End Function

'============================='
' Attempt to Solve the Puzzle '
'============================='
Private Sub cmdSOLVE_Click()

    Dim SolvedSquares As Byte, PrevSolved As Byte
    
    cmdCLEAR.Enabled = True
    cmdSOLVE.Enabled = False
    
    For Cntr = 1 To 81
        POS = Get_POS(Cntr)
        If CD(POS.GridX, POS.GridY, POS.BlockX, POS.BlockY).VAL > 0 Then
            SolvedSquares = SolvedSquares + 1
        Else
            txtCELL(Cntr).ForeColor = vbBlack
        End If
    Next Cntr
    pgbSOLVED.Min = SolvedSquares
    
    'Save the current puzzle before starting the SOLVE
    SP = String(81, "0")
    For Cntr = 1 To 81
        If txtCELL(Cntr).Text > "" Then
            Mid(SP, Cntr, 1) = txtCELL(Cntr).Text
        End If
    Next Cntr
    
    CrtSolved = SolvedCount()
    Do
        Do
            Call PrimarySOLVE
        Loop Until SolveStatus > [Solving Puzzle]
        
        If SolveStatus = [Puzzle Stalled] Then
            Call AdvancedSOLVE_1
            Call AdvancedSOLVE_2
            For Cntr = 2 To 4
                Call AdvancedSOLVE_3(Cntr)
                Call AdvancedSOLVE_4(Cntr)
            Next Cntr
        End If
    
    Loop Until SolveStatus = [Invalid Solution] Or SolveStatus = [Puzzle Solved]
    
    Call Paint_Solved(SuDoku.SOLVED.BackColour, SuDoku.SOLVED.TextColour)
    
    If SolveStatus = [Invalid Solution] Then
        Select Case MsgBox("Unable to solve this puzzle. Would you" + vbCrLf + _
                           "like to change the puzzle and try again?", _
                            vbInformation Or vbYesNoCancel, "SuDoku Solver")
            Case vbYes
                Call LoadPUZZLE(SP)
            Case vbNo
                Call cmdCLEAR_Click
        End Select
    End If

End Sub

Private Sub Paint_Solved(BackColour As OLE_COLOR, TextColour As OLE_COLOR)

    For Cntr = 1 To 81
        POS = Get_POS(Cntr)
        'When a Cell is empty, but it's value is greater than zero,
        'then it has been solved, but not painted, so we paint it here
        If CD(POS.GridX, POS.GridY, POS.BlockX, POS.BlockY).VAL > 0 _
        And txtCELL(Cntr).Text = "" Then
            txtCELL(Cntr).BackColor = BackColour
            txtCELL(Cntr).ForeColor = TextColour
            txtCELL(Cntr).Text = Trim(Str(CD(POS.GridX, POS.GridY, POS.BlockX, POS.BlockY).VAL))
        End If
    Next Cntr

End Sub

Private Sub PrimarySOLVE()

    For Cntr = 1 To 81
        POS = Get_POS(Cntr)
        If CD(POS.GridX, POS.GridY, POS.BlockX, POS.BlockY).VAL = 0 Then
            'Find and remove all occurances of numbers in the same 3X3 cell
            For ALT.BlockY = 1 To 3
            For ALT.BlockX = 1 To 3
                If CD(POS.GridX, POS.GridY, ALT.BlockX, ALT.BlockY).VAL > 0 And (POS.BlockX <> ALT.BlockX Or POS.BlockY <> ALT.BlockY) Then
                    Call Exclude_Candidate(Cntr, CInt(CD(POS.GridX, POS.GridY, ALT.BlockX, ALT.BlockY).VAL))
                End If
            Next ALT.BlockX
            Next ALT.BlockY
            'Find and remove all occurances of numbers in the same row
            For ALT.GridX = 1 To 3
            For ALT.BlockX = 1 To 3
                If CD(ALT.GridX, POS.GridY, ALT.BlockX, POS.BlockY).VAL > 0 And POS.GridX <> ALT.GridX Then
                    Call Exclude_Candidate(Cntr, CInt(CD(ALT.GridX, POS.GridY, ALT.BlockX, POS.BlockY).VAL))
                End If
            Next ALT.BlockX
            Next ALT.GridX
            'Find and remove all occurances of numbers in the same column
            For ALT.GridY = 1 To 3
            For ALT.BlockY = 1 To 3
                If CD(POS.GridX, ALT.GridY, POS.BlockX, ALT.BlockY).VAL > 0 And POS.GridY <> ALT.GridY Then
                    Call Exclude_Candidate(Cntr, CInt(CD(POS.GridX, ALT.GridY, POS.BlockX, ALT.BlockY).VAL))
                End If
            Next ALT.BlockY
            Next ALT.GridY
        End If
    Next Cntr
    
    For Cntr = 1 To 81
        POS = Get_POS(Cntr)
        Select Case BitCount(CD(POS.GridX, POS.GridY, POS.BlockX, POS.BlockY).BMK)
            Case 0
                'When no Candidate Values are left and the Cell Value
                'is still Zero, then the puzzle must be malformed
                If CD(POS.GridX, POS.GridY, POS.BlockX, POS.BlockY).VAL = 0 Then
                    SolveStatus = [Invalid Solution]
                End If
            Case 1
                'When only 1 Candidate number is left then a Cell is Solved
                Call SetCellVal(Cntr, 1 + LOG2(CD(POS.GridX, POS.GridY, POS.BlockX, POS.BlockY).BMK))
                'CD(POS.GridX, POS.GridY, POS.BlockX, POS.BlockY).VAL = _
                    1 + LOG2(CD(POS.GridX, POS.GridY, POS.BlockX, POS.BlockY).BMK)
                '(POS.GridX, POS.GridY, POS.BlockX, POS.BlockY).BMK = 0
        End Select
    Next Cntr
    
    Call CheckSolveStatus

End Sub

Private Sub CheckSolveStatus()

    If SolveStatus = [Puzzle Solved] _
    Or SolveStatus = [Invalid Solution] Then Exit Sub
    
    NewSolved = SolvedCount()
    If NewSolved > CrtSolved Then
        CrtSolved = NewSolved
        If CrtSolved < 81 Then
            SolveStatus = [Solving Puzzle]
        Else
            SolveStatus = [Puzzle Solved]
        End If
    Else
        If SolveStatus = [Puzzle Stalled] Then
            SolveStatus = [Invalid Solution]
        Else
            SolveStatus = [Puzzle Stalled]
        End If
    End If

End Sub

'=================================='
' Find Potential Values that Occur '
' only ONCE in any ROW or COLUMN   '
'=================================='
Private Sub AdvancedSOLVE_1()

    Dim CrtSolved As Byte, NewSolved As Byte
    
    CrtSolved = SolvedCount()
    
    Dim Cntx As Integer
    Dim BitMask As Integer
    
    For Cntx = 1 To 81
        POS = Get_POS(Cntx)
        If CD(POS.GridX, POS.GridY, POS.BlockX, POS.BlockY).VAL = 0 Then
            BitMask = CD(POS.GridX, POS.GridY, POS.BlockX, POS.BlockY).BMK
            'AND NOT the Bitmask with the BitMasks of all
            'the other unsolved Cells in the same ROW
            For ALT.GridX = 1 To 3: For ALT.BlockX = 1 To 3
                If (POS.BlockX <> ALT.BlockX Or POS.GridX <> ALT.GridX) _
                And CD(ALT.GridX, POS.GridY, ALT.BlockX, POS.BlockY).VAL = 0 Then
                    BitMask = BitMask And Not CD(ALT.GridX, POS.GridY, ALT.BlockX, POS.BlockY).BMK
                End If
            Next ALT.BlockX: Next ALT.GridX
            If BitMask > 0 Then
                If WholeNumber(LOG2(BitMask)) Then
                    'Debug.Print "ROW" + Str(POS.BlockY + 3 * (POS.GridY - 1)) + _
                                " has an exclusive" + Str(CInt(1 + LOG2(BitMask))) + _
                                " in COLUMN" + Str(POS.BlockX + 3 * (POS.GridX - 1))
                    Call SetCellVal(Cntx, CInt(1 + LOG2(BitMask)))
                End If
            End If
        End If
        If CD(POS.GridX, POS.GridY, POS.BlockX, POS.BlockY).VAL = 0 Then
            BitMask = CD(POS.GridX, POS.GridY, POS.BlockX, POS.BlockY).BMK
            'AND NOT the Bitmask with the BitMasks of all
            'the other unsolved Cells in the same COLUMN
            For ALT.GridY = 1 To 3: For ALT.BlockY = 1 To 3
                If (POS.BlockY <> ALT.BlockY Or POS.GridY <> ALT.GridY) _
                And CD(POS.GridX, ALT.GridY, POS.BlockX, ALT.BlockY).VAL = 0 Then
                    BitMask = BitMask And Not CD(POS.GridX, ALT.GridY, POS.BlockX, ALT.BlockY).BMK
                End If
            Next ALT.BlockY: Next ALT.GridY
            If BitMask > 0 Then
                If WholeNumber(LOG2(BitMask)) Then
                    'Debug.Print "COLUMN" + Str(POS.BlockX + 3 * (POS.GridX - 1)) + _
                                " has an exclusive" + Str(CInt(1 + LOG2(BitMask))) + _
                                " in ROW" + Str(POS.BlockY + 3 * (POS.GridY - 1))
                    Call SetCellVal(Cntx, CInt(1 + LOG2(BitMask)))
                End If
            End If
        End If
    Next Cntx

End Sub

'===================================================='
' Find Potential Values that Occur only within ONE   '
' BLOCK within any ROW or COLUMN then eliminate this '
' value from the other CELLS within the same BLOCK   '
'===================================================='
Private Sub AdvancedSOLVE_2()

    Dim BitMask As Integer
    Dim AS2_GridX As Byte
    Dim AS2_GridY As Byte
    Dim AS2_BlockX As Byte
    Dim AS2_BlockY As Byte
    
    For POS.GridY = 1 To 3
    For POS.BlockY = 1 To 3
    For POS.GridX = 1 To 3
        BitMask = -1
        For POS.BlockX = 1 To 3
            If CD(POS.GridX, POS.GridY, POS.BlockX, POS.BlockY).BMK > 0 Then
                If BitMask < 0 Then
                    BitMask = CD(POS.GridX, POS.GridY, POS.BlockX, POS.BlockY).BMK
                Else
                    BitMask = BitMask And CD(POS.GridX, POS.GridY, POS.BlockX, POS.BlockY).BMK
                End If
            End If
        Next POS.BlockX
        For AS2_GridX = 1 To 3
        For AS2_BlockX = 1 To 3
            If AS2_GridX <> POS.GridX Then
                BitMask = BitMask And Not CD(AS2_GridX, POS.GridY, AS2_BlockX, POS.BlockY).BMK
            End If
        Next AS2_BlockX
        Next AS2_GridX
        If BitMask > 0 Then
            If WholeNumber(LOG2(BitMask)) Then
                For AS2_BlockY = 1 To 3
                For AS2_BlockX = 1 To 3
                    If AS2_BlockY <> POS.BlockY Then
                        Call Exclude_Candidate(CD(POS.GridX, POS.GridY, AS2_BlockX, AS2_BlockY).IDX, CInt(1 + LOG2(BitMask)))
                    End If
                Next AS2_BlockX
                Next AS2_BlockY
                'Debug.Print "Row" + Str(3 * (POS.GridY - 1) + POS.BlockY) + " Grid" + Str(POS.GridX), 1 + LOG2(BitMask), StrBitMask(BitMask)
            End If
        End If
    Next POS.GridX
    Next POS.BlockY
    Next POS.GridY
    
    For POS.GridX = 1 To 3
    For POS.BlockX = 1 To 3
    For POS.GridY = 1 To 3
        'If POS.GridX = 1 And POS.BlockX = 3 And POS.GridY = 2 Then Stop
        BitMask = -1
        For POS.BlockY = 1 To 3
            If CD(POS.GridX, POS.GridY, POS.BlockX, POS.BlockY).BMK > 0 Then
                If BitMask < 0 Then
                    BitMask = CD(POS.GridX, POS.GridY, POS.BlockX, POS.BlockY).BMK
                Else
                    BitMask = BitMask And CD(POS.GridX, POS.GridY, POS.BlockX, POS.BlockY).BMK
                End If
            End If
        Next POS.BlockY
        For AS2_GridY = 1 To 3
        For AS2_BlockY = 1 To 3
            If AS2_GridY <> POS.GridY Then
                BitMask = BitMask And Not CD(POS.GridX, AS2_GridY, POS.BlockX, AS2_BlockY).BMK
            End If
        Next AS2_BlockY
        Next AS2_GridY
        If BitMask > 0 Then
            If WholeNumber(LOG2(BitMask)) Then
                For AS2_BlockY = 1 To 3
                For AS2_BlockX = 1 To 3
                    If AS2_BlockX <> POS.BlockX Then
                        Call Exclude_Candidate(CD(POS.GridX, POS.GridY, AS2_BlockX, AS2_BlockY).IDX, CInt(1 + LOG2(BitMask)))
                    End If
                Next AS2_BlockX
                Next AS2_BlockY
                'Debug.Print "Col" + Str(3 * (POS.GridX - 1) + POS.BlockX) + " Grid" + Str(POS.GridY), 1 + LOG2(BitMask), StrBitMask(BitMask)
            End If
        End If
    Next POS.GridY
    Next POS.BlockX
    Next POS.GridX

End Sub

'==========================================================='
' Find candidate value groups (2,3 or 4) that occur as many '
' times in any ROW or COLUMN, then eliminate their values   '
' from any other candidate values in the same ROW or COLUMN '
'==========================================================='
Private Sub AdvancedSOLVE_3(GroupSize As Integer)

    Dim BitMask As Integer
    Dim FoundCnt As Integer
    Dim Cnt3 As Integer
    Dim AS3_GridX As Byte
    Dim AS3_GridY As Byte
    Dim AS3_BlockX As Byte
    Dim AS3_BlockY As Byte
    
    'Solve by ROWs
    For POS.GridY = 1 To 3
    For POS.BlockY = 1 To 3
        BitMask = -1
        For POS.GridX = 1 To 3
        For POS.BlockX = 1 To 3
            If BitCount(CD(POS.GridX, POS.GridY, POS.BlockX, POS.BlockY).BMK) = GroupSize Then
                BitMask = CD(POS.GridX, POS.GridY, POS.BlockX, POS.BlockY).BMK
                FoundCnt = 1
                For AS3_GridX = 1 To 3
                For AS3_BlockX = 1 To 3
                    If (AS3_GridX <> POS.GridX Or AS3_BlockX <> POS.BlockX) _
                    And BitCount(CD(AS3_GridX, POS.GridY, AS3_BlockX, POS.BlockY).BMK) >= 2 _
                    And BitCount(CD(AS3_GridX, POS.GridY, AS3_BlockX, POS.BlockY).BMK) <= GroupSize _
                    And (BitMask And CD(AS3_GridX, POS.GridY, AS3_BlockX, POS.BlockY).BMK) = BitMask Then
                        FoundCnt = FoundCnt + 1
                    End If
                Next AS3_BlockX
                Next AS3_GridX
                If FoundCnt = GroupSize Then
                    For AS3_GridX = 1 To 3
                    For AS3_BlockX = 1 To 3
                        If (AS3_GridX <> POS.GridX Or AS3_BlockX <> POS.BlockX) _
                        And BitCount(CD(AS3_GridX, POS.GridY, AS3_BlockX, POS.BlockY).BMK) >= 1 _
                        And (BitMask And CD(AS3_GridX, POS.GridY, AS3_BlockX, POS.BlockY).BMK) <> BitMask Then
                            For Cnt3 = 1 To GroupSize
                                Call Exclude_Candidate(CD(AS3_GridX, POS.GridY, AS3_BlockX, POS.BlockY).IDX, _
                                                       GetSetBitVAL(BitMask, Cnt3))
                            Next Cnt3
                        End If
                    Next AS3_BlockX
                    Next AS3_GridX
                End If
            End If
        Next POS.BlockX
        Next POS.GridX
    Next POS.BlockY
    Next POS.GridY
    
    'Solve by COLUMNs
    For POS.GridX = 1 To 3
    For POS.BlockX = 1 To 3
        BitMask = -1
        For POS.GridY = 1 To 3
        For POS.BlockY = 1 To 3
            If BitCount(CD(POS.GridX, POS.GridY, POS.BlockX, POS.BlockY).BMK) = GroupSize Then
                BitMask = CD(POS.GridX, POS.GridY, POS.BlockX, POS.BlockY).BMK
                FoundCnt = 1
                For AS3_GridY = 1 To 3
                For AS3_BlockY = 1 To 3
                    If (AS3_GridY <> POS.GridY Or AS3_BlockY <> POS.BlockY) _
                    And BitCount(CD(POS.GridX, AS3_GridY, POS.BlockX, AS3_BlockY).BMK) >= 2 _
                    And BitCount(CD(POS.GridX, AS3_GridY, POS.BlockX, AS3_BlockY).BMK) <= GroupSize _
                    And (BitMask And CD(POS.GridX, AS3_GridY, POS.BlockX, AS3_BlockY).BMK) = BitMask Then
                        FoundCnt = FoundCnt + 1
                    End If
                Next AS3_BlockY
                Next AS3_GridY
                If FoundCnt = GroupSize Then
                    For AS3_GridY = 1 To 3
                    For AS3_BlockY = 1 To 3
                        If (AS3_GridY <> POS.GridY Or AS3_BlockY <> POS.BlockY) _
                        And BitCount(CD(POS.GridX, AS3_GridY, POS.BlockX, AS3_BlockY).BMK) >= 1 _
                        And (BitMask And CD(POS.GridX, AS3_GridY, POS.BlockX, AS3_BlockY).BMK) <> BitMask Then
                            For Cnt3 = 1 To GroupSize
                                Call Exclude_Candidate(CD(POS.GridX, AS3_GridY, POS.BlockX, AS3_BlockY).IDX, _
                                                       GetSetBitVAL(BitMask, Cnt3))
                            Next Cnt3
                        End If
                    Next AS3_BlockY
                    Next AS3_GridY
                End If
            End If
        Next POS.BlockY
        Next POS.GridY
    Next POS.BlockX
    Next POS.GridX

End Sub

'=========================================================='
' When a ROW, COLUMN or BOX contains a set of N candidate  '
' groups contain all occurrences of a set of N candidate   '
' numbers, other numbers can be removed from those groups. '
'=========================================================='
Private Sub AdvancedSOLVE_4(GroupSize As Integer)
'this is for extreme puzzles and hasn't been coded yet
End Sub

Private Sub SetCellVal(Index As Integer, Nmbr As Integer)

    Dim SCV As CP
    SCV = Get_POS(Index)
    
    CD(SCV.GridX, SCV.GridY, SCV.BlockX, SCV.BlockY).VAL = Nmbr
    CD(SCV.GridX, SCV.GridY, SCV.BlockX, SCV.BlockY).BMK = 0
    
    'Find and remove all occurances of the Candidate number in the same 3X3 BLOCK
    For ALT.BlockY = 1 To 3: For ALT.BlockX = 1 To 3
        If (SCV.BlockX <> ALT.BlockX Or SCV.BlockY <> ALT.BlockY) _
        And CD(SCV.GridX, SCV.GridY, ALT.BlockX, ALT.BlockY).BMK And Nmbr = Nmbr Then
            Call Exclude_Candidate(CD(SCV.GridX, SCV.GridY, ALT.BlockX, ALT.BlockY).IDX, Nmbr)
        End If
    Next ALT.BlockX: Next ALT.BlockY
    'Find and remove all occurances of the Candidate number in the same ROW
    For ALT.GridX = 1 To 3: For ALT.BlockX = 1 To 3
        If POS.GridX <> ALT.GridX _
        And CD(ALT.GridX, POS.GridY, ALT.BlockX, POS.BlockY).BMK And Nmbr = Nmbr Then
            Call Exclude_Candidate(CD(ALT.GridX, POS.GridY, ALT.BlockX, POS.BlockY).IDX, Nmbr)
        End If
    Next ALT.BlockX: Next ALT.GridX
    'Find and remove all occurances of the Candidate number in the same COLUMN
    For ALT.GridY = 1 To 3: For ALT.BlockY = 1 To 3
        If POS.GridY <> ALT.GridY _
        And CD(POS.GridX, ALT.GridY, POS.BlockX, ALT.BlockY).BMK And Nmbr = Nmbr Then
            Call Exclude_Candidate(CD(POS.GridX, ALT.GridY, POS.BlockX, ALT.BlockY).IDX, Nmbr)
        End If
    Next ALT.BlockY: Next ALT.GridY

End Sub

'======================================='
' Exclude a Candidate Value from a Cell '
'======================================='
Private Sub Exclude_Candidate(Index As Integer, Nmbr As Integer)

    Dim XCL As CP
    XCL = Get_POS(Index)
    
    If CD(XCL.GridX, XCL.GridY, XCL.BlockX, XCL.BlockY).BMK > 0 Then
        CD(XCL.GridX, XCL.GridY, XCL.BlockX, XCL.BlockY).BMK = _
            CD(XCL.GridX, XCL.GridY, XCL.BlockX, XCL.BlockY).BMK And Not (2 ^ (Nmbr - 1))
    End If

End Sub

'========================================'
' Clear ALL Cell's Attributes and Values '
'========================================'
Private Sub cmdCLEAR_Click()

    For Cntr = 1 To 81: Call Reset_Cell(Cntr): Next Cntr
    cmdSOLVE.Enabled = True
    cmdCLEAR.Enabled = False
    pgbSOLVED.Min = 0: pgbSOLVED.Value = 0
    SolveStatus = [Solving Puzzle]

End Sub

'============================================='
' Reset a Single Cell's Attributes and Values '
'============================================='
Private Sub Reset_Cell(Index As Integer)

    POS = Get_POS(Index)
    CD(POS.GridX, POS.GridY, POS.BlockX, POS.BlockY).VAL = 0
    'CD(POS.GridX, POS.GridY, POS.BlockX, POS.BlockY).CAN = "123456789"
    CD(POS.GridX, POS.GridY, POS.BlockX, POS.BlockY).BMK = &O777 'Octal 777 is Binary 111111111
    txtCELL(Index).Text = ""
    txtCELL(Index).BackColor = SuDoku.SEED.BackColour
    txtCELL(Index).ForeColor = SuDoku.SEED.TextColour

End Sub

'===================================='
' Returns the number of SOLVED Cells '
'===================================='
Private Function SolvedCount() As Byte
    
    SolvedCount = 0
    For Cntr = 1 To 81
        POS = Get_POS(Cntr)
        If CD(POS.GridX, POS.GridY, POS.BlockX, POS.BlockY).VAL > 0 Then
            SolvedCount = SolvedCount + 1
        End If
    Next Cntr
    
    pgbSOLVED.Value = SolvedCount

End Function
