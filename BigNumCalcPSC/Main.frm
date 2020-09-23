VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7230
   ClientLeft      =   150
   ClientTop       =   150
   ClientWidth     =   7020
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   482
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   468
   Begin VB.CommandButton cmdLogicCalc 
      Caption         =   "Combinations"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   13
      Left            =   4740
      TabIndex        =   100
      TabStop         =   0   'False
      ToolTipText     =   " N!/(r!(N-r)!) -> r 1st from N 2nd num "
      Top             =   4425
      Width           =   1365
   End
   Begin VB.CommandButton cmdLogicCalc 
      Caption         =   "Permutations"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   12
      Left            =   4755
      TabIndex        =   99
      TabStop         =   0   'False
      ToolTipText     =   " N!/(N-r)! -> r 1st from N 2nd num "
      Top             =   4140
      Width           =   1335
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Random 2"
      Height          =   270
      Index           =   1
      Left            =   5145
      TabIndex        =   98
      TabStop         =   0   'False
      Top             =   2475
      Width           =   1260
   End
   Begin VB.Frame fraVBASM 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5160
      TabIndex        =   95
      Top             =   1260
      Width           =   1650
      Begin VB.OptionButton optVBASM 
         Caption         =   "ASM"
         Height          =   225
         Index           =   1
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   97
         TabStop         =   0   'False
         Top             =   45
         Width           =   795
      End
      Begin VB.OptionButton optVBASM 
         Caption         =   "VB"
         Height          =   225
         Index           =   0
         Left            =   75
         Style           =   1  'Graphical
         TabIndex        =   96
         TabStop         =   0   'False
         Top             =   45
         Width           =   540
      End
   End
   Begin VB.CommandButton cmdLogicCalc 
      Caption         =   "N3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   11
      Left            =   6120
      Picture         =   "Main.frx":0E42
      TabIndex        =   90
      TabStop         =   0   'False
      ToolTipText     =   "  Cube of 2nd number "
      Top             =   4425
      Width           =   435
   End
   Begin VB.CommandButton cmdLogicCalc 
      Caption         =   "N2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   10
      Left            =   6120
      Picture         =   "Main.frx":13CC
      TabIndex        =   89
      TabStop         =   0   'False
      ToolTipText     =   " Square of 2nd number "
      Top             =   4140
      Width           =   435
   End
   Begin VB.CommandButton cmdLogicCalc 
      Caption         =   "N!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   9
      Left            =   6120
      Picture         =   "Main.frx":1956
      TabIndex        =   87
      TabStop         =   0   'False
      ToolTipText     =   " Factorial of 2nd number <= 4 decimal digits "
      Top             =   3855
      Width           =   435
   End
   Begin VB.CommandButton cmdLogicCalc 
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   8
      Left            =   5535
      TabIndex        =   80
      TabStop         =   0   'False
      ToolTipText     =   " Abs Subtraction "
      Top             =   3855
      Width           =   555
   End
   Begin VB.CommandButton cmdLogicCalc 
      Caption         =   "| Sub |"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   7
      Left            =   4800
      TabIndex        =   79
      TabStop         =   0   'False
      ToolTipText     =   " Abs Subtraction "
      Top             =   3855
      Width           =   705
   End
   Begin VB.CommandButton cmdLogicCalc 
      Caption         =   "Mul"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   6
      Left            =   4275
      TabIndex        =   78
      TabStop         =   0   'False
      Top             =   3855
      Width           =   495
   End
   Begin VB.CommandButton cmdLogicCalc 
      Caption         =   "Div/Mod"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   5
      Left            =   3270
      TabIndex        =   77
      TabStop         =   0   'False
      Top             =   3855
      Width           =   975
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "?"
      Height          =   285
      Left            =   6555
      TabIndex        =   65
      TabStop         =   0   'False
      ToolTipText     =   " Help "
      Top             =   945
      Width           =   270
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Random 1 && 2"
      Height          =   300
      Index           =   0
      Left            =   5130
      TabIndex        =   64
      TabStop         =   0   'False
      Top             =   930
      Width           =   1365
   End
   Begin VB.CommandButton cmdSR 
      Caption         =   "Not"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   4
      Left            =   2925
      TabIndex        =   60
      TabStop         =   0   'False
      ToolTipText     =   " Invert bits "
      Top             =   1275
      Width           =   510
   End
   Begin VB.CommandButton cmdLogicCalc 
      Caption         =   "Imp"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   4
      Left            =   2595
      TabIndex        =   59
      TabStop         =   0   'False
      Top             =   3855
      Width           =   495
   End
   Begin VB.CommandButton cmdLogicCalc 
      Caption         =   "Eqv"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   3
      Left            =   2070
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   3855
      Width           =   495
   End
   Begin VB.CommandButton cmdSR 
      Caption         =   "roR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   3
      Left            =   2295
      TabIndex        =   56
      TabStop         =   0   'False
      ToolTipText     =   " Roll right "
      Top             =   1275
      Width           =   510
   End
   Begin VB.CommandButton cmdSR 
      Caption         =   "shR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   2
      Left            =   1725
      TabIndex        =   55
      TabStop         =   0   'False
      ToolTipText     =   " Shift right "
      Top             =   1275
      Width           =   510
   End
   Begin VB.CommandButton cmdSR 
      Caption         =   "roL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   1095
      TabIndex        =   54
      TabStop         =   0   'False
      ToolTipText     =   " Roll left "
      Top             =   1275
      Width           =   510
   End
   Begin VB.CommandButton cmdSR 
      Caption         =   "shL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   510
      TabIndex        =   53
      TabStop         =   0   'False
      ToolTipText     =   " Shift left "
      Top             =   1275
      Width           =   510
   End
   Begin VB.CommandButton cmdClear2 
      Caption         =   "Clr 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3420
      TabIndex        =   52
      TabStop         =   0   'False
      ToolTipText     =   " Clear 2nd numbers & results "
      Top             =   2490
      Width           =   675
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   5835
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   15
      Width           =   600
   End
   Begin VB.Frame fraHexLen 
      Caption         =   "Max hex input length"
      Height          =   555
      Left            =   5115
      TabIndex        =   32
      ToolTipText     =   " Max hex input "
      Top             =   330
      Width           =   1725
      Begin VB.HScrollBar HSHexLen 
         Height          =   255
         Left            =   930
         Max             =   39
         Min             =   1
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   210
         Value           =   16
         Width           =   525
      End
      Begin VB.Label LabHexLen 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LabHexLen"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   195
         TabIndex        =   57
         Top             =   225
         Width           =   660
      End
   End
   Begin VB.CommandButton cmdClear1 
      Caption         =   "Clr 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   3435
      TabIndex        =   22
      TabStop         =   0   'False
      ToolTipText     =   " Clear 1st number & enable hex sizing "
      Top             =   15
      Width           =   615
   End
   Begin VB.CommandButton cmdConvCalc 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Calc"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   810
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   " Toggle calculator "
      Top             =   15
      Width           =   750
   End
   Begin VB.CommandButton cmdConvCalc 
      Caption         =   "Conv"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   75
      TabIndex        =   18
      TabStop         =   0   'False
      ToolTipText     =   " Converter only "
      Top             =   30
      Width           =   690
   End
   Begin VB.CommandButton cmdDisplay 
      Caption         =   "Disp && ClipB"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4335
      TabIndex        =   12
      TabStop         =   0   'False
      ToolTipText     =   " Full Display, Clipboard, Save "
      Top             =   15
      Width           =   1320
   End
   Begin VB.CommandButton delLR 
      Caption         =   "delR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   2415
      TabIndex        =   11
      TabStop         =   0   'False
      ToolTipText     =   " Delete from right "
      Top             =   15
      Width           =   750
   End
   Begin VB.CommandButton delLR 
      Caption         =   "delL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   1650
      TabIndex        =   10
      TabStop         =   0   'False
      ToolTipText     =   " Delete from left "
      Top             =   15
      Width           =   720
   End
   Begin VB.Frame fraInput 
      Caption         =   "Input or Kyb"
      Height          =   1245
      Left            =   3810
      TabIndex        =   9
      Top             =   330
      Width           =   1245
      Begin VB.CommandButton cmdPad 
         BackColor       =   &H00C0FFFF&
         Caption         =   "F"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   15
         Left            =   870
         Style           =   1  'Graphical
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   915
         Width           =   255
      End
      Begin VB.CommandButton cmdPad 
         BackColor       =   &H00C0FFFF&
         Caption         =   "E"
         Height          =   225
         Index           =   14
         Left            =   615
         Style           =   1  'Graphical
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   915
         Width           =   255
      End
      Begin VB.CommandButton cmdPad 
         BackColor       =   &H00C0FFFF&
         Caption         =   "D"
         Height          =   225
         Index           =   13
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   915
         Width           =   255
      End
      Begin VB.CommandButton cmdPad 
         BackColor       =   &H00C0FFFF&
         Caption         =   "C"
         Height          =   225
         Index           =   12
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   915
         Width           =   255
      End
      Begin VB.CommandButton cmdPad 
         BackColor       =   &H00C0FFFF&
         Caption         =   "B"
         Height          =   225
         Index           =   11
         Left            =   870
         Style           =   1  'Graphical
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   705
         Width           =   255
      End
      Begin VB.CommandButton cmdPad 
         BackColor       =   &H00C0FFFF&
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   10
         Left            =   615
         Style           =   1  'Graphical
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   690
         Width           =   255
      End
      Begin VB.CommandButton cmdPad 
         BackColor       =   &H00C0FFFF&
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   9
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   690
         Width           =   255
      End
      Begin VB.CommandButton cmdPad 
         BackColor       =   &H00C0FFFF&
         Caption         =   "8"
         Height          =   225
         Index           =   8
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   690
         Width           =   255
      End
      Begin VB.CommandButton cmdPad 
         BackColor       =   &H00C0FFFF&
         Caption         =   "7"
         Height          =   225
         Index           =   7
         Left            =   870
         Style           =   1  'Graphical
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   465
         Width           =   255
      End
      Begin VB.CommandButton cmdPad 
         BackColor       =   &H00C0FFFF&
         Caption         =   "6"
         Height          =   225
         Index           =   6
         Left            =   630
         Style           =   1  'Graphical
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   465
         Width           =   255
      End
      Begin VB.CommandButton cmdPad 
         BackColor       =   &H00C0FFFF&
         Caption         =   "5"
         Height          =   225
         Index           =   5
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   465
         Width           =   255
      End
      Begin VB.CommandButton cmdPad 
         BackColor       =   &H00C0FFFF&
         Caption         =   "4"
         Height          =   225
         Index           =   4
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   465
         Width           =   255
      End
      Begin VB.CommandButton cmdPad 
         BackColor       =   &H00C0FFFF&
         Caption         =   "3"
         Height          =   225
         Index           =   3
         Left            =   870
         Style           =   1  'Graphical
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton cmdPad 
         BackColor       =   &H00C0FFFF&
         Caption         =   "2"
         Height          =   225
         Index           =   2
         Left            =   615
         Style           =   1  'Graphical
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton cmdPad 
         BackColor       =   &H00C0FFFF&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton cmdPad 
         BackColor       =   &H00C0FFFF&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   240
         Width           =   255
      End
   End
   Begin VB.CommandButton cmdLogicCalc 
      Caption         =   "Xor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   2
      Left            =   1545
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3855
      Width           =   495
   End
   Begin VB.CommandButton cmdLogicCalc 
      Caption         =   "Or"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   1
      Left            =   1080
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3855
      Width           =   435
   End
   Begin VB.CommandButton cmdLogicCalc 
      Caption         =   "And"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   540
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3855
      Width           =   510
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   3
      Left            =   480
      TabIndex        =   3
      Text            =   "Text1(3)"
      Top             =   2685
      Width           =   1860
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   4
      Left            =   480
      TabIndex        =   4
      Text            =   "Text1(4)"
      Top             =   3045
      Width           =   2580
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   5
      Left            =   480
      TabIndex        =   5
      Text            =   "Text1(5)"
      Top             =   3405
      Width           =   5610
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   480
      TabIndex        =   2
      Text            =   "Text1(2)"
      Top             =   1590
      Width           =   5610
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   480
      TabIndex        =   1
      Text            =   "Text1(1)"
      Top             =   885
      Width           =   2580
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Text            =   "Text1(0)"
      Top             =   525
      Width           =   1860
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   2310
      Left            =   480
      ScaleHeight     =   2310
      ScaleWidth      =   6255
      TabIndex        =   66
      Top             =   4380
      Width           =   6255
      Begin VB.TextBox txtResult 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   15
         TabIndex        =   72
         Text            =   "txtResult(5)"
         Top             =   2010
         Width           =   5610
      End
      Begin VB.TextBox txtResult 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   4
         Left            =   15
         TabIndex        =   71
         Text            =   "txtResult(4)"
         Top             =   1665
         Width           =   2580
      End
      Begin VB.TextBox txtResult 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   15
         TabIndex        =   70
         Text            =   "txtResult(3)"
         Top             =   1305
         Width           =   1860
      End
      Begin VB.TextBox txtResult 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   15
         TabIndex        =   69
         Text            =   "txtResult(2)"
         Top             =   705
         Width           =   5610
      End
      Begin VB.TextBox txtResult 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   15
         TabIndex        =   68
         Text            =   "txtResult(1)"
         Top             =   375
         Width           =   2580
      End
      Begin VB.TextBox txtResult 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   30
         TabIndex        =   67
         Text            =   "txtResult(0)"
         Top             =   60
         Width           =   1860
      End
      Begin VB.Label LabMsg 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Caption         =   "LabMsg(1)"
         ForeColor       =   &H000000FF&
         Height          =   240
         Index           =   1
         Left            =   3315
         TabIndex        =   101
         Top             =   375
         Width           =   2760
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "L 8"
         Height          =   255
         Index           =   8
         Left            =   5655
         TabIndex        =   83
         Top             =   720
         Width           =   630
      End
      Begin VB.Label LabMarks 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   2
         Left            =   30
         TabIndex        =   94
         Top             =   960
         Width           =   5700
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "L 11"
         Height          =   255
         Index           =   11
         Left            =   5685
         TabIndex        =   86
         Top             =   2025
         Width           =   480
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "L 10"
         Height          =   255
         Index           =   10
         Left            =   2655
         TabIndex        =   85
         Top             =   1695
         Width           =   480
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "L 9"
         Height          =   255
         Index           =   9
         Left            =   1950
         TabIndex        =   84
         Top             =   1335
         Width           =   480
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "L 7"
         Height          =   255
         Index           =   7
         Left            =   2640
         TabIndex        =   82
         Top             =   390
         Width           =   570
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "L 6"
         Height          =   255
         Index           =   6
         Left            =   1935
         TabIndex        =   81
         Top             =   75
         Width           =   570
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Remainder"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   11
         Left            =   90
         TabIndex        =   76
         Top             =   1095
         Width           =   930
      End
   End
   Begin VB.Label LabMarks 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Index           =   3
      Left            =   510
      TabIndex        =   93
      Top             =   6645
      Width           =   5700
   End
   Begin VB.Label LabMarks 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Index           =   1
      Left            =   525
      TabIndex        =   92
      Top             =   3690
      Width           =   5670
   End
   Begin VB.Label LabMarks 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Index           =   0
      Left            =   510
      TabIndex        =   91
      Top             =   1875
      Width           =   5670
   End
   Begin VB.Label LabMsg 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "LabMsg(0)"
      ForeColor       =   &H000000FF&
      Height          =   570
      Index           =   0
      Left            =   5475
      TabIndex        =   88
      Top             =   2775
      Width           =   1245
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      FillColor       =   &H00FFFFFF&
      Height          =   4620
      Index           =   1
      Left            =   210
      Shape           =   4  'Rounded Rectangle
      Top             =   2340
      Width           =   6630
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "B"
      Height          =   240
      Index           =   10
      Left            =   300
      TabIndex        =   75
      Top             =   6405
      Width           =   105
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "D"
      Height          =   240
      Index           =   9
      Left            =   300
      TabIndex        =   74
      Top             =   6060
      Width           =   135
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "H"
      Height          =   240
      Index           =   8
      Left            =   300
      TabIndex        =   73
      Top             =   5685
      Width           =   135
   End
   Begin VB.Label LabOP 
      BackStyle       =   0  'Transparent
      Caption         =   "LabOP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1185
      TabIndex        =   63
      Top             =   4185
      Width           =   1155
   End
   Begin VB.Label Label1 
      Caption         =   "lo"
      Height          =   180
      Index           =   1
      Left            =   2130
      TabIndex        =   62
      Top             =   330
      Width           =   165
   End
   Begin VB.Label Label1 
      Caption         =   "hi"
      Height          =   180
      Index           =   0
      Left            =   495
      TabIndex        =   61
      Top             =   330
      Width           =   240
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Result"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   7
      Left            =   570
      TabIndex        =   21
      Top             =   4185
      Width           =   825
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Second number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   615
      TabIndex        =   20
      Top             =   2445
      Width           =   1350
   End
   Begin VB.Label Lab1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "L 0"
      Height          =   255
      Index           =   0
      Left            =   2415
      TabIndex        =   34
      Top             =   555
      Width           =   480
   End
   Begin VB.Label Label7 
      Caption         =   "B"
      Height          =   240
      Index           =   8
      Left            =   225
      TabIndex        =   31
      Top             =   1590
      Width           =   135
   End
   Begin VB.Label Label7 
      Caption         =   "D"
      Height          =   240
      Index           =   7
      Left            =   225
      TabIndex        =   30
      Top             =   915
      Width           =   135
   End
   Begin VB.Label Label7 
      Caption         =   "H"
      Height          =   240
      Index           =   6
      Left            =   225
      TabIndex        =   29
      Top             =   540
      Width           =   135
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "B"
      Height          =   240
      Index           =   5
      Left            =   300
      TabIndex        =   28
      Top             =   5100
      Width           =   105
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "D"
      Height          =   240
      Index           =   4
      Left            =   300
      TabIndex        =   27
      Top             =   4755
      Width           =   135
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "H"
      Height          =   240
      Index           =   3
      Left            =   300
      TabIndex        =   26
      Top             =   4410
      Width           =   135
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "B"
      Height          =   240
      Index           =   2
      Left            =   300
      TabIndex        =   25
      Top             =   3465
      Width           =   135
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "D"
      Height          =   240
      Index           =   1
      Left            =   300
      TabIndex        =   24
      Top             =   3090
      Width           =   135
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "H"
      Height          =   240
      Index           =   0
      Left            =   300
      TabIndex        =   23
      Top             =   2730
      Width           =   135
   End
   Begin VB.Label Lab1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "L 5"
      Height          =   255
      Index           =   5
      Left            =   6195
      TabIndex        =   17
      Top             =   3450
      Width           =   480
   End
   Begin VB.Label Lab1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "L 4"
      Height          =   255
      Index           =   4
      Left            =   3180
      TabIndex        =   16
      Top             =   3075
      Width           =   480
   End
   Begin VB.Label Lab1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "L 3"
      Height          =   255
      Index           =   3
      Left            =   2460
      TabIndex        =   15
      Top             =   2715
      Width           =   480
   End
   Begin VB.Label Lab1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "L2"
      Height          =   255
      Index           =   2
      Left            =   6150
      TabIndex        =   14
      Top             =   1620
      Width           =   480
   End
   Begin VB.Label Lab1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "L 1"
      Height          =   255
      Index           =   1
      Left            =   3135
      TabIndex        =   13
      Top             =   915
      Width           =   480
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFC0C0&
      BorderColor     =   &H00808080&
      BorderWidth     =   4
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   4725
      Index           =   0
      Left            =   150
      Shape           =   4  'Rounded Rectangle
      Top             =   2280
      Width           =   6750
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Main.frm


' Update  1 July 2007

'1. Corrected FulIntDivision


' BigNum Calculator  by Robert Rayment April 2006

' Permutations & Combinations added 6 Apr

' Multiply routines base on code by
' EKabiljo  PSC CodeId=24960 - Not available anymore!

Option Explicit

Option Base 1

Private Declare Sub InitCommonControls Lib "comctl32" ()
Private Declare Function LoadLibrary Lib "Kernel32" Alias "LoadLibraryA" ( _
    ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "Kernel32" ( _
   ByVal hLibModule As Long) As Long
Private m_hMod As Long


Private aT1Focus() As Boolean
Private aBlock As Boolean

Private MaxHexInput As Long
Private MaxDecInput As Long
Private MaxBinInput As Long

Private IP As Long   ' Textbox Insertion Point

Private Const NumberHex$ = "0123456789ABCDEF"
Private Const NumberDec$ = "0123456789"
Private Const NumberBin$ = "01"

'For routines with Multiplier
Private aa() As Long
Private bb() As Long
Private cc() As Long
Private LengthA As Long
Private LengthB As Long
Private LengthC As Long

Const MaxLong = 10000
Private TLong(MaxLong) As Long
Private LenTLong As Long

Private Sub Form_Initialize()
   m_hMod = LoadLibrary("shell32.dll")
   InitCommonControls
End Sub

Private Sub Form_Load()
   STX = Screen.TwipsPerPixelX
   STY = Screen.TwipsPerPixelY
   
   PathSpec$ = App.Path
   If Right$(PathSpec$, 1) <> "\" Then PathSpec$ = PathSpec$ & "\"
   CurrPath$ = PathSpec$

   Top = 1000
   Left = Screen.Width \ 2 - Me.Width \ 2
   Height = 2595 + 90
   Caption = "BigNum Calculator  HexDecBin"
   Show
   
   ReDim aT1Focus(0 To 5)  ' False
   aT1Focus(0) = True
   Text1_GotFocus (0)
   HSHexLen.Value = 9 ' Gives MaxHexInput = 16   ' default
   
   SetTextBoxProperties
   
   MaxHexInput = 16 '1-256 sure?
   SetLengths ' Needs MaxHexInput

   cmdClear1_Click
   cmdClear2_Click
   
   aBlock = False
   aLogic = False
   DisplayRes = 0
   LogicOp = bNop
   
   optVBASM(0) = True
   aVBASM = True
'--------------ASM------------------------
' Loading separate asm bin files for easier debugging
Dim FSpec$
         FSpec$ = PathSpec$ & "Dec2Bytes.bin"
         If FileExists(FSpec$) Then
            Loadmcode FSpec$
         Else
            MsgBox FSpec & vbCrLf & " Not there!", vbCritical, "ASM loading"
            End
         End If
         
         FSpec$ = PathSpec$ & "Bytes2Hex.bin"
         If FileExists(FSpec$) Then
            Loadmcode2 FSpec$
         Else
            MsgBox FSpec & vbCrLf & " Not there!", vbCritical, "ASM loading"
            End
         End If
         
         FSpec$ = PathSpec$ & "Bytes2Bits.bin"
         If FileExists(FSpec$) Then
            Loadmcode3 FSpec$
         Else
            MsgBox FSpec & vbCrLf & " Not there!", vbCritical, "ASM loading"
            End
         End If
         
         FSpec$ = PathSpec$ & "Bytes2Dec.bin"
         If FileExists(FSpec$) Then
            Loadmcode4 FSpec$
         Else
            MsgBox FSpec & vbCrLf & " Not there!", vbCritical, "ASM loading"
            End
         End If
         '-----------------------------------------
   
End Sub

Private Sub cmdTest_Click(Index As Integer)
Dim k As Long, Num As Long
Dim A$

   '     "0" - "9"
   ' Rnd  48 - 57
   '     10  -  15
   ' Rnd "A" - "F"
   
   Screen.MousePointer = vbHourglass
   If aLogic = False Then
      cmdConvCalc_Click 1
   End If
   If Index = 0 Then
   ' First number
   cmdClear1_Click
   A$ = ""
   'DoEvents
   For k = 1 To MaxHexInput
      If (k Mod 100) = 0 Then Refresh
      Num = CLng(15 * Rnd)
      If k = 1 And Num = 0 Then Num = 1
      'If Num > 15 Then Stop
      If Num < 10 Then
         A$ = A$ & Chr$(Num + 48)
      Else
         A$ = A$ & Chr$(Num + 55)
      End If
   Next k
      aBlock = False
      aT1Focus(0) = True
      Text1(0) = A$
   End If
   If Index = 0 Or Index = 1 Then
   ' Second number
   cmdClear2_Click
   A$ = ""
   For k = 1 To MaxHexInput
      If (k Mod 100) = 0 Then Refresh
      Num = CLng(15 * Rnd)
      If k = 1 And Num = 0 Then Num = 1
      'If Num > 15 Then Stop
      If Num < 10 Then
         A$ = A$ & Chr$(Num + 48)
      Else
         A$ = A$ & Chr$(Num + 55)
      End If
   Next k
   aBlock = False
   aT1Focus(3) = True
   Text1(3) = A$
   End If
   Screen.MousePointer = vbDefault
End Sub


Private Sub cmdSR_Click(Index As Integer)
'Shift, Roll & Not. Only applies to first number
Dim k As Long
Dim A$
   'Operate on bin string
   Screen.MousePointer = vbHourglass
   ClearResults
   ShiftRoll = bNop
   If Text1(2).Text = "" Then
      Screen.MousePointer = vbDefault
      Exit Sub
   End If
   If DisplayRes > 0 Then Unload frmZoom
   ShiftRoll = Index ' For Display
   A$ = ""
   ReDim aT1Focus(0 To 5)  ' False
   Text1(2).SetFocus
   Text1(2).BackColor = vbYellow
   aT1Focus(2) = True
   BinResult$ = ""
   Select Case ShiftRoll
   Case shL
      BinResult$ = BinString1$
      k = InStr(1, BinString$, "1")
      If k = 0 Then
         Screen.MousePointer = vbDefault
         Exit Sub
      End If
      BinResult$ = Mid$(BinResult$, k)
      Text1(2) = BinResult$
      If Len(BinResult$) < 4 * MaxHexInput Then
         BinResult$ = BinResult$ & "0"
      Else
         BinResult$ = BinString1$
      End If
   Case roL
      A$ = Left$(BinString1$, 1)
      BinResult$ = Mid$(BinString1$, 2)
      BinResult$ = BinResult$ & A$
   Case shR
      BinResult$ = Mid$(BinString1$, 1, Len(BinString1$) - 1)
      BinResult$ = "0" & BinResult$
   Case roR
      A$ = Right$(BinString1$, 1)
      BinResult$ = Mid$(BinString1$, 1, Len(BinString1$) - 1)
      BinResult$ = A$ & BinResult$
   Case bNot
      For k = 1 To Len(BinString1$)
         If Mid$(BinString1$, k, 1) = "0" Then
            Mid$(BinString1$, k, 1) = "1"
         Else
            Mid$(BinString1$, k, 1) = "0"
         End If
      Next k
      BinResult$ = BinString1$
   End Select
   
   BinString1$ = BinResult$
   A$ = BinResult$
   aBlock = True
   Text1(2).Text = A$
   aBlock = False
   
   Lab1(2) = Len(A$)
   BinString$ = A$
   Bin2Hex2Dec A$
   If HexString$ = "" Then HexString$ = "0"
   Text1(0).Text = HexString$
   Lab1(0) = Len(HexString$)
   HexString1$ = HexString$
   If DecString$ = "" Then DecString$ = "0"
   Text1(1).Text = DecString$
   Lab1(1) = Len(DecString$)
   DecString1$ = DecString$
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmdLogicCalc_Click(Index As Integer)
' Logic And, Or, Xor, Eqv, Imp
' Calc  Div/Mod, Mul, Abs(Sub), Add
Dim j As Long
Dim L1 As Long, L2 As Long, TheLen As Long
Dim A$
Dim aLenEqual As Boolean
Dim k As Long
Dim LTGT As Long
' For perms
Dim N$, R$
Dim NF$
Dim N1mR$
Dim N2$
Dim RF$
   
   Screen.MousePointer = vbUpArrow 'vbHourglass
   ClearResults
   LabMsg(0) = "": LabMsg(1) = ""
   LabOP = cmdLogicCalc(Index).Caption
   aLenEqual = False
   
   If DisplayRes > 0 Then Unload frmZoom
   
   If Text1(3) = "" Then
      Screen.MousePointer = vbDefault
      LogicOp = bNop
      LabOP = ""
      Exit Sub
   End If
   Select Case Index
   Case Is < bFac, bPerm, bComb
      If Text1(0).Text = "" Then
         Screen.MousePointer = vbDefault
         LogicOp = bNop
         LabOP = ""
         Exit Sub
      End If
   End Select
      
   LogicOp = Index ' For Display
   ' Carry out logic ops
   ' Make Len(BinString1$) = Len(BinString2$)
   BinResult$ = ""
   L1 = Len(BinString1$)
   L2 = Len(BinString2$)
   TheLen = L1
   Select Case LogicOp
   Case Is <> bDiv, Is <> bPerm, Is <> bComb
      If L1 > L2 Then
         BinString2$ = String$(L1 - L2, "0") & BinString2$
      ElseIf L2 > L1 Then
         BinString1$ = String$(L2 - L1, "0") & BinString1$
         TheLen = L2
      End If
   End Select
   
   aBlock = True
   Text1(2).Text = BinString1$
   Lab1(2) = Len(BinString1$)
   Text1(5).Text = BinString2$
   Lab1(5) = Len(BinString2$)
   aBlock = False
   
   ReDim aT1Focus(0 To 5)  ' False
   Text1(5).SetFocus
   Text1(5).BackColor = vbYellow
   aT1Focus(5) = True
   Select Case LogicOp
   Case bAnd
      For j = 1 To TheLen
         If Mid$(BinString2$, j, 1) = "1" Then
            BinResult$ = BinResult$ & Mid$(BinString1$, j, 1)
         ElseIf Len(BinResult$) > 0 Then
            BinResult$ = BinResult$ & "0"
         End If
      Next j
   Case bOr
      For j = 1 To TheLen
         If Mid$(BinString1$, j, 1) = "1" Or Mid$(BinString2$, j, 1) = "1" Then
            BinResult$ = BinResult$ & "1"
         ElseIf Len(BinResult$) > 0 Then
            BinResult$ = BinResult$ & "0"
         End If
      Next j
   Case bXor
      For j = 1 To TheLen
         If Mid$(BinString1$, j, 1) = Mid$(BinString2$, j, 1) Then
            BinResult$ = BinResult$ & "0"
         Else
            BinResult$ = BinResult$ & "1"
         End If
      Next j
   Case bEqv
      For j = 1 To TheLen
         If Mid$(BinString1$, j, 1) = Mid$(BinString2$, j, 1) Then
            BinResult$ = BinResult$ & "1"
         Else
            BinResult$ = BinResult$ & "0"
         End If
      Next j
   Case bImp
      For j = 1 To TheLen
         If Mid$(BinString1$, j, 1) = Mid$(BinString2$, j, 1) _
            Or Mid$(BinString2$, j, 1) = "1" Then
            BinResult$ = BinResult$ & "1"
         Else
            BinResult$ = BinResult$ & "0"
         End If
      Next j
      
   Case bDiv
      ' Numerator\Denominator  =  DecString1$\DecString2$
      If DecString2$ = "" Or DecString2$ = "0" Then
         MsgBox "Division by zero", vbCritical
         Exit Sub
      End If
      ClearLeadingZeros DecString1$
      ClearLeadingZeros DecString2$
      If Len(DecString2$) > Len(DecString1$) Then ' Denominator > Numerator
            SimpleIntDiv
      ElseIf Len(DecString2$) = Len(DecString1$) Then
         ' Check magnitudes
         k = 0
         LTGT = 2
         For j = 1 To Len(DecString1$)
            If Mid$(DecString1$, j, 1) > Mid$(DecString2$, j, 1) Then
               LTGT = 1
               Exit For
            End If
            If Mid$(DecString1$, j, 1) < Mid$(DecString2$, j, 1) Then
               LTGT = 0
               Exit For
            End If
         Next j
         If LTGT = 0 Then ' Numerator < Denominator
            SimpleIntDiv
         Else
            ' Denominator < Numerator
            FullIntDivision DecString1$, DecString2$
         End If
      Else  ' Len(DecString2$) < Len(DecString1$)
            ' Denominator < Numerator
         FullIntDivision DecString1$, DecString2$
      End If
      
   Case bMul
      L1 = MaxHexInput
      A$ = DecResult$
      L2 = Len(A$)
      MaxHexInput = 2 * L2
      SetLengths
      ' EKabiljo multiply routine
      ReDim aa(MaxLong)
      ReDim bb(MaxLong)
      ReDim cc(MaxLong)
      Call String2LongArray(DecString1$, aa, LengthA)
      Call String2LongArray(DecString2$, bb, LengthB)
      Call Multiply(aa, LengthA, bb, LengthB, cc, LengthC)
      DecResult$ = LongArray2String(cc, LengthC)
      MaxHexInput = L1
      SetLengths
   Case bSub
      DecResult$ = bSub_Routine(DecString1$, DecString2$)   ' |DecString1$-DecString2$|
   Case bAdd
         DecResult$ = Add(DecString1$, DecString2$)
   Case bFac
      LabOP = "Factorial"
      If Len(DecString2$) > 4 Then
         Screen.MousePointer = vbDefault
         LabMsg(0) = "Factorial N " & vbCrLf & "Only for <= 4 decimal digits"
         LogicOp = bNop
         Exit Sub
      Else
         LabMsg(0) = "": LabMsg(1) = ""
         DoEvents
         DecResult$ = bFac_Routine(DecString2$)
      End If
   
   Case bSqa
      LabOP = "Squared"
      ' EKabiljo multiply routine
      ReDim aa(MaxLong)
      ReDim bb(MaxLong)
      ReDim cc(MaxLong)
      Call String2LongArray(DecString2$, aa, LengthA)
      Call String2LongArray(DecString2$, bb, LengthB)
      Call Multiply(aa, LengthA, bb, LengthB, cc, LengthC)
      DecResult$ = LongArray2String(cc, LengthC)
   Case bCub
      LabOP = "Cubed"
      ' EKabiljo multiply routine
      ReDim aa(MaxLong)
      ReDim bb(MaxLong)
      ReDim cc(MaxLong)
      Call String2LongArray(DecString2$, aa, LengthA)
      Call String2LongArray(DecString2$, bb, LengthB)
      Call Multiply(aa, LengthA, bb, LengthB, cc, LengthC)
      Call Multiply(aa, LengthA, cc, LengthC, cc, LengthC)
      DecResult$ = LongArray2String(cc, LengthC)
      
   Case bPerm
      LabOP = "Permutations"  ' R from N,   N!/(N-R)!,  N=DecString2$ R=DecString1$
      R$ = DecString1$
      N$ = DecString2$
      If CheckMags(N$, R$) = 2 Then Exit Sub
      LabMsg(0) = "": LabMsg(1) = ""
      DoEvents
      If CheckMags(N$, R$) = 1 Then    ' N = r
         ' N!/0!
         DecResult$ = bFac_Routine(DecString2$)
      Else
         NF$ = bFac_Routine(N$)  ' N!
         N1mR$ = Subtract(N$, R$)   ' (N-R)
         N2$ = bFac_Routine(N1mR$)  ' (N-R)!
         FullIntDivision NF$, N2$   ' N!/(N-R)!
         ' OUT: Public DecResult$
      End If
      
   Case bComb
      LabOP = "Combinations"  ' N!/(r!(N-r)!)  N=N2 r=N1
      N$ = DecString2$
      R$ = DecString1$
      If CheckMags(N$, R$) = 2 Then Exit Sub
      LabMsg(0) = "": LabMsg(1) = ""
      DoEvents
      If CheckMags(N$, R$) = 1 Then   ' N = r
         ' N!/r! =1
         DecResult$ = "1"
      Else
         NF$ = bFac_Routine(N$)  ' N!
         RF$ = bFac_Routine(R$)  ' R!
         N1mR$ = Subtract(N$, R$)   ' (N-R)
         N2$ = bFac_Routine(N1mR$)  ' (N-R)!
         ' RF$ x N2$
         L1 = MaxHexInput
         A$ = DecResult$
         L2 = Len(A$)
         MaxHexInput = 2 * L2
         SetLengths
         
         ' EKabiljo multiply routine
         ReDim aa(MaxLong)
         ReDim bb(MaxLong)
         ReDim cc(MaxLong)
         Call String2LongArray(RF$, aa, LengthA)
         Call String2LongArray(N2$, bb, LengthB)
         Call Multiply(aa, LengthA, bb, LengthB, cc, LengthC)
         N2$ = LongArray2String(cc, LengthC)
         ' NF$/N2$     N!/(r!(N-r)!)
         FullIntDivision NF$, N2$   ' N!/(R!(N-R)!)
         ' OUT: Public DecResult$
         MaxHexInput = L1
         SetLengths
      End If
   
   End Select
   
   Screen.MousePointer = vbHourglass
   
   Select Case LogicOp
   Case bDiv    ' Div.       ' Show Result & Remainder
      
      ConvertDBH DecResult$
      
      A$ = DecRemain$
      Dec2Bin2Hex A$
      If HexString$ = "" Then HexString$ = "0"
      txtResult(3) = HexString$
      Lab1(9) = Len(HexString$)
      HexRemain$ = HexString$
      txtResult(4) = DecRemain$
      Lab1(10) = Len(DecRemain$)
      If BinString$ = "" Then BinString$ = "0"
      txtResult(5) = BinString$
      Lab1(11) = Len(BinString$)
      BinRemain$ = BinString$
   
   Case bMul, bSqa, bCub   ' DecResult$
      L1 = MaxHexInput
      L2 = Len(DecResult$)
      MaxHexInput = 2 * L2
      If LogicOp = bCub Then
         MaxHexInput = 3 * L2
      End If
      SetLengths
      
      ConvertDBH DecResult$
      
      MaxHexInput = L1
      SetLengths
   Case bSub
      
      ConvertDBH DecResult$
   
   Case bAdd   ' DecResult$
      L1 = MaxHexInput
      L2 = Len(DecResult$)
      MaxHexLen = (5 * (L2 - 1)) \ 6 + 2
      MaxHexInput = MaxHexLen - 4
      SetLengths
      
      ConvertDBH DecResult$
      
      MaxHexInput = L1
      SetLengths
   Case bFac   ' DecResult$
      L1 = MaxHexInput
      MaxHexInput = Len(DecResult$)
      SetLengths
      
      DecResult$ = DecResult$
      ConvertDBH DecResult$
      
      MaxHexInput = L1
      SetLengths
   
   Case bPerm, bComb ' DecResult$
      
      L1 = MaxHexInput
      MaxHexInput = Len(DecResult$)
      SetLengths
      
      A$ = DecResult$
      
      ConvertDBH A$
      
      MaxHexInput = L1
      SetLengths
      
   Case Else
      L1 = Len(BinString2$)
      L2 = Len(BinResult$)
      If L2 < L1 Then
         BinResult$ = String$(L1 - L2, "0") & BinResult$
      End If
      A$ = BinResult$
      Bin2Hex2Dec A$
      txtResult(2) = BinResult$
      Lab1(8) = Len(BinResult$)
      If HexString$ = "" Then HexString$ = "0"
      txtResult(0) = HexString$
      Lab1(6) = Len(HexString$)
      If DecString$ = "" Then DecString$ = "0"
      txtResult(1) = DecString$
      Lab1(7) = Len(DecString$)
      txtResult(1).Refresh
      HexResult$ = HexString$
      DecResult$ = DecString$
   End Select
   
   ' Show low ends of strings
   For k = 0 To 5
      Text1(k).SelStart = Len(Text1(k))
   Next k
   Picture1.Enabled = True
   For k = 0 To 5
      txtResult(k).SelStart = Len(txtResult(k))
   Next k
   Picture1.Enabled = False
   
   Erase aa(), bb(), cc()
   Screen.MousePointer = vbDefault
End Sub

Private Function CheckMags(N$, R$) As Long
   CheckMags = 2
   If Len(N$) > 4 Or Len(R$) > 4 Then
      Screen.MousePointer = vbDefault
      LabMsg(1) = "N or r > 4 dec digits"
      LogicOp = bNop
      Exit Function
   End If
   If (Len(R$) > Len(N$)) Then
      Screen.MousePointer = vbDefault
      LabMsg(1) = "r > N"
      LogicOp = bNop
      Exit Function
   End If
   If (Val(R$) > Val(N$)) Then
      Screen.MousePointer = vbDefault
      LabMsg(1) = "r > N"
      LogicOp = bNop
      Exit Function
   End If
   If (Val(R$) = Val(N$)) Then
      CheckMags = 1
      Exit Function
   End If
   CheckMags = 0
End Function
   

Private Sub ConvertDBH(ByVal A$)
' Get Hex & Bin string from Dec string
   Dec2Bin2Hex A$
   If HexString$ = "" Then HexString$ = "0"
   txtResult(0) = HexString$
   Lab1(6) = Len(HexString$)
   HexResult$ = HexString$
   txtResult(1) = DecResult$
   Lab1(7) = Len(DecResult$)
   If BinString$ = "" Then BinString$ = "0"
   txtResult(2) = BinString$
   Lab1(8) = Len(BinString$)
   BinResult$ = BinString$
End Sub

Private Sub ClearLeadingZeros(A$)
Dim aZero As Boolean
Dim k As Long
Dim b$, c$
   aZero = True
   b$ = A$
   A$ = ""
   For k = 1 To Len(b$)
      c$ = Mid$(b$, k, 1)
      If c$ <> "0" Then
         aZero = False
      End If
      If Not aZero Then
         A$ = A$ & c$
      End If
   Next k
   b$ = ""
End Sub

Private Function bSub_Routine(Num1$, Num2$) As String   ' |Num1$-Num2$|
Dim j As Long
Dim LTGT As Long
   If Len(Num2$) > Len(Num1$) Then ' Second > First number
         bSub_Routine = Subtract(Num2$, Num1$)
   ElseIf Len(Num2$) = Len(Num1$) Then
      ' Check magnitudes
      LTGT = 2
      For j = 1 To Len(Num1$)
         If Mid$(Num1$, j, 1) > Mid$(Num2$, j, 1) Then
            LTGT = 1
            Exit For
         End If
         If Mid$(Num1$, j, 1) < Mid$(Num2$, j, 1) Then
            LTGT = 0
            Exit For
         End If
      Next j
      If LTGT = 0 Then ' Second > First number
         bSub_Routine = Subtract(Num2$, Num1$)
      Else
         ' Second < First number
         bSub_Routine = Subtract(Num1$, Num2$)
      End If
   Else  ' Len(Num2$) < Len(Num1$)
         ' Second < First number
      bSub_Routine = Subtract(Num1$, Num2$)
   End If
End Function

Private Function bFac_Routine(Num2$) As String
Dim L2 As Long
Dim k As Long
   ' EKabiljo factorial routine
   ReDim aa(MaxLong)
   ReDim bb(MaxLong)
   L2 = Val(Num2$)
   aa(1) = 1
   LengthA = 1
   For k = 2 To L2
       bb(1) = k
       LengthB = 1
       Call Multiply(aa, LengthA, bb, LengthB, aa, LengthA)
   Next k
   bFac_Routine = LongArray2String(aa, LengthA)
End Function

Private Sub SimpleIntDiv()
   If DecString1$ = DecString2$ Then
      DecResult$ = "1"
      DecRemain$ = DecString1$
   Else
      DecResult$ = "0"
      DecRemain$ = "0"
      LabMsg(1) = "Denominator > Numerator"
    End If
End Sub

Private Sub MakeDecLensEqual()
' NOT USED
Dim L1 As Long, L2 As Long
   L1 = Len(DecString1$)
   L2 = Len(DecString2$)
   If L1 > L2 Then
      DecString2$ = String$(L1 - L2, "0") & DecString2$
   ElseIf L2 > L1 Then
      DecString1$ = String$(L2 - L1, "0") & DecString1$
   End If
End Sub

Private Sub FullIntDivision(ByVal D1$, ByVal D2$)
' Public DecString1$, DecString2$
' Public HexResult$, DecResult$, BinResult$  ' Logic result
' Public HexRemain$, DecRemain$, BinRemain$  ' Div remainder
' Public Counter() As Byte                   ' Div subtraction counter

' DecString1$\DecString2$ = D1$\D2$ = Quotient: Remainder
' Count continuous subtractions until remainder <= divider

' OUT: Public DecResult$, DecRemain$

Dim R1$, C1$, C2$
Dim CopyD2$
Dim A$
Dim k As Long
Dim L As Long
   
   CopyD2$ = D2$

   L = Len(D1$)
   If Len(D2$) < Len(D1$) Then
      D2$ = String$(Len(D1$) - Len(D2$), "0") & D2$
   End If
   
   Subtractor D1$, D2$
   
   R1$ = ""
   C1$ = ""
   For k = L To 1 Step -1
      If PDec1(k) <> 0 Then
        R1$ = R1$ & (PDec1(k))
      End If
      
      C1$ = C1$ & (Counter(k))
   Next k
   
   ' Code chunk here removed - unnecessary ??
   
   ClearLeadingZeros C1$
   
   DecResult$ = C1$
   If R1$ = "" Then R1$ = "0"
   DecRemain$ = R1$
   
   A$ = ""
   C1$ = ""
   C2$ = ""
   R1$ = ""
   CopyD2$ = ""
   Erase Counter()
   Erase PDec1()
End Sub

Private Sub Subtractor(ByVal a1$, ByVal a2$)
' a1$ - n * a2$  & rem
Dim L As Long, k As Long
Dim N1 As Byte, N2 As Byte
Dim PartNum As Byte
Dim Carry As Byte
      If a1$ = "" Then a1$ = "0" 'Stop
      L = Len(a1$)
      
      ReDim Dec1(L), Dec2(L)
      ReDim PDec1(L)
      ReDim Counter(L)
      
      If Len(a2$) < Len(a1$) Then
         a2$ = String$(Len(a1$) - Len(a2$), "0") & a2$
      End If
      
      For k = 1 To L
         Dec1(k) = Val(Mid$(a1$, Len(a1$) - (k - 1), 1))
         Dec2(k) = Val(Mid$(a2$, Len(a2$) - (k - 1), 1))
         PDec1(k) = Dec1(k)
      Next k
      Carry = 0
      DoEvents ' Help stop ridiculous XP whiteout
      Do
         Carry = 0
         For k = 1 To L
            PDec1(k) = Dec1(k)
            N1 = Dec1(k)
            N2 = Dec2(k) + Carry
            Carry = 0
            If N1 >= N2 Then
               PartNum = N1 - N2
            Else
               PartNum = (N1 + 10) - N2
               Carry = 1
               PDec1(k) = Dec1(k)         ' <<<<<<<<<<<
            End If
            Dec1(k) = PartNum
         Next k
      
         If Carry = 1 Then Exit Do
         IncrementCount Counter(), L
      Loop
End Sub


Private Sub IncrementCount(b() As Byte, L As Long)
' Public Counter() As Byte but general with B()
' DecByte Counter Incrementer for numbers > VB max
' SLIGHTLY FASTER THAN FOR LOOP IN EXE !
Dim k As Long, byt As Byte
   k = 1    ' For Option Base 1
   Do
      byt = b(k)
      If byt = 9 Then
         b(k) = 0
      Else
         b(k) = byt + 1
         Exit Do
      End If
      k = k + 1
      If k > L Then Exit Sub
   Loop
End Sub


'###############################################################

' Multiply routines base on code by
' EKabiljo  PSC CodeId=24960

Private Sub String2LongArray(A$, aa() As Long, LengthA As Long)
' RR Moves 4 characters at time into a long array
Dim i As Long
Dim MaxLen As Long
Dim Lmod As Long
   A$ = Trim$(A$)
   If Len(A$) = 0 Or A$ = "0" Then
      aa(1) = 0
      LengthA = 1
   Else
      MaxLen = Len(A$)
      If (MaxLen \ 4) * 4 = MaxLen Then
         LengthA = MaxLen \ 4
         For i = 1 To LengthA
            aa(i) = Mid$(A$, MaxLen - i * 4 + 1, 4)
         Next i
      Else
         LengthA = MaxLen \ 4 + 1
         Lmod = MaxLen Mod 4
         For i = 1 To LengthA - 1
            aa(i) = Mid$(A$, MaxLen - i * 4 + 1, 4)
         Next i
         aa(LengthA) = Mid$(A$, 1, Lmod)
      End If
   End If
End Sub

Private Function LongArray2String(c() As Long, Length As Long) As String
'RR Move long array into string
Dim i As Long
Dim Temp$
   Temp$ = Temp$ & Format$(c(Length), "0")
   For i = Length - 1 To 1 Step -1
       Temp$ = Temp$ & Format$(c(i), "0000")
   Next i
   LongArray2String = Temp$
End Function

Private Sub Multiply(a1() As Long, LengthA As Long, a2() As Long, LengthB As Long, res() As Long, LengthC As Long)
'RR Multiply a1() x a2() -> res() [LengthC]
Dim Carry As Long
Dim i As Long
Dim j As Long
Dim k As Long
   If (LengthB = 1 And a2(1) = 0) Or (LengthA = 1 And a1(1) = 0) Then
      res(1) = 0
      LengthC = 1
      Exit Sub
   End If
   If LengthB = 1 And a2(1) = 1 Then
      LengthC = LengthA
      CopyMemory res(1), a1(1), 4 * LengthC
      Exit Sub
   End If
   If LengthA = 1 And a1(1) = 1 Then
      LengthC = LengthB
      CopyMemory res(1), a2(1), 4 * LengthC
      Exit Sub
   End If
   Carry = 0
   FillMemory TLong(1), 4 * (LengthA + LengthB), 0
   For i = 1 To LengthB
      For j = 1 To LengthA
         k = i + j - 1
         TLong(k) = TLong(k) + a1(j) * a2(i)
         Carry = TLong(k) \ 10000
         'TLong(k) = TLong(k) Mod 10000
         TLong(k) = TLong(k) - 10000 * Carry
         TLong(k + 1) = TLong(k + 1) + Carry
      Next j
   Next i
   LenTLong = LengthA + LengthB
   Do Until TLong(LenTLong) <> 0 Or LenTLong = 1
      LenTLong = LenTLong - 1
   Loop
   LengthC = LenTLong
   CopyMemory res(1), TLong(1), 4 * LengthC
End Sub

'###############################################################

Private Function Subtract(ByVal a1$, ByVal a2$) As String
' Dec A1$ - Dec A2$
Dim L As Long, k As Long
Dim N1 As Byte, N2 As Byte
Dim PartNum As Byte
Dim Carry As Byte
Dim R1$ ', C2$
Dim aZero As Boolean
      
      L = Len(a1$)
      
      ReDim Dec1(L), Dec2(L)
      
      If Len(a2$) < Len(a1$) Then
         a2$ = String$(Len(a1$) - Len(a2$), "0") & a2$
      End If
      
      For k = 1 To L
         Dec1(k) = Val(Mid$(a1$, Len(a1$) - (k - 1), 1))
         Dec2(k) = Val(Mid$(a2$, Len(a2$) - (k - 1), 1))
      Next k
      Carry = 0
      DoEvents ' Help stop ridiculous XP whiteout
      For k = 1 To L                   ' 1 2 3 4
         N1 = Dec1(k)
         N2 = Dec2(k) + Carry          ' 0 5 2 3
         Carry = 0
         If N1 >= N2 Then
            PartNum = N1 - N2
         Else
            PartNum = (N1 + 10) - N2
            Carry = 1
         End If
         Dec1(k) = PartNum
      Next k
      R1$ = ""
      aZero = True
      For k = L To 1 Step -1
         If Dec1(k) <> 0 Then aZero = False
         If Not aZero Then
            R1$ = R1$ & (Dec1(k))
         End If
      Next k
      Subtract = R1$
      R1$ = ""
End Function

Private Function Add(ByVal a1$, ByVal a2$) As String
' Dec A1$ + Dec A2$
Dim L As Long, k As Long
Dim N1 As Byte, N2 As Byte
Dim PartNum As Byte
Dim Carry As Byte
Dim R1$ ', C2$
Dim aZero As Boolean
      
      If Len(a2$) < Len(a1$) Then
         a2$ = String$(Len(a1$) - Len(a2$), "0") & a2$
      End If
      a1$ = "0" & a1$
      a2$ = "0" & a2$
      L = Len(a1$)
      
      ReDim Dec1(L), Dec2(L)
      
      For k = 1 To L
         Dec1(k) = Val(Mid$(a1$, Len(a1$) - (k - 1), 1))
         Dec2(k) = Val(Mid$(a2$, Len(a2$) - (k - 1), 1))
      Next k
      Carry = 0
      DoEvents ' Help stop ridiculous XP whiteout
      For k = 1 To L                   ' 1 2 3 4
         N1 = Dec1(k)
         N2 = Dec2(k)         ' 0 5 2 3
         PartNum = N1 + N2 + Carry
         Carry = PartNum \ 10
         Dec1(k) = PartNum - Carry * 10
      Next k
      R1$ = ""
      aZero = True
      For k = L To 1 Step -1
         If Dec1(k) <> 0 Then aZero = False
         If Not aZero Then
            R1$ = R1$ & (Dec1(k))
         End If
      Next k
      Add = R1$
      R1$ = ""
End Function

Private Sub cmdPad_Click(Index As Integer)
' On-screen keypad, Mouse Input
Dim Key As Integer, i As Integer
Dim k As Long
Dim A$
   LabMsg(0) = "": LabMsg(1) = ""
   For k = 0 To 5
      txtResult(k) = ""
   Next k
   Key = 48 + Index
   If Key > 57 Then Key = Key + 7
   
   For k = 0 To 5
      If aT1Focus(k) Then
         i = k
         Exit For
      End If
   Next k
   If k = 6 Then Exit Sub
   IP = Text1(i).SelStart
   Text1_KeyPress i, Key
   If Key <> 0 Then
      A$ = Text1(i).Text
      If IP <> 0 Then
         A$ = Left$(A$, IP) & Chr$(Key) & Mid$(A$, IP + 1)
      Else
         A$ = Chr$(Key) & A$
      End If
      Text1(i).Text = A$
      fraHexLen.Enabled = False
      HSHexLen.Enabled = False
      LabOP = ""
   End If
   Text1(i).SelStart = IP + 1 'Len(A$)
   Lab1(i) = Len(Text1(i).Text)
   Text1(i).SetFocus
End Sub

Private Sub LabMsg_Click(Index As Integer)
   LabMsg(0) = "": LabMsg(1) = ""
End Sub

Private Sub optVBASM_Click(Index As Integer)
   aVBASM = Index - 1
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
' Limit input digits
   ClearResults
   Select Case Index
   Case 0, 3  ' Hex
      If KeyAscii <> 8 Then
         If Len(Text1(Index).Text) >= MaxHexInput Then
         KeyAscii = 0: Exit Sub
         End If
         ' Capitalize lower case letters
         If (KeyAscii >= 97 And KeyAscii <= 102) Then KeyAscii = KeyAscii - 32
         ' Take numbers 0-9 letters A-F & Backspace
         If (KeyAscii >= 48 And KeyAscii <= 57) Or _
            (KeyAscii >= 65 And KeyAscii <= 70) Then
         Else: KeyAscii = 0
         End If
      End If
      Lab1(Index) = Len(Text1(Index).Text) ' ActHexLen
   Case 1, 4  ' Dec
      If KeyAscii <> 8 Then
         If Len(Text1(Index).Text) >= MaxDecInput Then KeyAscii = 0: Exit Sub
         ' Take numbers 0-9
         If (KeyAscii < 48 Or KeyAscii > 57) Then KeyAscii = 0: Exit Sub
      End If
      Lab1(Index) = Len(Text1(Index).Text) ' ActDecLen
   Case 2, 5  ' Bin
      If KeyAscii <> 8 Then
         If Len(Text1(Index).Text) >= MaxBinInput Then KeyAscii = 0: Exit Sub
         ' Take numbers 0-1
         If (KeyAscii < 48 Or KeyAscii > 49) Then KeyAscii = 0: Exit Sub
      End If
      Lab1(Index) = Len(Text1(Index).Text) ' ActBinLen
   End Select
End Sub

Private Sub Text1_Change(Index As Integer)
' Show results
' 0,3 Hex  1,4 Dec  2,5 Bin input
Dim A$
Dim k As Long
Dim aCHK  As Boolean
   aCHK = False
   If aBlock Then Exit Sub
   If aT1Focus(Index) = False Then Exit Sub
   Select Case Index
   Case 0, 3  ' Hex
      If CheckInput(Index, A$) Then
         aCHK = True
         Lab1(Index) = Len(Text1(Index).Text)         ' ActHexLen
         HexString$ = A$
         Hex2Bin2Dec A$
         Text1(Index + 1).Text = DecString$
         Lab1(Index + 1) = Len(Text1(Index + 1).Text) ' ActDecLen
         Text1(Index + 2).Text = BinString$
         Lab1(Index + 2) = Len(Text1(Index + 2).Text) ' ActBinLen
      End If
   Case 1, 4  ' Dec
      If CheckInput(Index, A$) Then
         aCHK = True
         Lab1(Index) = Len(Text1(Index).Text)         ' ActDecLen
         DecString$ = A$
         Dec2Bin2Hex A$
         Text1(Index - 1).Text = HexString$
         Lab1(Index - 1) = Len(Text1(Index - 1).Text) ' ActHexLen
         Text1(Index + 1).Text = BinString$
         Lab1(Index + 1) = Len(Text1(Index + 1).Text) ' ActBinLen
      End If
   Case 2, 5  ' Bin
      If CheckInput(Index, A$) Then
         aCHK = True
         Lab1(Index) = Len(Text1(Index).Text)         ' ActBinLen
         BinString$ = A$
         Bin2Hex2Dec A$
         Text1(Index - 2).Text = HexString$
         Lab1(Index - 2) = Len(Text1(Index - 2).Text) ' ActHexLen
         Text1(Index - 1).Text = DecString$
         Lab1(Index - 1) = Len(Text1(Index - 1).Text) ' ActDecLen
      End If
   End Select
   If A$ <> "0" Then
   If A$ <> "" And Not aCHK Then
      ' Probably pasting into filled text box
      ' exceeding allowed length
      HexString$ = ""
      DecString$ = ""
      BinString$ = ""
      If Index < 3 Then
         For k = 0 To 2
            Text1(k).Text = ""
            Lab1(k) = 0
         Next k
      Else
         For k = 3 To 5
            Text1(k).Text = ""
            Lab1(k) = 0
         Next k
      End If
      MsgBox "Exceeded Max hex input", vbCritical, "Pasting?"
   End If
   End If
   If Index < 3 Then
      HexString1$ = HexString$
      DecString1$ = DecString$
      BinString1$ = BinString$
   ElseIf Index < 6 Then
      HexString2$ = HexString$
      DecString2$ = DecString$
      BinString2$ = BinString$
   End If

End Sub

Private Function CheckInput(Index As Integer, A$) As Boolean
' IN:  Index Text1(Index)
' OUT: A$ = Textbox string for processing or Error
' An error could arise from Paste into text box
Dim k As Long, L As Long
Dim L1 As Long
   CheckInput = False
   Select Case Index
   Case 0, 3  ' Any Hex input?
      A$ = Text1(Index).Text
      If A$ = "" Then
         If Index = 0 Then cmdClear1_Click Else cmdClear2_Click
         Exit Function
      End If
      For k = 1 To Len(A$)
         If InStr(NumberHex$, Mid$(A$, k, 1)) = 0 Then
            MsgBox "Hex error" & vbCrLf & " Not 0 to 9 or A to F", vbCritical
            If Index = 0 Then cmdClear1_Click Else cmdClear2_Click
            A$ = ""
            Exit Function
         End If
      Next k
      If Len(A$) > MaxHexInput Then Exit Function
      L = Len(A$)
      If L = 1 And A$ = "0" Then
         If Index = 0 Then cmdClear1_Click Else cmdClear2_Click
         Exit Function
      ElseIf Left$(A$, 1) = "0" Then
         A$ = Right$(A$, L - 1)
         aBlock = True
         Text1(Index).Text = A$
         aBlock = False
      End If
      CheckInput = True
      fraHexLen.Enabled = False
      HSHexLen.Enabled = False
      LabOP = ""
   Case 1, 4   ' Any Dec input?
      A$ = Text1(Index).Text
      If A$ = "" Then
         If Index = 1 Then cmdClear1_Click Else cmdClear2_Click
         Exit Function
      End If
      For k = 1 To Len(A$)
         If InStr(NumberDec$, Mid$(A$, k, 1)) = 0 Then
            MsgBox "Dec error" & vbCrLf & " Not 0 to 9", vbCritical
            If Index = 1 Then cmdClear1_Click Else cmdClear2_Click
            A$ = ""
            Exit Function
         End If
      Next k
      If Len(A$) > MaxDecInput Then Exit Function
      L = Len(A$)
      If L = 1 And A$ = "0" Then
         If Index = 1 Then cmdClear1_Click Else cmdClear2_Click
         Exit Function
      ElseIf Left$(A$, 1) = "0" Then
         A$ = Mid$(A$, 2)
         aBlock = True
         Text1(Index).Text = A$
         aBlock = False
      End If
      CheckInput = True
      fraHexLen.Enabled = False
      HSHexLen.Enabled = False
      LabOP = ""
   Case 2, 5  ' Any Bin input?
      A$ = Text1(Index).Text
      If A$ = "" Then
         If Index = 2 Then cmdClear1_Click Else cmdClear2_Click
         Exit Function
      End If
      For k = 1 To Len(A$)
         If InStr(NumberBin$, Mid$(A$, k, 1)) = 0 Then
            MsgBox "Bin error" & vbCrLf & " Not 0 or 1", vbCritical
            If Index = 2 Then cmdClear1_Click Else cmdClear2_Click
            A$ = ""
            Exit Function
         End If
      Next k
      If Len(A$) > MaxBinInput Then Exit Function
      L = Len(A$)
      If L = 1 And A$ = "0" Then
         If Index = 2 Then cmdClear1_Click Else cmdClear2_Click
         Exit Function
      ElseIf Left$(A$, 1) = "0" Then   ' Slow for long strings
         A$ = Right$(A$, L - 1)
         aBlock = True
         Text1(Index).Text = A$
         aBlock = False
      End If
      CheckInput = True
      fraHexLen.Enabled = False
      HSHexLen.Enabled = False
      'LabOP = ""
   End Select
End Function

Private Sub delLR_Click(Index As Integer)
'Delete digit of focussed textbox from Left or Right
Dim k As Long, i As Integer
   ClearResults
   For k = 0 To 5
      If aT1Focus(k) Then
         i = k
         Exit For
      End If
   Next k
   If k = 6 Then Exit Sub

   Select Case Index
   Case 0   ' Delete from left
      If Text1(i).Text <> "" Then
         Text1(i).Text = Mid$(Text1(i).Text, 2)
         Text1(i).SelStart = 0
      End If
   Case 1   ' Delete from right
      If Text1(i).Text <> "" Then
         Text1(i).Text = Left$(Text1(i).Text, Len(Text1(i).Text) - 1)
         k = Len(Text1(i).Text)
         Text1(i).SelStart = k
      End If
   End Select
End Sub


Private Sub SetLengths()
' IN: MaxHexInput
   MaxHexLen = MaxHexInput + 4
   MaxDecLen = MaxHexLen + (MaxHexLen \ 5) + 1
   MaxBinLen = 4 * MaxHexLen
   MaxBinLen = 8 * (MaxBinLen \ 8)  ' Make multiple of 8
   ReDim HexBytes(MaxHexLen)
   ReDim DecBytes(MaxDecLen)
   ReDim BinBits(MaxBinLen)
   MaxByteLen = MaxBinLen \ 8
   ReDim BinBytes(MaxByteLen)
   MaxDecInput = MaxDecLen - 2
   MaxBinInput = MaxBinLen - 8
End Sub

Private Sub HSHexLen_Scroll()
   Call HSHexLen_Change
End Sub
Private Sub HSHexLen_Change()
' Max hex input len
Dim V As Long
   'MaxHexInput = HSHexLen.Value
   V = HSHexLen.Value
   Select Case V
   Case 1 To 7: MaxHexInput = V
   Case Else: MaxHexInput = 8 * (V - 7)
   End Select
   LabHexLen = Str$(MaxHexInput)
   SetLengths
End Sub

Private Sub Text1_GotFocus(Index As Integer)
Dim k As Long
   ReDim aT1Focus(0 To 5)  ' False
   For k = 0 To 5
      Text1(k).BackColor = vbWhite
   Next k
   Text1(Index).BackColor = &HC0FFFF 'vbYellow
   aT1Focus(Index) = True
End Sub


Private Sub cmdConvCalc_Click(Index As Integer)
Dim k As Long
   LabMsg(0) = "": LabMsg(1) = ""
   If aLogic = True Then ' Conv
      Me.Height = 2595
      aLogic = False
   ElseIf Index = 1 And aLogic = False Then ' Logic Calc
      Me.Height = 7800
      aLogic = True
   End If
   Me.Refresh
End Sub

Private Sub cmdDisplay_Click()
' Display all results
   LabMsg(0) = "": LabMsg(1) = ""
   If DisplayRes > 0 Then
      Unload frmZoom
      DisplayRes = 0
   Else
      DisplayRes = 1
      Load frmZoom
   End If
End Sub

Private Sub cmdHelp_Click()
   LabMsg(0) = "": LabMsg(1) = ""
   If DisplayRes > 0 Then
      Unload frmZoom
      DisplayRes = 0
   Else
      DisplayRes = 2
      Load frmZoom
   End If
End Sub

Private Sub cmdClear1_Click()
' Clear all Text1(0 to 2)
Dim k As Long
   For k = 0 To 2
      Text1(k) = ""
      Lab1(k) = ""
   Next k
   fraHexLen.Enabled = True
   HSHexLen.Enabled = True
   HexString$ = ""
   DecString$ = ""
   BinString$ = ""
   HexString1$ = ""
   DecString1$ = ""
   BinString1$ = ""
   LabOP = ""
   LabMsg(0) = "": LabMsg(1) = ""
   If DisplayRes > 0 Then Unload frmZoom
End Sub

Private Sub cmdClear2_Click()
' Clear all Text1(3 to 5), Text2, LabRemain
Dim k As Long
   For k = 3 To 5
      Text1(k) = ""
      Lab1(k) = ""
   Next k
   ClearResults
   HexString$ = ""
   DecString$ = ""
   BinString$ = ""
   HexString2$ = ""
   DecString2$ = ""
   BinString2$ = ""
   LabMsg(0) = "": LabMsg(1) = ""
End Sub

Private Sub ClearResults()
Dim k As Long
   For k = 0 To 5
      txtResult(k) = ""
      Lab1(k + 6) = ""
   Next k
   HexResult$ = ""
   DecResult$ = ""
   BinResult$ = ""
   LabOP = ""
   If DisplayRes > 0 Then Unload frmZoom
   LogicOp = bNop
   LabMsg(0) = "": LabMsg(1) = ""
End Sub

Private Sub SetTextBoxProperties()
Dim k As Long
Dim M$, A$
   
cmdLogicCalc(10).FontName = "Arial"
cmdLogicCalc(10).FontSize = 10

cmdLogicCalc(10).Caption = "N " & Chr$(178)
cmdLogicCalc(11).FontName = "Arial"
cmdLogicCalc(11).FontSize = 10
cmdLogicCalc(11).Caption = "N " & Chr$(179)
   ' Locked True/False  disallow/allow keyboard input
   For k = 0 To 5
      With Text1(k)
         .Locked = False 'True
         .Alignment = vbRightJustify
         .MaxLength = 0
         .FontName = "Courier New"
         .FontSize = 9 '8
      End With
   Next k
   For k = 0 To 5
      With txtResult(k)
         .Enabled = True
         .Alignment = vbRightJustify
         .FontName = "Courier New"
         .FontSize = 9 '8
         .Enabled = True
         .Alignment = vbRightJustify
         .FontName = "Courier New"
         .FontSize = 9 '8
      End With
   Next k
   Shape1(0).FillColor = &HFFD3D3
   Picture1.BackColor = Shape1(0).FillColor
   Picture1.Enabled = False
   cmdConvCalc(1).BackColor = Shape1(0).FillColor
   LabMsg(0) = "": LabMsg(1) = ""
   LabMsg(0).BackColor = Shape1(0).FillColor
   LabMsg(1).BackColor = Shape1(0).FillColor

M$ = "|   "
A$ = ""
For k = 1 To 93
   A$ = A$ & M$
Next k
A$ = A$
LabMarks(0).Caption = A$
For k = 1 To 3
   LabMarks(k).Caption = A$
   LabMarks(k).BackColor = Shape1(0).FillColor
Next k
   LabOP = ""
   k = Rnd
End Sub

Private Sub cmdExit_Click()
' Exit
   Form_Unload 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim Form As Form
   FreeLibrary m_hMod
   For Each Form In Forms
      Unload Form
      Set Form = Nothing
   Next Form
   End
End Sub

