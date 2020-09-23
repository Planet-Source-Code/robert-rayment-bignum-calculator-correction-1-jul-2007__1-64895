VERSION 5.00
Begin VB.Form frmZoom 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0C0&
   Caption         =   "frmZoom"
   ClientHeight    =   4260
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   4680
   Icon            =   "frmZoom.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   284
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCBViewer 
      Caption         =   "View C'Board"
      Height          =   300
      Left            =   2370
      TabIndex        =   4
      Top             =   3885
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save As"
      Height          =   285
      Left            =   105
      TabIndex        =   3
      Top             =   3900
      Width           =   870
   End
   Begin VB.CommandButton cmdClipBoard 
      Caption         =   "--> ClipBoard"
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      Top             =   3900
      Width           =   1170
   End
   Begin VB.TextBox txtZoom 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3810
      Left            =   30
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   30
      Width           =   4560
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   285
      Left            =   3720
      TabIndex        =   0
      Top             =   3900
      Width           =   825
   End
End
Attribute VB_Name = "frmZoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' frmZoom.frm

Option Explicit

' WIN API Stay on top and position form
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, _
ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, _
ByVal wi As Long, ByVal ht As Long, ByVal wFlags As Long) As Long

Private Const hWndInsertAfter = -1
Private Const wFlags = &H40 Or &H20

Private Declare Function GetSystemDirectory Lib "Kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Private A$
Private CommonDialog1 As OSDialog


Private Sub cmdCBViewer_Click()
Dim F$
Dim ret As Long
   F$ = Space$(255)
   ret = GetSystemDirectory(F$, 255)
   F$ = Left$(F$, ret) & "\clipbrd.exe"
   If FileExists(F$) Then
      Shell F$, vbMaximizedFocus
      Unload Me
   Else
      MsgBox F$ & "  " & vbCrLf & " Not there!", vbInformation, "View Clipborad"
   End If
End Sub

Private Sub Form_Load()
Dim k As Long
Dim L1$, L2$, L3$
Dim FileNum As Long
   txtZoom.Text = ""
   Me.Caption = " Full Display"
   Me.BackColor = &HFFD3D3
    
   k = SetWindowPos(frmZoom.hWnd, hWndInsertAfter, 20, 1000 \ STY, 4800 / STX, 7800 / STY, wFlags)
   Show
  
   If DisplayRes = 1 Then
  
      If HexString1$ <> "" Or LogicOp = bFac Then
         
         L1$ = " [" & Trim$(Str$(Len(HexString1$))) & "]"
         L2$ = " [" & Trim$(Str$(Len(DecString1$))) & "]"
         L3$ = " [" & Trim$(Str$(Len(BinString1$)))
         If Left$(BinString1$, 1) = "0" Then L3$ = L3$ & " inc lead 0s"
         L3$ = L3$ & "]"
         A$ = "First number" & vbCrLf & "Hex" & L1$ & vbCrLf
         A$ = A$ & HexString1$ & vbCrLf & "Dec" & L2$ & vbCrLf & _
                   DecString1$ & vbCrLf & "Bin" & L3$ & vbCrLf & _
                   BinString1$ & vbCrLf & vbCrLf
         If aLogic And (HexString2$ <> "") Then
            L1$ = " [" & Trim$(Str$(Len(HexString2$))) & "]"
            L2$ = " [" & Trim$(Str$(Len(DecString2$))) & "]"
            L3$ = " [" & Trim$(Str$(Len(BinString2$)))
            If Left$(BinString2$, 1) = "0" Then L3$ = L3$ & " inc lead 0s"
            L3$ = L3$ & "]"
            A$ = A$ & "Second number" & vbCrLf & "Hex" & L1$ & vbCrLf
            A$ = A$ & HexString2$ & vbCrLf & "Dec" & L2$ & vbCrLf & _
                      DecString2$ & vbCrLf & "Bin" & L3$ & vbCrLf & _
                      BinString2$ & vbCrLf & vbCrLf
            If LogicOp <> bNop Then
               Select Case LogicOp
               Case bAnd: A$ = A$ & "And"
               Case bOr: A$ = A$ & "Or"
               Case bXor: A$ = A$ & "Xor"
               Case bEqv: A$ = A$ & "Eqv"
               Case bImp: A$ = A$ & "Imp"
               Case bDiv: A$ = A$ & "Div"
               Case bMul: A$ = A$ & "Mul"
               Case bSub: A$ = A$ & "Abs Sub"
               Case bAdd: A$ = A$ & "Add"
               Case bFac: A$ = A$ & "Factorial"
               Case bSqa: A$ = A$ & "Squared"
               Case bCub: A$ = A$ & "Cubed"
               Case bPerm: A$ = A$ & "Permutations"
               Case bComb: A$ = A$ & "Combinations"
               End Select
               L1$ = " [" & Trim$(Str$(Len(HexResult$))) & "]"
               L2$ = " [" & Trim$(Str$(Len(DecResult$))) & "]"
               L3$ = " [" & Trim$(Str$(Len(BinResult$))) & "]"
               A$ = A$ & " result" & vbCrLf & "Hex" & L1$ & vbCrLf
               A$ = A$ & HexResult$ & vbCrLf & "Dec" & L2$ & vbCrLf _
                       & DecResult$ & vbCrLf & "Bin" & L3$ & vbCrLf _
                       & BinResult$
               A$ = A$ & vbCrLf
               If LogicOp = bDiv Then
                  L1$ = " [" & Trim$(Str$(Len(HexRemain$))) & "]"
                  L2$ = " [" & Trim$(Str$(Len(DecRemain$))) & "]"
                  L3$ = " [" & Trim$(Str$(Len(BinRemain$))) & "]"
                  A$ = A$ & vbCrLf & "Div remainder" & vbCrLf & "Hex" & L1$ & vbCrLf
                  A$ = A$ & HexRemain$ & vbCrLf & "Dec" & L2$ & vbCrLf _
                          & DecRemain$ & vbCrLf & "Bin" & L3$ & vbCrLf _
                          & BinRemain$
               End If
            End If
         End If
      End If
  
  Else   ' ? Help
  
      If FileExists(PathSpec$ & "OPS.txt") Then
         FileNum = FreeFile
         Open PathSpec$ & "OPS.txt" For Binary As #FileNum
         A$ = Space$(LOF(FileNum))
         Get #FileNum, , A$
         Close FileNum
      Else
         MsgBox "OPS.txt file not there"
      End If
  
  End If
  
  txtZoom.Text = A$ & vbCrLf
  txtZoom.Refresh
End Sub

Private Sub cmdClipBoard_Click()
'ClipB
   If A$ <> "" Then
      Clipboard.Clear
      Clipboard.Clear
      Clipboard.SetText A$
      A$ = ""
   End If
End Sub

Private Sub cmdSave_Click()
' Save txtZoom contents to txt file
Dim Title$, Filt$, InDir$
Dim FIndex As Long, FileNum As Long
   If HexString1$ <> "" Or LogicOp = bFac Then
      Filt$ = "Text(*.txt)|*.txt"
      FileSpec$ = ""
      Title$ = "Save Result"
      InDir$ = CurrPath$
      
      Set CommonDialog1 = New OSDialog
      CommonDialog1.ShowSave FileSpec$, Title$, Filt$, InDir$, "", Me.hWnd, FIndex
      Set CommonDialog1 = Nothing
      
      If Len(FileSpec$) > 0 Then
         FileNum = FreeFile
         Open FileSpec$ For Output As #FileNum
         Print #FileNum, A$
         Close FileNum
      End If
   End If
   A$ = ""
End Sub

Private Sub Form_Resize()
Dim pos As Long
   If WindowState <> vbMinimized Then
      txtZoom.Width = Me.Width \ STX - 16
      txtZoom.Height = (Me.Height \ STY) - 70
      pos = txtZoom.Top + txtZoom.Height + 4
      cmdClose.Top = pos 'txtZoom.Top + txtZoom.Height + 4
      cmdClipBoard.Top = pos 'txtZoom.Top + txtZoom.Height + 4
      cmdCBViewer.Top = pos ' txtZoom.Top + txtZoom.Height + 4
      cmdSave.Top = pos 'txtZoom.Top + txtZoom.Height + 4
   End If
End Sub

Private Sub cmdClose_Click()
   Form_Unload 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
   DisplayRes = 0
   A$ = ""
   Unload Me
End Sub
