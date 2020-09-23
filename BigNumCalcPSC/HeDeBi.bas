Attribute VB_Name = "HeDeBi"
' HeDeBi.bas
' By Robert Rayment

' Method applicable to any input string lengths.

Option Explicit

Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

'------------------------------------------------
' Only these necessary for VB conversions
Option Base 1

Public Declare Sub FillMemory Lib "kernel32.dll" Alias "RtlFillMemory" _
(Destination As Any, ByVal Length As Long, ByVal Fill As Byte)

Public HexString$, DecString$, BinString$    ' General

Public HexBytes() As Byte
Public MaxHexLen As Long

Public DecBytes() As Byte
Public MaxDecLen As Long

Public BinBytes() As Byte
Public MaxByteLen As Long

Public BinBits() As Byte
Public MaxBinLen As Long

'------------------------------------------------

Public HexString1$, DecString1$, BinString1$ ' 1st number
Public HexString2$, DecString2$, BinString2$ ' 2nd number

Public HexResult$, DecResult$, BinResult$    ' Logic result

Public HexRemain$, DecRemain$, BinRemain$    ' Div remainder
Public Counter() As Byte                     ' Div subtraction counter
Public PDec1() As Byte
Public SumCount() As Byte

Public aLogic As Boolean
Public DisplayRes As Long

Public Enum Logics
   bAnd = 0
   bOr
   bXor
   bEqv
   bImp
   bDiv
   bMul
   bSub
   bAdd
   bFac
   bSqa
   bCub
   bPerm
   bComb
   bNop = 14
End Enum
Public LogicOp

Public Enum ShiftRol
   shL = 0
   roL
   shR
   roR
   bNot = 4
End Enum
Public ShiftRoll
   
Public PathSpec$, CurrPath$, FileSpec$

Public STX As Long  ' ScreenTwipsPerPixelX/Y
Public STY As Long

'------------------------------------------------
' For ASM if wanted
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" _
(ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, _
ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'' For machine code if used
' All separated for easier debugging
Public MCCode() As Byte
Public MCCode2() As Byte
Public MCCode3() As Byte
Public MCCode4() As Byte
Public ptrMC As Long
Public ptrMC2 As Long
Public ptrMC3 As Long
Public ptrMC4 As Long
Public ptrBinBytes As Long
Public ptrDecBytes As Long
Public ptrHexBytes As Long
Public ptrBinBits As Long

Public aVBASM As Boolean
 '------------------------------------------------


Public Sub Hex2Bin2Dec(A$)
' IN:  A$ = Hex string
' OUT: DecString$, BinString$
Dim k As Long, j As Long
Dim b As Byte
   b = 48
   FillMemory HexBytes(1), MaxHexLen, b    ' "0" to HexBytes()
   ' Fill HexBytes() from A$
   ' NB HexBytes(1) is Right char of A$ ie @ Len(A$)
   ' ie CopyMemory cannot be used here
   For k = 1 To Len(A$)
      HexBytes(k) = Asc(Mid$(A$, Len(A$) - (k - 1), 1))
   Next k
    
   Hex2Bytes A$
   A$ = ""
   Bytes2Bits   ' BinBits()
   Bytes2Dec    ' DecBytes()   ' Slow in VB
   
   ' Get Dec result
   DecString$ = ""
   For k = MaxDecLen To 1 Step -1
      If DecBytes(k) <> 48 Then Exit For
   Next k
   For j = k To 1 Step -1
         DecString$ = DecString$ + Chr$(DecBytes(j))
   Next j
   
   ' Get Binary result
   BinString$ = ""
   For k = MaxBinLen To 1 Step -1
      If BinBits(k) <> 48 Then Exit For
   Next k
   For j = k To 1 Step -1
         BinString$ = BinString$ + Chr$(BinBits(j))
   Next j
End Sub

Public Sub Dec2Bin2Hex(A$)
'IN:  A$ = Dec string
'OUT: HexString$, BinString$
Dim k As Long, j As Long
Dim b As Byte
Dim U As Long
   b = 48
   FillMemory DecBytes(1), MaxDecLen, b    ' "0" to DecBytes()
   ' Fill DecBytes() from A$ (DecString$)
   U = UBound(DecBytes(), 1)
   For k = 1 To Len(A$)
      DecBytes(k) = Asc(Mid$(A$, Len(A$) - (k - 1), 1))
   Next k
   A$ = ""
    
   Dec2Bytes
   Bytes2Bits   ' BinBits()
   Bytes2Hex    ' HexBytes()   ' Bit slow in VB
   
   ' Get Hex result
   HexString$ = ""
   For k = MaxHexLen To 1 Step -1
      If HexBytes(k) <> 48 Then Exit For
   Next k
   For j = k To 1 Step -1
         HexString$ = HexString$ + Chr$(HexBytes(j))
   Next j
   
   ' Get Binary result
   BinString$ = ""
   For k = MaxBinLen To 1 Step -1
      If BinBits(k) <> 48 Then Exit For
   Next k
   For j = k To 1 Step -1
         BinString$ = BinString$ + Chr$(BinBits(j))
   Next j
End Sub

Public Sub Bin2Hex2Dec(A$)
'IN:  A$ = Bin string
'OUT: HexString$, DecString$
Dim k As Long, j As Long
   ' Zero BinBits
   ReDim BinBits(MaxBinLen)  ' to zero
   
   ' Fill BinBits() from A$
   k = Len(A$)
   For k = 1 To Len(A$)
      BinBits(k) = Asc(Mid$(A$, Len(A$) - (k - 1), 1))
   Next k
   A$ = ""
   
   Bits2Bytes
   Bytes2Hex    ' HexBytes()    ' Bit slow in VB
   Bytes2Dec    ' DecBytes()    ' Very Slow in VB
   
   ' Get Hex result
   HexString$ = ""
   For k = MaxHexLen To 1 Step -1
      If HexBytes(k) <> 48 Then Exit For
   Next k
   For j = k To 1 Step -1
         HexString$ = HexString$ + Chr$(HexBytes(j))
   Next j
   
   ' Get Dec result
   DecString$ = ""
   For k = MaxDecLen To 1 Step -1
      If DecBytes(k) <> 48 Then Exit For
   Next k
   For j = k To 1 Step -1
         DecString$ = DecString$ + Chr$(DecBytes(j))
   Next j
End Sub

'################################################################################

Private Sub Hex2Bytes(HexString$)
'IN:  HexString$
'OUT: BinBytes(MaxByteLen)
Dim A$
Dim LengthHexStr As Long
Dim k As Long, N As Long
Dim b As Byte
   
   ' Ensure LengthHexStr even
   If (Len(HexString$) And 1) <> 0 Then HexString$ = "0" & HexString$
   LengthHexStr = Len(HexString$)
   b = 48
   FillMemory DecBytes(1), MaxDecLen, b    ' "0" to DecBytes()
   ReDim BinBytes(MaxByteLen)  ' to zero
   ' Transfer 2 nybble values to BinBytes()
   N = 1
   For k = LengthHexStr To 2 Step -2
      A$ = Mid$(HexString$, (k - 1), 2)
      BinBytes(N) = Val("&H" & A$)
      N = N + 1
      If N > MaxByteLen Then Exit For
   Next k
End Sub

Private Sub Bits2Bytes()
'IN:  BinBits(MaxBinLen)
'OUT: BinBytes(MaxByteLen)
Dim i As Long, k As Long
Dim sum As Byte, bit As Byte
Dim Carry As Byte
   ReDim BinBytes(MaxByteLen)  ' to zero
   For i = 1 To MaxByteLen
      sum = 0
      Carry = 0
      For k = 8 To 1 Step -1
         bit = BinBits(k + (i - 1) * 8)
         If (bit And 1) <> 0 Then Carry = 1
         sum = sum * 2 + Carry
         Carry = 0
      Next k
      BinBytes(i) = sum
   Next i
End Sub

Private Sub Bytes2Bits()
'IN:  BinBytes(MaxByteLen)
'OUT: BinBits(MaxBinLen)
Dim i As Long, j As Long, k As Long
Dim one As Byte
Dim b As Byte
   b = 48
   FillMemory BinBits(1), MaxBinLen, b    ' "0" to BinBits()
   
   ' Bytes2Bits.asm .bin
   ' ASM routine inputs
   '  ptrBinBytes
   '  ptrBinBits
   '  MaxByteLen
   '  MaxBinLen
   If Not aVBASM Then
      ptrBinBytes = VarPtr(BinBytes(1))
      ptrBinBits = VarPtr(BinBits(1))
      i = CallWindowProc(ptrMC3, ptrBinBytes, ptrBinBits, MaxByteLen, MaxBinLen)
      Exit Sub
   End If
   
   ' VB routine
   one = 49 ' "1"
   i = 1
   For j = 1 To MaxByteLen
      b = BinBytes(j)
      For k = 0 To 7
         If (b And 1) <> 0 Then BinBits(i) = one
         b = b \ 2
         i = i + 1
      Next k
   Next j
End Sub

Private Sub Bytes2Hex()
' Bit slow in VB
'IN:  BinBytes(MaxByteLen)
'OUT: HexBytes(MaxHexLen)
Dim i As Long, k As Long
Dim b As Byte, lo As Byte, hi As Byte
   b = 48
   FillMemory HexBytes(1), MaxHexLen, b    ' "0" to HexBytes()
   
   ' Bytes2Hex.asm .bin
   ' ASM routine inputs
   '  ptrBinBytes
   '  ptrHexBytes
   '  MaxByteLen
   '  MaxHexLen

   If Not aVBASM Then
      ptrBinBytes = VarPtr(BinBytes(1))
      ptrHexBytes = VarPtr(HexBytes(1))
      i = CallWindowProc(ptrMC2, ptrBinBytes, ptrHexBytes, MaxByteLen, MaxHexLen)
      Exit Sub
   End If
   
   ' VB routine
   i = 1
   For k = 1 To MaxByteLen
      b = BinBytes(k)
      lo = b And &HF
      hi = (b And &HF0) \ 16
      lo = lo + 48
      hi = hi + 48
      If lo > 57 Then lo = lo + 7
      If hi > 57 Then hi = hi + 7
      HexBytes(i) = lo
      HexBytes(i + 1) = hi
      i = i + 2
   Next k
End Sub

Private Sub Dec2Bytes()
' SLOW IN VB!
'IN:  DecBytes(MaxDecLen) from Dec String
'OUT: BinBytes(MaxByteLen)
Dim i As Long, j As Long
Dim byt As Long, lo As Long, hi As Long
Dim ival As Long
   
   ReDim BinBytes(MaxByteLen)  ' to zero

   ' Dec2Bytes.asm .bin
   ' ASM routine inputs
   '  ptrBinBytes
   '  ptrDecBytes
   '  MaxByteLen
   '  MaxDecLen

   If Not aVBASM Then
      ptrBinBytes = VarPtr(BinBytes(1))
      ptrDecBytes = VarPtr(DecBytes(1))
      i = CallWindowProc(ptrMC, ptrBinBytes, ptrDecBytes, MaxByteLen, MaxDecLen)
      Exit Sub
   End If
   ' VB routine
   For i = MaxDecLen - 1 To 1 Step -1
      byt = 0
      For j = 1 To MaxByteLen
         ival = 10 * BinBytes(j)
         lo = ival And &HFF
         hi = (ival And &HFF00) \ 256
         lo = lo + byt
         If lo > 255 Then
            lo = lo - 256
            hi = hi + 1
         End If
         BinBytes(j) = CByte(lo)
         byt = hi
      Next j
      j = 1
      ival = DecBytes(i)
      ival = ival - 48  ' 0 - 9
      ival = BinBytes(j) + ival
      If ival > 255 Then
         ival = ival - 256
         byt = 1
         BinBytes(j) = CByte(ival)
         Do
            j = j + 1
            ival = 1& * BinBytes(j) + byt '1
            If ival > 255 Then
               ival = ival - 256
               byt = 1
            Else
               byt = 0
            End If
            BinBytes(j) = CByte(ival)
         Loop While byt = 1
      Else
         byt = 0
         BinBytes(j) = CByte(ival)
      End If
   Next i
End Sub

Private Sub Bytes2Dec()
' SLOW IN VB!
'IN:  BinBytes(MaxByteLen)
'OUT: DecBytes(MaxDecLen)
Dim i As Long, k As Long
Dim Carry1 As Byte, Carry2 As Byte
Dim bits As Integer, sum As Integer
Dim b As Byte
   b = 48
   ReDim DecBytes(MaxDecLen + 4)
   FillMemory DecBytes(1), MaxDecLen, b    ' "0" to DecBytes()
   
   ' Bytes2Dec.asm .bin
   ' ASM routine inputs
   '  ptrBinBytes
   '  ptrDecBytes
   '  MaxByteLen
   '  MaxBinLen

   If Not aVBASM Then
      ptrBinBytes = VarPtr(BinBytes(1))
      ptrDecBytes = VarPtr(DecBytes(1))
      i = CallWindowProc(ptrMC4, ptrBinBytes, ptrDecBytes, MaxByteLen, MaxBinLen)
      Exit Sub
   End If
   
   ' VB routine
   k = 1
   Do
XX:
      bits = MaxBinLen
      sum = 0
      Do Until bits = 0
         bits = bits - 1
         Carry1 = 0
         ' Shift bits to left with carry
         For i = 1 To MaxByteLen
            ' Check if * 2 will give a carry
            If BinBytes(i) > 127 Then
               Carry2 = 1
               BinBytes(i) = BinBytes(i) - 128
            Else
               Carry2 = 0
            End If
            
            BinBytes(i) = BinBytes(i) * 2 + Carry1  ' Shift << 1 + 1/0
            Carry1 = Carry2
         Next i
         sum = sum * 2 + Carry1    ' Shift << 1 + 1/0
         If sum >= 10 Then
            sum = sum - 10
            BinBytes(1) = BinBytes(1) + 1
         End If
      Loop
      DecBytes(k) = sum + 48  ' Store ASCII digit
      k = k + 1
      ' Check if finished
      For i = MaxByteLen To 1 Step -1
         If BinBytes(i) <> 0 Then GoTo XX ' GoTo used for comparison with ASM
      Next i
      Exit Do
   Loop

End Sub

'###############################################################

Public Function FileExists(FSpec$) As Boolean
  On Error Resume Next
  FileExists = FileLen(FSpec$)
End Function

'###############################################################
' ASM bin file loaders
Public Sub Loadmcode(Infile$)
' Dec2Bytes
' Load machine code from bin file (Option Base 1)
' can be placed in res file when fully debugged
Dim MCSize As Long
Dim fnum As Long
   fnum = FreeFile
   Open Infile$ For Binary As #fnum
   MCSize = LOF(fnum)
   ReDim MCCode(MCSize)
   Get #fnum, , MCCode
   Close #fnum
   ptrMC = VarPtr(MCCode(1))
End Sub

Public Sub Loadmcode2(Infile$)
' Bytes2Hex
' Load machine code from bin file (Option Base 1)
' can be placed in res file when fully debugged
Dim MCSize As Long
Dim fnum As Long
   fnum = FreeFile
   Open Infile$ For Binary As #fnum
   MCSize = LOF(fnum)
   ReDim MCCode2(MCSize)
   Get #fnum, , MCCode2
   Close #fnum
   ptrMC2 = VarPtr(MCCode2(1))
End Sub

Public Sub Loadmcode3(Infile$)
' Bytes2Bits
' Load machine code from bin file (Option Base 1)
' can be placed in res file when fully debugged
Dim MCSize As Long
Dim fnum As Long
   fnum = FreeFile
   Open Infile$ For Binary As #fnum
   MCSize = LOF(fnum)
   ReDim MCCode3(MCSize)
   Get #fnum, , MCCode3
   Close #fnum
   ptrMC3 = VarPtr(MCCode3(1))
End Sub

Public Sub Loadmcode4(Infile$)
' Bytes2Dec
' Load machine code from bin file (Option Base 1)
' can be placed in res file when fully debugged
Dim MCSize As Long
Dim fnum As Long
   fnum = FreeFile
   Open Infile$ For Binary As #fnum
   MCSize = LOF(fnum)
   ReDim MCCode4(MCSize)
   Get #fnum, , MCCode4
   Close #fnum
   ptrMC4 = VarPtr(MCCode4(1))
End Sub

