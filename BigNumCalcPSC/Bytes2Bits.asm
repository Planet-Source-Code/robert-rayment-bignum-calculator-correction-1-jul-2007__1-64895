; Bytes2Bits.asm  FASM  by  Robert Rayment Apr 2006

; Using FlatAssembler  <www.flatassembler.net>

; res=CallWindowProc(ptrMC, ptrBinBytes, ptrBinBits, MaxByteLen, MaxBinLen)
;                           8            12           16          20

macro movab %1,%2
 {
        push dword %2
        pop dword %1
 }
format binary
Use32

ptrBinBytes equ [ebp-4]
ptrBinBits equ  [ebp-8]
MaxByteLen equ  [ebp-12]
MaxBinLen equ   [ebp-16]

    push ebp
    mov ebp,esp
    sub esp,16
    push edi
    push esi
    push ebx
    
    movab ptrBinBytes,[ebp+8]
    movab ptrBinBits, [ebp+12]
    movab MaxByteLen, [ebp+16]
    movab MaxBinLen,  [ebp+20]

    call Bytes2Bits
    
GETOUT:
pop ebx
pop esi
pop edi
mov esp,ebp
pop ebp
ret 16
;##########################################################
Bytes2Bits:
    ; CONVERT BINARY BYTES TO BINARY BITS
    mov esi,ptrBinBytes     ; pointer to BinBytes(1)
    mov edi,ptrBinBits      ; pointer to BinBits(1)
    
    mov ah,49          ; "1"
    mov ecx,MaxByteLen ; j = 1 To MaxByteLen

nxbinbyt:
    mov al,[esi]       ; pick up BinByte
    mov ebx,8          ; k = 0 To 7
getbits:
    shr al,1
    jnc nxedi
    mov [edi],ah       ; "1"
nxedi:
    inc edi
    dec ebx
    jnz getbits        ; Next k

    inc esi
    dec ecx
    jnz  nxbinbyt      ; Next j
ret
;=========================================================
