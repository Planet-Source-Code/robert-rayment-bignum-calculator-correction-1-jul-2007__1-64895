; Bytes2Hex.asm  FASM  by  Robert Rayment Apr 2006

; Using FlatAssembler  <www.flatassembler.net>

; res=CallWindowProc(ptrMC, ptrBinBytes, ptrHexBytes, MaxByteLen, MaxHexLen)
;                           8            12           16          20

macro movab %1,%2
 {
        push dword %2
        pop dword %1
 }
format binary
Use32

ptrBinBytes equ [ebp-4]
ptrHexBytes equ [ebp-8]
MaxByteLen equ  [ebp-12]
MaxHexLen equ   [ebp-16]

    push ebp
    mov ebp,esp
    sub esp,16
    push edi
    push esi
    push ebx
    
    movab ptrBinBytes,[ebp+8]
    movab ptrHexBytes,[ebp+12]
    movab MaxByteLen, [ebp+16]
    movab MaxHexLen,  [ebp+20]

    call Bytes2Hex
    
GETOUT:
pop ebx
pop esi
pop edi
mov esp,ebp
pop ebp
ret 16
;##########################################################
Bytes2Hex:
    ; CONVERT BINARY BYTES TO HEX BYTES
    mov esi,ptrBinBytes     ; pointer to BinBytes(1)
    mov edi,ptrHexBytes     ; pointer to HexBytes(1)
    mov bx,03030h           ; 48 48
    mov ecx,MaxByteLen      ; k = 1 To MaxByteLen
nxbyt:
    mov al,[esi]    ; pick up BinByte
    aam 16          ; ah=high nyb al=lo nyb
                    ; 48-57(for 0-9)58-63 (for 10-15) 
                    ; 48-57(for 0-9)58-63 (for 10-15) 
    add ax,bx       ; ah+48, al+48 
    cmp al,57
    jbe testah
    add al,7        ; 58+7=65
testah:
    cmp ah,57
    jbe storehex
    add ah,7        ; 58+7=65
storehex:
    mov [edi],ax
    inc edi
    inc edi
    inc esi
    dec ecx
    jnz  nxbyt      ; Next k
ret
;=========================================================
