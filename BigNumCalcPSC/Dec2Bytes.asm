; Dec2Bytes.asm  FASM  by  Robert Rayment Apr 2006

; Using FlatAssembler  <www.flatassembler.net>

; res=CallWindowProc(ptrMC, ptrBinBytes, ptrDecBytes, MaxByteLen, MaxDecLen)
;                           8            12           16          20

macro movab %1,%2
 {
        push dword %2
        pop dword %1
 }
format binary
Use32

ptrBinBytes equ [ebp-4]
ptrDecBytes equ [ebp-8]
MaxByteLen equ  [ebp-12]
MaxDecLen equ   [ebp-16]

    push ebp
    mov ebp,esp
    sub esp,16
    push edi
    push esi
    push ebx
    
    movab ptrBinBytes,[ebp+8]
    movab ptrDecBytes,[ebp+12]
    movab MaxByteLen, [ebp+16]
    movab MaxDecLen,  [ebp+20]

    call Dec2Bytes
    
GETOUT:
    pop ebx
    pop esi
    pop edi
    mov esp,ebp
    pop ebp
    ret 16

; #########################################
Dec2Bytes:
        ; CONVERT ASCII DECIMAL TO BINARY BYTES
        mov edi,ptrBinBytes     ; pointer to BinBytes(0)
        mov edx,edi
        mov esi,ptrDecBytes     ; pointer to DecBytes(0) -(19)
        add esi,MaxDecLen       ; -> DecBytes(19) I = MaxDecLen to 1 Step -1
        dec esi
        cld                     ; ensure inc edi
        
        mov bh,10
NXScan:
        mov ecx,MaxByteLen      ; j = 1 To MaxByteLen
        xor bl,bl               ; bl = 0 = byt
MUL10:
        mov al,[edi]            ; pick up binbyte & x10
        mul bh                  ; al*10 -> ax, (ah,al)
        add al,bl               ; carry if al > 255
        adc ah,0                ; adds any carry
        stosb                   ; al->[edi]  edi+1
        mov bl,ah
        dec ecx
        jnz MUL10
        
        ; add next decimal
        mov edi,edx             ; pointer to BinBytes(1)
        mov al,[esi]            ; pick up next ascii dec
        and al,0Fh              ; 0-9
        add [edi],al            ; add to BinBytes array
        jnc testend

        ; send carry back up thru BinBytes
        ; determined by input number magnitude else error
addc:
        inc edi
        adc [edi],byte 0
        jc addc
        mov edi,edx             ; pointer to BinBytes(1)
testend:
        dec esi
        cmp esi,ptrDecBytes     ; pointer to DecBytes(1)
        jae NXScan

ret
;-----------------------------------------
