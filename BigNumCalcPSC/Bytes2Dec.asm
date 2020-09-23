; Bytes2Dec.asm  FASM  by  Robert Rayment Apr 2006

; Using FlatAssembler  <www.flatassembler.net>

; res=CallWindowProc(ptrMC, ptrBinBytes, ptrDecBytes, MaxByteLen, MaxBinLen)
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
MaxBinLen equ   [ebp-16]

    push ebp
    mov ebp,esp
    sub esp,16
    push edi
    push esi
    push ebx
    
    movab ptrBinBytes,[ebp+8]
    movab ptrDecBytes,[ebp+12]
    movab MaxByteLen, [ebp+16]
    movab MaxBinLen,  [ebp+20]

    call Bytes2Dec
    
GETOUT:
    pop ebx
    pop esi
    pop edi
    mov esp,ebp
    pop ebp
    ret 16

; #########################################
Bytes2Dec:
        ;CONVERT BINARY BYTES TO ASCII DECIMAL
        cld
        mov edi,ptrDecBytes
B2DStart:
        mov ebx,MaxByteLen 
        mov edx,MaxBinLen
        inc edx                 ; bits = MaxBinLen+1
        mov eax,0               ; sum = 0
scan1:                         
        dec edx
        cmp edx,0               ; cf=0
        je storeB               ; 1 fullbit scan done, store digit
        clc
        mov esi,ptrBinBytes     ; ptr BinBytes(1)
        mov ecx,ebx             ; For i = 1 To MaxByteLen
shm:
        rcl byte[esi],1         ;shift bit all way up BinBytes
        inc esi
        dec ecx
        jnz shm

        rcl ax,1                ;build remainder
        cmp ax,10                       
        jb scan1                ;still< 10

        sub esi,ebx             ; ptr BinBytes(1)
        sub ax,10               ; subtract 10 &
        inc byte[esi]           ; bump BinBytes(1)
        jmp scan1               ; Loop
storeB:
        ; DecBytes(k) = sum + 48: k = k + 1
        or al,48                ; make ascii
        stosb                   ; al->[edi] edi+1
        ; Check quotient, when zero done
        mov esi,ptrBinBytes     ; ptr BinBytes(1)
        mov ecx,ebx             ; MaxByteLen
testquo:
        cmp byte[esi],0
        jne B2DStart
        inc esi
        dec ecx
        jnz testquo
ret

