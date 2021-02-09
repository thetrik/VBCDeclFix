; Thunks for patched procedures
; by The trick 2021

format binary

use32

include "win32wx.inc"
include "macros.inc"

start_data:

dd MODULE_SIGNATURE
dd MODULE_VERSION

CREATE_IMPORT _imp_LastTokenId, \
	      _imp_CurrentTokenStr, \
	      _imp_ExtractCurrentToken, \
	      _imp_ContinueParseId, \
	      _imp_ParseFailed, \
	      _imp_ContinueBSCRBuild, \
	      _imp_ContinueProcDescBuild

align 4

CREATE_EXPORT PCR_ParseThunk, \
	      BSCR_CreateThunk, \
	      ProcDesc_Thunk

align 4

dd end_code - ($ + 4)	; Size of code

; This procedure check CDecl keyword after function ID
PCR_ParseThunk:

mov_rel4 eax, _imp_LastTokenId
mov eax, [eax]
cmp eax, 0x1B				; 1B - CDecl keyword
jne @f
or word [esp + 0x10], 0x4000		; 0x4000 - CDecl flag in PCR
push_rel4 _imp_CurrentTokenStr
invoke_rel _imp_ExtractCurrentToken
mov_rel4 eax, _imp_LastTokenId
mov eax, [eax]
cmp eax, 0xF8				; F8 - left bracket
jne @f
jmp_rel _imp_ContinueParseId

@@:
jmp_rel _imp_ParseFailed

; This procedure updates BSCR tree node
; EAX - pointer to BSCR
; [ebp + 8] - pointer to PCR
BSCR_CreateThunk:

mov dl, [eax + 1]		 ; Accept only object modules
and dl, 7			 ; Mask field

.if dl = 2			 ; 2 - Standard module

    mov edx, [ebp + 8]		 ; EDX <- PCR
    test byte [edx + 5], 0x40	 ; Check CDecl flag

    je @f

    and cl, 0xF8		 ; Set CDecl flag in BSCR
    or cl, 1

.else

    mov edx, [ebp + 8]

.endif

@@:

mov byte [eax + 0x3A], cl	; Update BSCR flags
jmp_rel _imp_ContinueBSCRBuild

; This procedure updates P-code procedure descriptor
; [ebp + 8] - offset within BSCR tree to function BSCR node
ProcDesc_Thunk:

.if dword [ebp + 8] <> -1	; Check offset within BSCR-tree

    mov eax, [esi + 0x2c]
    mov eax, [eax]		; EAX <- BSCR tree
    add eax, dword [ebp + 8]	; BSCR node of function

    mov al, [eax + 0x3a]	; Check calling convention
    and al, 7

    .if al = 1			; if Cdecl

	mov ax, 4
	jmp @f

    .endif

.endif

mov ax, [ebp + 0x18]		; For StdCall get the original value

@@:

mov word [ebp - 0x24],ax

jmp_rel _imp_ContinueProcDescBuild

end_code:

; Relocations
dd RELOCATIONS_COUNT RELOCATIONS