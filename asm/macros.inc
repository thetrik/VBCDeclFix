
;
; Macros for P-code handlers / patch procedures
; by The tick 2021
;

if defined MSVBVM60
else
    MSVBVM60 fix
end if

macro SAVE_INSTRUCTIONS_POINTER {
    mov [ebp - 0x30], esi
}

macro RESTORE_INSTRUCTION_POINTER {
    mov esi, [ebp - 0x30]
}

macro LOAD_FROM_CONST_TABLE {
    mov edx, [ebp - 0x54]
    mov eax, [ecx * 4 + edx]
}

macro GOTO_NEXT_OPCODE {
    xor eax, eax
    mov al, BYTE [esi]
    inc esi
    mov_rel4 ecx, _g_DispTable
    jmp DWORD [eax * 4 + ecx]
}

macro LOAD_CALL_PARAMETERS {

    movzx ecx, WORD [esi + tFuncDesc.wFuncIndex]
    movzx edi, WORD [esi + tFuncDesc.wArgsSize]

    add esi, sizeof.tFuncDesc ; Next opcode
    add edi, esp	      ; Save valid stack value

}

macro CHECK_LOAD_FUNCTION {

    local IS_LOADED

    .if ~eax

	; HrDefSetIndex
	push ecx
	invoke_rel __imp_HrDefSetIndex, edx, ecx
	pop ecx

    .endif

}

RELOCATIONS equ
RELOCATIONS_COUNT = 0

macro invoke_rel proc*, [args] {
common

    local lbl

    invoke proc, args

    lbl:

    RELOCATIONS equ RELOCATIONS, lbl - 4
    RELOCATIONS_COUNT = RELOCATIONS_COUNT + 1

}

macro jmp_rel proc* {
    local lbl

    jmp [proc]

    lbl:

    RELOCATIONS equ RELOCATIONS, lbl - 4
    RELOCATIONS_COUNT = RELOCATIONS_COUNT + 1

}

macro mov_rel4 reg*, ptr* {
    local lbl

    mov reg, [ptr]

    lbl:

    RELOCATIONS equ RELOCATIONS, lbl - 4
    RELOCATIONS_COUNT = RELOCATIONS_COUNT + 1

}

macro push_rel4 ptr* {
    local lbl

    push [ptr]

    lbl:

    RELOCATIONS equ RELOCATIONS, lbl - 4
    RELOCATIONS_COUNT = RELOCATIONS_COUNT + 1

}

macro CALL_FUNCTION_TRACE {

    mov_rel4 ecx, _g_pEventMonitorsEnabled

    .if BYTE [ecx]

	pushad
	invoke_rel __imp_VBAEventImportCall, dword [ebp - 0x50], ecx, esp
	popad
	call eax
	pushad
	invoke_rel __imp_VBAEventProcExit, esp
	popad

    .else

	call eax

    .endif

    movzx ecx, WORD [esi - 2]
    add esp, ecx

    .if edi <> esp
	invoke_rel __imp_EbRaiseExceptionCode, 0x31
    .endif

if MSVBVM60 eq

    .if DWORD [ebp - 0x38]
	invoke_rel __imp_AllocStackUnk
    .endif

end if

}

macro TEST_EXCEPT_FLAGS {

if MSVBVM60 eq TRUE
    mov edx, [ebp - 0x44]
    test WORD [edx + 0x76], 2
else
    mov_rel4 ecx, _g_ExceptFlags
    test WORD [ecx], 2
end if
    jnz HrCheck

}

; Create export
macro CREATE_EXPORT [pSym] {
    _counter = 0

    common

	dd _number_of_exports

    forward

	dd pSym
	_counter = _counter + 1

    forward

	db `pSym , 0

    common

	_number_of_exports = _counter

}

; Create imported addresses table
macro CREATE_IMPORT [pSym] {
    _counter = 0

    common

	dd _number_of_imports

    forward

	pSym dd 0
	_counter = _counter + 1

    forward

	db `pSym , 0

    common

	_number_of_imports = _counter

}

; Macro to define an opcode handler
; 1st param - opcode value
; 2nd param - handler
macro DEFINE_OPCODES [op*, hdlr*] {
    _counter = 0

    common

	dd _number_of_handlers

    forward

	if op >= 0 & op <= 255
	    db op
	    dd hdlr
	else
	    ; Swap
	    db ((op and 0x0ff00) shr 8), (op and 0x0ff)
	    dd hdlr
	end if

	_counter = _counter + 1

    common

	_number_of_handlers = _counter

}