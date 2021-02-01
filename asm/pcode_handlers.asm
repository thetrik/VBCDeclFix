; P-code handlers for CDECL functions
; By The trick 2021
; FASM compiler used

;MODULE_SIGNATURE = 0
;MODULE_VERSION = 0

format binary

use32

include "win32wx.inc"
include "macros.inc"

struct tFuncDesc
    wFuncIndex	dw ?
    wArgsSize	dw ?
ends

start_data:

dd MODULE_SIGNATURE
dd MODULE_VERSION

; List of opcodes handlers
;DEFINE_OPCODES 0x09, ImpAdCallHresult_Cdecl, \        ; need fix
DEFINE_OPCODES 0xFE39, ImpAdCallFPR4_Cdecl, \
	       0xFE3A, ImpAdCallFPR8_Cdecl, \
	       0xFE35, ImpAdCallCy_Cdecl, \
	       0xFE30, ImpAdCallUI1_Cdecl, \
	       0xFE32, ImpAdCallI4_Cdecl, \
	       0xFE38, ImpAdCallUnk_Cdecl, \
	       0xFE37, ImpAdCallStr_Cdecl, \
	       0xFE31, ImpAdCallI2_Cdecl, \
	       0xFE3C, ImpAdCallVoid_Cdecl

align 4

; List of symbols
CREATE_IMPORT __imp_HrDefSetIndex, \
	      __imp_VBAEventImportCall, \
	      __imp_VBAEventProcExit, \
	      _g_pEventMonitorsEnabled, \
	      __imp_EbRaiseExceptionCode, \
	      __imp_AllocStackUnk, \
	      __imp_HresultCheck, \
	      _g_ExceptFlags, \
	      _g_DispTable


; Code
dd end_code - start_code

start_code:

HrCheck:
    invoke_rel __imp_HresultCheck, eax, 0

; An API function which returns HRESULT.
ImpAdCallHresult_Cdecl:

    LOAD_CALL_PARAMETERS
    SAVE_INSTRUCTIONS_POINTER
    LOAD_FROM_CONST_TABLE
    CHECK_LOAD_FUNCTION
    CALL_FUNCTION_TRACE
    RESTORE_INSTRUCTION_POINTER

    .if signed eax < 0
	jmp HrCheck
    .endif

    TEST_EXCEPT_FLAGS
    GOTO_NEXT_OPCODE

; An API function which returns a float/double value
ImpAdCallFPR4_Cdecl:
ImpAdCallFPR8_Cdecl:

    LOAD_CALL_PARAMETERS
    SAVE_INSTRUCTIONS_POINTER
    LOAD_FROM_CONST_TABLE
    CHECK_LOAD_FUNCTION
    CALL_FUNCTION_TRACE
    RESTORE_INSTRUCTION_POINTER
    GOTO_NEXT_OPCODE

; An API function which returns a currency value
ImpAdCallCy_Cdecl:

    LOAD_CALL_PARAMETERS
    SAVE_INSTRUCTIONS_POINTER
    LOAD_FROM_CONST_TABLE
    CHECK_LOAD_FUNCTION
    CALL_FUNCTION_TRACE
    RESTORE_INSTRUCTION_POINTER

    push edx
    push eax

    GOTO_NEXT_OPCODE

; An API function which returns a byte value
ImpAdCallUI1_Cdecl:
    LOAD_CALL_PARAMETERS
    SAVE_INSTRUCTIONS_POINTER
    LOAD_FROM_CONST_TABLE
    CHECK_LOAD_FUNCTION
    CALL_FUNCTION_TRACE
    RESTORE_INSTRUCTION_POINTER

    xor ah, ah
    push eax

    GOTO_NEXT_OPCODE

; An API function which returns a long value
ImpAdCallI4_Cdecl:
ImpAdCallUnk_Cdecl:
ImpAdCallStr_Cdecl:
ImpAdCallI2_Cdecl:
    LOAD_CALL_PARAMETERS
    SAVE_INSTRUCTIONS_POINTER
    LOAD_FROM_CONST_TABLE
    CHECK_LOAD_FUNCTION
    CALL_FUNCTION_TRACE
    RESTORE_INSTRUCTION_POINTER

    push eax

    GOTO_NEXT_OPCODE

ImpAdCallVoid_Cdecl:
    LOAD_CALL_PARAMETERS
    SAVE_INSTRUCTIONS_POINTER
    LOAD_FROM_CONST_TABLE
    CHECK_LOAD_FUNCTION
    CALL_FUNCTION_TRACE
    RESTORE_INSTRUCTION_POINTER
    GOTO_NEXT_OPCODE

end_code:

; Relocations
dd RELOCATIONS_COUNT RELOCATIONS