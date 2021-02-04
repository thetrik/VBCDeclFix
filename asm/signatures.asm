;
; Signatures
; This file produces binary file with signatures descriptions
; by The tick 2021
;
format binary

use32

include "win32wx.inc"

TAG_MODULE = 1
TAG_SIGNATURE = 2
TAG_RANGE_LIST = 3
TAG_MODULE_RANGE = 4
TAG_PATTERN = 5
TAG_MASK = 6
TAG_TARGET = 7

TARGET_ABSOLUTE = 0	 ; Target is offset itself
TARGET_PTR = 1		 ; Target is pointer to pointer
TARGET_RELATIVE = 2	 ; Target is value relative target address (call/jmp DWORD)

macro __START_DECLARE id* {
common
    local __end_decl
    __END_DECL equ __end_decl
    db id
    dd __END_DECL - $ - 4
}

macro __END_DECLARE {
common
    __END_DECL:
    restore __END_DECL
}

; //
; // Create module definition
; //
macro START_MODULE name* {
common
    if defined __INSIDE_MODULE
	err; 'unallowed nested modules'
    end if
    __INSIDE_MODULE equ 1
    __START_DECLARE TAG_MODULE
    db name, 0
}

macro END_MODULE {
    __END_DECLARE
    restore __INSIDE_MODULE
}

; //
; // Create signature definition
; //
macro START_SIGNATURE name* {
common
    if ~defined __INSIDE_MODULE
	err; 'only inside module'
    end if
    if defined __INSIDE_SIGNATURE
	err; 'unallowed nested signatures'
    end if
    __HAS_PATTERN equ
    __HAS_MASK equ
    __INSIDE_SIGNATURE equ 1
    __START_DECLARE TAG_SIGNATURE
    db name, 0
}

macro END_SIGNATURE {
    __END_DECLARE
    restore __INSIDE_SIGNATURE
    restore __HAS_PATTERN
    restore __HAS_MASK
}

; //
; // Create range list
; //
macro START_RANGE_LIST {
    if ~defined __INSIDE_SIGNATURE
	err; 'only inside signature'
    end if
    if defined __INSIDE_RANGE
	err; 'unallowed nested ranges'
    end if
    __INSIDE_RANGE equ 1
    __START_DECLARE TAG_RANGE_LIST
}

macro END_RANGE_LIST {
    __END_DECLARE
    restore __INSIDE_RANGE
}

; //
; // Add range
; //
macro ADD_MODULE_RANGE_TO_LIST section*, start_offset*, end_offset* {
common
    if ~defined __INSIDE_RANGE
	err; 'only inside rangr'
    end if
    __START_DECLARE TAG_MODULE_RANGE
    db section, 0
    dd start_offset
    dd end_offset
    __END_DECLARE
}

; //
; // Define signature pattern
; //
macro SET_SIGNATURE_PATTERN [data*] {
common
    if ~defined __INSIDE_SIGNATURE
	err; 'only inside signature'
    end if
    if __HAS_PATTERN eq 1
	err; 'already defined"
    end if
    __START_DECLARE TAG_PATTERN
    __HAS_PATTERN equ 1
    db data
    __END_DECLARE
}

; //
; // Define signature mask
; //
macro SET_SIGNATURE_MASK [data*] {
common
    if ~defined __INSIDE_SIGNATURE
	err; 'only inside signature'
    end if
    if __HAS_MASK eq 1
	err; 'already defined'
    end if
    __START_DECLARE TAG_MASK
    db data
    __END_DECLARE
}

; //
; // Define signature target. This is offset from start of pattern
; //
macro SET_SIGNATURE_TARGET name*, offset*, type* {
common
    if ~defined __INSIDE_SIGNATURE
	err; 'only inside signature'
    end if
    __START_DECLARE TAG_TARGET
    db type
    db name, 0
    dd offset
    __END_DECLARE
}

START_MODULE "vba6"

    START_SIGNATURE "Bug_Signature"

	; Search within .text section
	START_RANGE_LIST

	    ADD_MODULE_RANGE_TO_LIST ".text", 0, 0

	END_RANGE_LIST

	; Search for this pattern (ECX, EDX can be replaced with other rigisters (see mask))
	; 83E1 03	AND ECX,00000003
	; 8D0C49	LEA ECX,[ECX*2+ECX]
	; 8D0C89	LEA ECX,[ECX*4+ECX]
	; 8D144D XXXXXX LEA EDX,[ECX*2+TBL]		    ; ECX = CL * 30

	SET_SIGNATURE_PATTERN 0x83, 0xE0, 0x03, 0x8D, 0x04, 0x40, 0x8D, 0x04, 0x80, 0x8D, 0x04, 0x45
	SET_SIGNATURE_MASK 0xFF, 0xF8, 0xFF, 0xFF, 0xC7, 0xC0, 0xFF, 0xC7, 0xC0, 0xFF, 0xC7, 0xC7

	; TBL at 0x0c offset
	SET_SIGNATURE_TARGET "BugTablePtr", 0x0C, TARGET_PTR

    END_SIGNATURE

    START_SIGNATURE "VM_DispTables"

	; Search within .ENGINE section
	START_RANGE_LIST

	    ADD_MODULE_RANGE_TO_LIST "ENGINE", 0, 0

	END_RANGE_LIST

	; Search for this pattern
	; FF2485 XXXXXX JMP DWORD PTR DS:[EAX*4+tblByteDisp]
	; 33C0		XOR EAX,EAX
	; 8A06		MOV AL,BYTE PTR DS:[ESI]
	; 46		INC ESI
	; FF2485 XXXXXX JMP DWORD PTR DS:[EAX*4+tblByteDisp_Lead1]
	; 33C0		XOR EAX,EAX
	; 8A06		MOV AL,BYTE PTR DS:[ESI]
	; 46		INC ESI
	; FF2485 XXXXXX JMP DWORD PTR DS:[EAX*4+tblByteDisp_Lead2]
	; 33C0		XOR EAX,EAX
	; 8A06		MOV AL,BYTE PTR DS:[ESI]
	; 46		INC ESI
	; FF2485 XXXXXX JMP DWORD PTR DS:[EAX*4+tblByteDisp_Lead3]
	; 33C0		XOR EAX,EAX
	; 8A06		MOV AL,BYTE PTR DS:[ESI]
	; 46		INC ESI
	; FF2485 XXXXXX JMP DWORD PTR DS:[EAX*4+tblByteDisp_Lead4]
	; 33C0		XOR EAX,EAX
	; 8A06		MOV AL,BYTE PTR DS:[ESI]
	; 46		INC ESI
	; 83F8 46	CMP EAX,46
	; 0F87 XXXXXX	JA XXXXXX
	; FF2485 XXXXXX JMP DWORD PTR DS:[EAX*4+tblByteDisp_Lead5]

	SET_SIGNATURE_PATTERN 0xFF, 0x24, 0x85, 0x00, 0x00, 0x00, 0x00, 0x33, 0xC0, 0x8A, 0x06, 0x46, 0xFF, 0x24, 0x85, 0x00, \
			      0x00, 0x00, 0x00, 0x33, 0xC0, 0x8A, 0x06, 0x46, 0xFF, 0x24, 0x85, 0x00, 0x00, 0x00, 0x00, 0x33, \
			      0xC0, 0x8A, 0x06, 0x46, 0xFF, 0x24, 0x85, 0x00, 0x00, 0x00, 0x00, 0x33, 0xC0, 0x8A, 0x06, 0x46, \
			      0xFF, 0x24, 0x85, 0x00, 0x00, 0x00, 0x00, 0x33, 0xC0, 0x8A, 0x06, 0x46, 0x83, 0xF8, 0x46, 0x0F, \
			      0x87, 0x00, 0x00, 0x00, 0x00, 0xFF, 0x24, 0x85, 0x00, 0x00, 0x00, 0x00

	SET_SIGNATURE_MASK 0xFF, 0xFF, 0xFF, 0x00, 0x00, 0x00, 0x00, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0x00, \
			   0x00, 0x00, 0x00, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0x00, 0x00, 0x00, 0x00, 0xFF, \
			   0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0x00, 0x00, 0x00, 0x00, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, \
			   0xFF, 0xFF, 0xFF, 0x00, 0x00, 0x00, 0x00, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, \
			   0xFF, 0x00, 0x00, 0x00, 0x00, 0xFF, 0xFF, 0xFF, 0x00, 0x00, 0x00, 0x00

	SET_SIGNATURE_TARGET "DispTablePtr", 0x03, TARGET_PTR
	SET_SIGNATURE_TARGET "DispTableLead1Ptr", 0x0F, TARGET_PTR
	SET_SIGNATURE_TARGET "DispTableLead2Ptr", 0x1B, TARGET_PTR
	SET_SIGNATURE_TARGET "DispTableLead3Ptr", 0x27, TARGET_PTR
	SET_SIGNATURE_TARGET "DispTableLead4Ptr", 0x33, TARGET_PTR
	SET_SIGNATURE_TARGET "DispTableLead5Ptr", 0x48, TARGET_PTR

    END_SIGNATURE

    START_SIGNATURE "ImpAdCallHresult"

	; Search within .ENGINE section
	START_RANGE_LIST

	    ADD_MODULE_RANGE_TO_LIST "ENGINE", 0, 0

	END_RANGE_LIST

	; 0FB70E	MOVZX ECX,WORD PTR DS:[ESI]
	; 0FB77E 02	MOVZX EDI,WORD PTR DS:[ESI+2]
	; 83C6 04	ADD ESI,4
	; 03FC		ADD EDI,ESP
	; 8975 D0	MOV DWORD PTR SS:[EBP-30],ESI
	; 8B55 AC	MOV EDX,DWORD PTR SS:[EBP-54]
	; 8B048A	MOV EAX,DWORD PTR DS:[ECX*4+EDX]
	; 0BC0		OR EAX,EAX
	; 75 09 	JNZ l3
	; 51		PUSH ECX
	; 51		PUSH ECX
	; 52		PUSH EDX
	; E8 XXXXXXXX	CALL HrDefSetIndex
	; 59		POP ECX
	; l3: 803D XXXX CMP BYTE PTR DS:[g_EventMonitorsEnabled]
	; 75 33 	JNE l2
	; FFD0		CALL EAX
	; l4: 3BFC	CMP EDI,ESP
	; 0F85 XXXXXXXX JNE BadCCErr
	; 837D C8 00	CMP DWORD PTR SS:[EBP-38],0
	; 74 05 	JE SHORT l1
	; E8 XXXXXXXX	CALL AllocStack_internal
	; l1: 8B75 D0	MOV ESI,DWORD PTR SS:[EBP-30]
	; 0BC0		OR EAX,EAX
	; 78 2F 	JS HrCheck
	; 66:F705 XXXXX TEST WORD PTR DS:[_g_ExceptFlags],0002
	; 75 24 	JNZ HrCheck
	; 33C0		XOR EAX,EAX
	; 8A06		MOV AL,BYTE PTR DS:[ESI]
	; 46		INC ESI
	; FF2485 XXXXXX JMP DWORD PTR DS:[EAX*4+tblByteDisp]
	; l2: 60	PUSHAD
	; 54		PUSH ESP
	; 51		PUSH ECX
	; FF75 B0	PUSH DWORD PTR SS:[EBP-50]
	; E8 XXXXXXXX	CALL VBAEventImportCall
	; 61		POPAD
	; FFD0		CALL EAX
	; 60		PUSHAD
	; 54		PUSH ESP
	; E8 XXXXXXXX	CALL VBAEventProcExit
	; 61		POPAD
	; EB B7 	JMP l4
	; HrCheck: 6A00 PUSH 0
	; 50		PUSH EAX
	; E8 XXXXXXXX	CALL HresultCheck

	SET_SIGNATURE_PATTERN 0x0F, 0xB7, 0x0E, 0x0F, 0xB7, 0x7E, 0x02, 0x83, 0xC6, 0x04, 0x03, 0xFC, 0x89, 0x75, 0xD0, 0x8B, \
			      0x55, 0xAC, 0x8B, 0x04, 0x8A, 0x0B, 0xC0, 0x75, 0x09, 0x51, 0x51, 0x52, 0xE8, 0x00, 0x00, 0x00, \
			      0x00, 0x59, 0x80, 0x3D, 0x00, 0x00, 0x00, 0x00, 0x00, 0x75, 0x33, 0xFF, 0xD0, 0x3B, 0xFC, 0x0F, \
			      0x85, 0x00, 0x00, 0x00, 0x00, 0x83, 0x7D, 0xC8, 0x00, 0x74, 0x05, 0xE8, 0x00, 0x00, 0x00, 0x00, \
			      0x8B, 0x75, 0xD0, 0x0B, 0xC0, 0x78, 0x2F, 0x66, 0xF7, 0x05, 0x00, 0x00, 0x00, 0x00, 0x02, 0x00, \
			      0x75, 0x24, 0x33, 0xC0, 0x8A, 0x06, 0x46, 0xFF, 0x24, 0x85, 0x00, 0x00, 0x00, 0x00, 0x60, 0x54, \
			      0x51, 0xFF, 0x75, 0xB0, 0xE8, 0x00, 0x00, 0x00, 0x00, 0x61, 0xFF, 0xD0, 0x60, 0x54, 0xE8, 0x00, \
			      0x00, 0x00, 0x00, 0x61, 0xEB, 0xB7, 0x6A, 0x00, 0x50, 0xE8, 0x00, 0x00, 0x00, 0x00

	SET_SIGNATURE_MASK 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, \
			   0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0x00, 0x00, 0x00, \
			   0x00, 0xFF, 0xFF, 0xFF, 0x00, 0x00, 0x00, 0x00, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, \
			   0xFF, 0x00, 0x00, 0x00, 0x00, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0x00, 0x00, 0x00, 0x00, \
			   0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0x00, 0x00, 0x00, 0x00, 0xFF, 0xFF, \
			   0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0x00, 0x00, 0x00, 0x00, 0xFF, 0xFF, \
			   0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0x00, 0x00, 0x00, 0x00, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0x00, \
			   0x00, 0x00, 0x00, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0x00, 0x00, 0x00, 0x00

	SET_SIGNATURE_TARGET "HrDefSetIndexCall", 0x1D, TARGET_RELATIVE
	SET_SIGNATURE_TARGET "EventMonitorsEnabledPtr", 0x24, TARGET_PTR
	SET_SIGNATURE_TARGET "AllocStackUnkCall", 0x3C, TARGET_RELATIVE
	SET_SIGNATURE_TARGET "ExceptFlagPtr", 0x4A, TARGET_PTR
	SET_SIGNATURE_TARGET "VbaEventImportCall", 0x65, TARGET_RELATIVE
	SET_SIGNATURE_TARGET "VbaEventProcExitCall", 0x6F, TARGET_RELATIVE
	SET_SIGNATURE_TARGET "HresultCheckPtr", 0x7A, TARGET_RELATIVE

    END_SIGNATURE

    START_SIGNATURE "Declare_Cdecl_Check"

	; Search within .text section
	START_RANGE_LIST

	    ADD_MODULE_RANGE_TO_LIST ".text", 0, 0

	END_RANGE_LIST

	; 8A47 3A	MOV AL,BYTE PTR DS:[EDI+3A]		 ; cdecl = 9; stdcall = 0xc
	; A8 10 	TEST AL,10
	; 0F85 XXXXXXXX JNZ ObjModule
	; 83E0 07	AND EAX,00000007
	; 83F8 01	CMP EAX,1
	; 0F85 XXXXXXXX JNE CompileDeclare
	; E9 XXXXXXXX	JMP CdeclCompileErr31

	SET_SIGNATURE_PATTERN 0x8A, 0x47, 0x3A, 0xA8, 0x10, 0x0F, 0x85, 0x00, 0x00, 0x00, 0x00, 0x83, 0xE0, 0x07, 0x83, 0xF8, \
			      0x01, 0x0F, 0x85, 0x00, 0x00, 0x00, 0x00, 0xE9, 0x00, 0x00, 0x00, 0x00

	SET_SIGNATURE_MASK 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0x00, 0x00, 0x00, 0x00, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, \
			   0xFF, 0xFF, 0xFF, 0x00, 0x00, 0x00, 0x00, 0xFF, 0x00, 0x00, 0x00, 0x00

	SET_SIGNATURE_TARGET "CompileDeclare", 0x13, TARGET_RELATIVE
	SET_SIGNATURE_TARGET "CdeclCompileErr31", 0x18, TARGET_ABSOLUTE

    END_SIGNATURE

END_MODULE
