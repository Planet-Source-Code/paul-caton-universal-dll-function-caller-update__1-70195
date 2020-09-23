
;******************************************************************************************
;
; Wrap a .bas module callback from a CDECL function
;
; v1.00 20071201 Original cut.......................................................... prc
; v1.01 20080203 Storing the return address as part of the push instruction............ prc
;******************************************************************************************

use32						;32bit

	call	L1				;Call the next instruction
L1:	pop	eax				;Pop the return address into eax (eax = L1)

	pop	dword [eax+(L2-L1)]		;Pop the calling function's return address into the immediate value 'push' instruction at L2

	db	0E8h				;Op-code for a relative address call
	dd	55555555h			;EIP-relative address of target vb module function, patched by cCallFunc2.CallbackCdecl

	sub	esp, 55h			;Adjust the stack... patched by cCallFunc2.CallbackCdecl

	db	068h				;Op-code for an immediate value push
L2:	dd	55555555h			;Return address, patched by the instruction after L2
	ret					;Return to caller
	nop
