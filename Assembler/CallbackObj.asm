
;******************************************************************************************
;
; Wrap an object (.cls/.frm/.ctl) callback from a cdecl or stdcall function
;
; v1.00 20071201 Original cut.......................................................... prc
; v1.01 20080203 Storing the return value as part of mov instruction................... prc
;******************************************************************************************

use32						;32bit

	mov	eax, esp			;Copy the stack pointer

	call	L1				;Call the next instruction
L1:	pop	edx				;Pop the return address into edx (edx = L1)

	add	edx, (L4-L1)			;Add the offeset to L4 (edx = L4)
	push	edx				;Push the return value location

	mov	ecx, 55h			;Number of parameters into ecx, patched by cCallFunc2.CallbackObj
	jecxz	L3				;If ecx = 0 (no parameters) then jump over the parameter push loop

L2:	push	dword [eax + ecx * 4]		;Push the parameter
	loop	L2				;Next parameter

L3:	push	55555555h			;Push the object address, patched by cCallFunc2.CallbackObj
	db	0E8h				;Op-code for call relative
	dd	55555555h			;EIP relative address of target object function, patched by cCallFunc2.CallbackObj

	db	0B8h				;Op-code for move eax immediate value
L4:	dd	55555555h			;immediate return value
	ret	55h				;Return to caller, stack adjustment patched by cCallFunc2.CallbackObj
	nop
	nop
