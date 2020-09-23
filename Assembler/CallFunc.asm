;******************************************************************************************
; cdecl, stdcall and fastcall function caller
;
; The stack on entry:
; [esp+12]  Address of the return value location
; [esp+ 8]  Address of the parameter block
; [esp+ 4]  Me object pointer
; [esp   ]  Return address
;
; v1.00 20071201 Original cut.......................................................... prc
; v1.01 20072309 Added fastcall support................................................ prc
; v1.02 20080212 Added support for additional return types............................. prc
; v1.03 20080303 Redundant fastcall jump removed....................................... prc
; v1.04 20080311 Return stack adjustment fixed......................................... prc
;******************************************************************************************

use32						 ;32bit
	call	L1				 ;Call the next instruction
L1:	pop	eax				 ;Pop the return address into eax (eax = L1)
	mov    [eax+(L4-L1)], esp		 ;Save the stack pointer to L4

	mov	eax, dword [esp+8]		 ;Address of the parameter block into eax
	mov	ecx, [eax]			 ;Number of parameters into ecx
	jecxz	L3				 ;If ecx = 0 (no parameters) then jump over the parameter push loop

	jmp	fastcall			 ;Patched at runtime to jmp or nop

L2:	push	dword [eax+ecx*4]		 ;Push the parameter
	loop	L2				 ;Next parameter

L3:	db	0E8h				 ;Call eip relative
	dd	55555555h			 ;EIP-relative address of the target function, patched by cCallFunc2.CallFunc/CallPointer

	db	0BCh				 ;mov esp, immediate value
L4:	dd	55555555h			 ;Immediate value, patched by the code after 'L1' -- restore the entry value of esp

	mov	ecx, [esp+12]			 ;Get the address of the return value location
	jmp	return				 ;Jump patched at runtime for appropriate return type

case_i08:
	mov	[ecx], al			 ;Write the int8 return value to the return value location
	jmp	return

case_i16:
	mov	[ecx], ax			 ;Write the int32 return value to the return value location
	jmp	 return

case_i32:
	mov	[ecx], eax			 ;Write the int32 return value to the return value location
	jmp	 return

case_i64:
	mov	[ecx], eax			 ;Write the int64 return value to the return value location
	mov	[ecx+4], edx
	jmp	 return

case_sng:
	fstp	dword [ecx]			 ;Write the float return value to the return value location
	jmp	 return

case_dbl:
	fstp	qword [ecx]			 ;Write the double return value to the return value location

return:
	xor	eax, eax			 ;Clear eax, indicates to VB that all is well
	ret	16				 ;Return

L5:	push	dword [eax+ecx*4]		 ;Push the parameter
	dec	ecx				 ;Decrement the parameter index
fastcall:
	cmp	ecx, 2
	jg	L5				 ;Next parameter

	mov	ecx, [eax+4]			 ;The first parameter into ecx
	mov	edx, [eax+8]			 ;The second parameter into edx
	jmp	L3				 ;Call the function
