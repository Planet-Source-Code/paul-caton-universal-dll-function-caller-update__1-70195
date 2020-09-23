#include <windows.h>

extern "C" { int _fltused; }; 

extern "C" BOOL WINAPI _DllMainCRTStartup(HMODULE  hModule, DWORD dwReason, LPVOID  lpreserved)
{
   return TRUE;
}

__int8  GetByte   () { return 0x12;                 }
__int16 GetInteger() { return 0x1234;               }
__int32 GetLong   () { return 0x12345678;           }
__int64 GetInt64  () { return 1234567890123;        }
void *  GetPointer() { return "Hello World";        }
float   GetSingle () { return 1.234567890123456789; }
double  GetDouble () { return 1.234567890123456789; }

double GetDblInc(double d) 
{ 
   return d + 1.00;  
}

void Subroutine(__int32 nValue)
{
   CHAR szText[80];

   wsprintfA(szText, "Parameter nValue: 0x%X", nValue);
   MessageBoxA(0, szText,"Subroutine test", MB_OK);
}