
/* 
** PlugIn.cpp
***************************************************************************************************/

#define WINVER         0x0501
#define _WIN32_WINNT   WINVER
#define _WIN32_WINDOWS WINVER

#include <windows.h>
#include <stdio.h>

typedef LONG (__stdcall * _Callback)(LONG nValue);

CHAR   szPath[MAX_PATH];
CHAR * szFileName;

/*
** dll entry point
***************************************************************************************************/
extern "C" BOOL WINAPI _DllMainCRTStartup(HMODULE  hModule, DWORD dwReason, LPVOID  lpreserved)
{
	if (dwReason == DLL_PROCESS_ATTACH)
	{
      //Get the full path of this module
      szFileName = szPath + GetModuleFileNameA(hModule, szPath, MAX_PATH);

      while (szFileName > szPath)
         if (*(--szFileName) == '\\') 
            break;

      *(szFileName++) = 0;
	}

	return TRUE;
}

/*
** MyFunc1, dll ordinal #1
***************************************************************************************************/
LONG __stdcall MyFunc1(LONG nParam1, LONG nParam2, HWND hWndParent)
{
   CHAR szBuf[512];
   CHAR szCap[256];

   wsprintfA(szBuf, "Plugin path:\t%s\nPlugin file:\t%s\nFunction name:\tMyFunc1\nParameter #1:\t%d\nParameter #2:\t%d\n\nClick any button to continue.", szPath, szFileName, nParam1, nParam2);
   wsprintfA(szCap, "%s - MyFunc1", szFileName);

   return MessageBoxA(hWndParent, szBuf, szCap, MB_ABORTRETRYIGNORE);
}

/*
** MyFunc2, dll ordinal #2
***************************************************************************************************/
LONG __stdcall MyFunc2(LONG nParam1, LONG nParam2, HWND hWndParent)
{
   CHAR szBuf[512];
   CHAR szCap[256];

   wsprintfA(szBuf, "Plugin path:\t%s\nPlugin file:\t%s\nFunction name:\tMyFunc2\nParameter #1:\t%d\nParameter #2:\t%d\n\nClick any button to continue.", szPath, szFileName, nParam1, nParam2);
   wsprintfA(szCap, "%s - MyFunc2", szFileName);

   return MessageBoxA(hWndParent, szBuf, szCap, MB_ABORTRETRYIGNORE);
}

/*
** MyCallback, dll ordinal #3
***************************************************************************************************/
LONG __stdcall MyCallback(_Callback fnCallback)
{
   for (int i=1; i<=100; i++) //Plugin_1.plugin
// for (int i=100; i>0; i--)  //Plugin_2.plugin
   {
      Sleep(100);

      if (fnCallback(i) == 0)
         break;
   }

   return 1;
}

/*
** MyShape, dll ordinal #4
***************************************************************************************************/
LONG __stdcall MyShape(HWND hWnd, HDC hdc)
{
   RECT   rc;
   HBRUSH hbrOld, hbrNew;
   
   GetClientRect(hWnd, &rc);

   rc.left = 8;
   rc.top = 8;
   rc.right = rc.right - 8;
   rc.bottom = rc.bottom - 46;

   hbrNew = CreateSolidBrush(RGB(0, 128, 0));      //Plugin_1.plugin
// hbrNew = CreateSolidBrush(RGB(128, 0, 0));      //Plugin_2.plugin

   hbrOld = (HBRUSH) SelectObject(hdc, hbrNew);

   Ellipse(hdc, rc.left, rc.top, rc.right, rc.bottom);           //Plugin_1.plugin
// RoundRect(hdc, rc.left, rc.top, rc.right, rc.bottom, 8, 8);   //Plugin_2.plugin

   SelectObject(hdc, hbrOld);
   DeleteObject(hbrNew);

   return 0;
}

/*
** Fast1, dll ordinal #5
***************************************************************************************************/
LONG __fastcall Fast1(LONG nParam1)
{
   CHAR szBuf[512];
   CHAR szCap[256];

   wsprintfA(szBuf, "Plugin path:\t%s\nPlugin file:\t%s\nFunction name:\tFast1\nParameter #1:\t%d\n\n\nClick any button to continue.", szPath, szFileName, nParam1);
   wsprintfA(szCap, "%s - Fast1", szFileName);

   return MessageBoxA(0, szBuf, szCap, MB_ABORTRETRYIGNORE);
}

/*
** Fast2, dll ordinal #6
***************************************************************************************************/
LONG __fastcall Fast2(LONG nParam1, LONG nParam2)
{
   CHAR szBuf[512];
   CHAR szCap[256];

   wsprintfA(szBuf, "Plugin path:\t%s\nPlugin file:\t%s\nFunction name:\tFast2\nParameter #1:\t%d\nParameter #2:\t%d\n\nClick any button to continue.", szPath, szFileName, nParam1, nParam2);
   wsprintfA(szCap, "%s - Fast2", szFileName);

   return MessageBoxA(0, szBuf, szCap, MB_ABORTRETRYIGNORE);
}

/*
** Fast3, dll ordinal #7
***************************************************************************************************/
LONG __fastcall Fast3(LONG nParam1, LONG nParam2, LONG nParam3)
{
   CHAR szBuf[512];
   CHAR szCap[256];

   wsprintfA(szBuf, "Plugin path:\t%s\nPlugin file:\t%s\nFunction name:\tFast3\nParameter #1:\t%d\nParameter #2:\t%d\nParameter #3:\t%d\n\nClick any button to continue.", szPath, szFileName, nParam1, nParam2, nParam3);
   wsprintfA(szCap, "%s - Fast3", szFileName);

   return MessageBoxA(0, szBuf, szCap, MB_ABORTRETRYIGNORE);
}