VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCallFunc2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'**********************************************************************************
'** cCallFunc2.cls - cCallFunc with added fastcall support, call by address and
'**                  additional return types
'**
'** Universal dll function/sub calling class
'**   cdecl/stdcall/fastcall calling convention
'**   Call by ordinal, name or address
'**   Module (.bas) callbacks for cdecl.
'**   Object (.cls/.frm/.ctl) callbacks for cdecl/stdcall
'**   Support for multiple callbacks.
'**   Support for multiple cCallFunc2 instances
'**   Support unicode path\module names
'**
'** If you wish to do crazy stuff like CallFunc with callbacks inside a callback
'** then the best solution is to make a copy of the class, eg cCallFunc1.cls, and
'** use an instance of that where needed.
'**
'** Calling conventions:
'**   stdcall:  parameters right to left, called routine adjusts the stack
'**   cdecl:    parameters right to left, caller adjusts the stack
'**   fastcall: first parameter, if present, in the ecx register
'**               second parameter, if present, in the edx register
'**               any other parameters are pushed to the stack
'**               called routine adjusts the stack
'**               N.B. fastcall isn't standardised, differing conventions exist.
'**               This class supports the Microsoft/GCC implementation.
'**
'** paul_caton@hotmail.com
'**
'** 20031029 First cut....................................................... v1.00
'** 20071129 Now using virtual memory to fix a DEP issue..................... v1.01
'** 20071130 Hacked from cCDECL, now supports stdcall and ordinals........... v1.02
'** 20071201 Added support for callback objects.............................. v1.03
'** 20071202 Unicode support for paths\modules where available............... v1.04
'** 20071213 Forked from cCallFunc.cls
'**          Added support for fastcall calling convention
'**          Added CallPointer
'**          Changed the interface to be more property like.................. v1.05
'** 20080212 Support Byte, Integer, Long, Single and Double return types..... v1.06
'** 20080311 Added IsValidDll and IsValidMethod
'**          Parameter block made global
'**          Eliminated MAX_ARG, VB has a limit of 60 parameters
'**          Various optimizations........................................... v1.07
'**********************************************************************************

Option Explicit

'API declarations
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Function GetProcByName Lib "kernel32" Alias "GetProcAddress" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetProcByOrdinal Lib "kernel32" Alias "GetProcAddress" (ByVal hModule As Long, ByVal nOrdinal As Long) As Long
Private Declare Function IsBadCodePtr Lib "kernel32" (ByVal lpfn As Long) As Long
Private Declare Function IsWindowUnicode Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function LoadLibraryA Lib "kernel32" (ByVal lpLibFileName As String) As Long
Private Declare Function LoadLibraryW Lib "kernel32" (ByVal lpLibFileName As Long) As Long
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long

Private Declare Sub GetMem1 Lib "msvbvm60" (ByVal Addr As Long, RetVal As Byte)
Private Declare Sub GetMem4 Lib "msvbvm60" (ByVal Addr As Long, RetVal As Long)
Private Declare Sub PutMem1 Lib "msvbvm60" (ByVal Addr As Long, ByVal NewVal As Byte)
Private Declare Sub PutMem2 Lib "msvbvm60" (ByVal Addr As Long, ByVal NewVal As Integer)
Private Declare Sub PutMem4 Lib "msvbvm60" (ByVal Addr As Long, ByVal NewVal As Long)
Private Declare Sub PutMem8 Lib "msvbvm60" (ByVal Addr As Long, ByVal NewVal As Currency)
Private Declare Sub SetLastError Lib "kernel32" (ByVal dwErrCode As Long)

Public Enum eObjType                                                          'Object type for CallbackObj... also incorporates vTable offsets
  objCls = &H1C                                                               'Class object callback
  objFrm = &H6F8                                                              'Form object callback
  objCtl = &H7A4                                                              'UserControl object callback
End Enum                                                                      '
                                                                              '
Public Enum eReturnType                                                       'CallFunc/CallPointer return types... also incorporates return type jump values
  retByte = &H0                                                               'Return Byte
  retInteger = &H4                                                            'Return Integer
  retLong = &H9                                                               'Return Long
  retInt64 = &HD                                                              'Return 64 bit value eg. Currency
  retSingle = &H14                                                            'Return Single
  retDouble = &H18                                                            'Return Double
  retSub = &H1A                                                               'No return value
End Enum                                                                      '
                                                                              '
Private Const SRC           As String = "cCallFunc2."                         'Error source
                                                                              '
Private Type tParamBlock                                                      'Parameter block type
  ParamCount                As Long                                           'Number of parameters
  Params(0 To 59)           As Long                                           'Array of parameters
End Type                                                                      '
                                                                              '
Private m_FastCall          As Boolean                                        'FastCall private property value
Private m_LastError         As Long                                           'LastError private property value
                                                                              
Private bUnicode            As Boolean                                        'Unicode flag                                                                              '
Private vCode               As Long                                           'Pointer to the machine-code thunks
Private vTable              As Long                                           'Class vTable address
Private nAddrPb             As Long                                           'Address of the parameter block
Private hModule             As Long                                           'Current/last-used dll handle
Private strLastDLL          As String                                         'Current/last-used dll name
Private strLastFunc         As String                                         'Current/last-used function/sub name
Private pb                  As tParamBlock                                    'Parameter block

'CallFunc:
'
' strDLL      - Name of the DLL
' RetType     - Function return type
' strFunc     - Name of the function or it's ordinal value preceded by a '#' eg. "#2"
' ParamLongs  - Any number [or none] of parameters As Long.
'                 To pass the address (ByRef) of a string use StrPtr, eg. StrPtr(strPath)
'                 To pass the address (ByRef) of a variable or UDT use VarPtr, eg. VarPtr(i)
Public Function CallFunc(ByRef strDll As String, _
                         ByVal RetType As eReturnType, _
                         ByRef strFunc As String, _
                    ParamArray ParamLongs() As Variant) As Variant            '
  Dim bNewDll As Boolean                                                      'New dll flag
                                                                              '
  If StrComp(strDll, strLastDLL, vbTextCompare) <> 0 Then                     'If the module is new
    Dim hMod As Long                                                          '
                                                                              '
    If bUnicode Then                                                          'If unicode
      hMod = LoadLibraryW(StrPtr(strDll & vbNullChar))                        'Load the module with the unicode version of LoadLibrary
    Else                                                                      '
      hMod = LoadLibraryA(strDll)                                             'Load the module with the ascii version of LoadLibrary
    End If                                                                    '
                                                                              '
    If hMod = 0 Then                                                          'If the load failed
      Debug.Assert False                                                      'Halt if running under the VB IDE
      Err.Raise vbObjectError + 0, SRC & "CallFunc", "DLL failed load"        'Raise an error if running compiled
    End If                                                                    '
                                                                              '
    If hModule <> 0 Then                                                      'If a module is already loaded
      FreeLibrary hModule                                                     'Free the last module
    End If                                                                    '
                                                                              '
    hModule = hMod                                                            'Save the module handle
    strLastDLL = strDll                                                       'Save the new module name
    bNewDll = True                                                            'Indicate that it's a new module
  End If                                                                      '
                                                                              '
  If bNewDll Or StrComp(strFunc, strLastFunc, vbBinaryCompare) <> 0 Then      'If the function or module is new
    Dim fnAddress As Long                                                     'Function address
                                                                              '
    If Asc(strFunc) = 35 Then                                                 'If "#..." eg "#2", ordinal 2
      fnAddress = GetProcByOrdinal(hModule, CLng(Mid$(strFunc, 2)))           'Get the address of the function by ordinal
    Else                                                                      '
      fnAddress = GetProcByName(hModule, strFunc)                             'Get the address of the function by name
    End If                                                                    '
                                                                              '
    If fnAddress = 0 Then                                                     'If the function wasn't found in the module
      Debug.Assert False                                                      'Halt if running under the VB IDE
      Err.Raise vbObjectError + 1, SRC & "CallFunc", "Function not found"     'Raise an error if running compiled
    End If                                                                    '
                                                                              '
    strLastFunc = strFunc                                                     'Save the function name
    PutMem4 vCode + &H19, fnAddress - vCode - (&H19 + 4)                      'Patch the code with the relative address to the target function
  End If                                                                      '
                                                                              '
  With pb                                                                     '
    Dim i As Long                                                             'Parameter loop vars
    Dim j As Long                                                             'Parameter loop vars
                                                                              '
    j = UBound(ParamLongs)                                                    'Get the upper parameter array bound
    For i = 0 To j                                                            'For each parameter
      .Params(i) = ParamLongs(i)                                              'Store the parameter in the parameter block
    Next i                                                                    '
                                                                              '
    .ParamCount = i                                                           'Store the parameter count (j + 1)
  End With                                                                    '
                                                                              '
  CallFunc = CallCommon(RetType)                                              'Call common code
End Function                                                                  '

'CallPointer: call a function by address
'
' RetType     - Function return type
' fnAddress   - Address of the target function
' ParamLongs  - Any number of parameters As Long, or none.
'                 To pass the address (ByRef) of a string use StrPtr, eg. StrPtr(strPath)
'                 To pass the address (ByRef) of a variable or UDT use VarPtr, eg. VarPtr(i)
Public Function CallPointer(ByVal RetType As eReturnType, _
                            ByVal fnAddress As Long, _
                       ParamArray ParamLongs() As Variant) As Variant         '
  Dim i  As Long                                                              'Parameter loop vars
  Dim j  As Long                                                              'Parameter loop vars
                                                                              '
  With pb                                                                     '
    j = UBound(ParamLongs)                                                    'Get the upper parameter array bound
    For i = 0 To j                                                            'For each parameter
      .Params(i) = ParamLongs(i)                                              'Store the parameter in the parameter block
    Next i                                                                    '
                                                                              '
    .ParamCount = i                                                           'Store the parameter count (j + 1)
  End With                                                                    '
                                                                              '
  strLastFunc = vbNullString                                                  'Ensure we don't clash with CallFunc caching
  PutMem4 vCode + &H19, fnAddress - vCode - (&H19 + 4)                        'Patch the code with the relative address to the target function
  CallPointer = CallCommon(RetType)                                           'Call common code
End Function

'CallbackCdecl: return a wrapper address for a bas module routine to be used as a callback for a cdecl function.
'               Note: stdcall functions don't need a thunk to use a bas module function as a callback, use direct.
'
' nModFuncAddr - The address of the bas module callback function, use AddressOf to get this value
' nParms       - The number of parameters that will be passed to the bas module callback function
' nIndex       - Allow for multiple simultaneous callbacks
Public Function CallbackCdecl(ByVal nModFuncAddr As Long, _
                              ByVal nParams As Long, _
                     Optional ByVal nIndex As Long = 1) As Long
  
  If nIndex < 1 Or nIndex > 60 Or nParams > 60 Then                           'Parameter sanity checks
    Debug.Assert False                                                        'Halt if running under the VB IDE
    Err.Raise vbObjectError + 2, SRC & "CallbackCdecl", "Invalid parameter"   'Raise error if running compiled
  End If                                                                      '
                                                                              '
  CallbackCdecl = vCode + 128 + ((nIndex - 1) * 64)                           'Address of the callback wrapper. Pass this return value as the callback address parameter of the cdecl function
                                                                              '
  PutMem8 CallbackCdecl + 0, 465203369712025.6232@                            'Callback wrapper machine code
  PutMem8 CallbackCdecl + 8, -140418483381718.8339@                           '
  PutMem8 CallbackCdecl + 16, -801546908679710.9163@                          '
                                                                              '
  PutMem4 CallbackCdecl + 10, nModFuncAddr - CallbackCdecl - (10 + 4)         'Patch the code to call the vb bas module callback function
  PutMem1 CallbackCdecl + 16, nParams * 4                                     'Patch the code to apply the necessary stack adjustment
End Function                                                                  '
                                                                              '
'CallbackObj: return a wrapper address for an object callback from a cdecl or stdcall function
'
' objType     - Callback object type
' objCallback - The callback object
' nParams     - The number of parameters that will be passed to the object callback function
' nOrdinal    - Callback ordinal. 1 = last private function in the callback object, 2 = second last private function in the callback object, etc
' bCDECL      - Specifes whether the callback calling function is cdecl or stdcall
' nIndex      - Allow for multiple simultaneous callbacks
Public Function CallbackObj(ByVal objType As eObjType, _
                            ByRef objCallback As Object, _
                            ByVal nParams As Long, _
                   Optional ByVal nOrdinal As Long = 1, _
                   Optional ByVal bCDECL As Boolean = False, _
                   Optional ByVal nIndex As Long = 1) As Long
  Dim o As Long                                                               'Object pointer
  Dim i As Long                                                               'vTable entry counter
  Dim j As Long                                                               'vTable address
  Dim n As Long                                                               'Method pointer
  Dim b As Byte                                                               'First method byte
  Dim m As Byte                                                               'Known good first method byte
                                                                              '
  If nIndex < 1 Or nIndex > 60 Or nParams > 60 Then                           'Parameter sanity checks
    Debug.Assert False                                                        'Halt if running under the VB IDE
    Err.Raise vbObjectError + 3, SRC & "CallbackObj", "Invalid parameter"     'Raise error if running compiled
  End If                                                                      '
                                                                              '
  o = ObjPtr(objCallback)                                                     'Get the callback object's address
  GetMem4 o, j                                                                'Get the address of the callback object's vTable
  j = j + objType                                                             'Increment to the the first user entry for this callback object type
  GetMem4 j, n                                                                'Get the method pointer
  GetMem1 n, m                                                                'Get the first method byte... &H33 if pseudo-code, &HE9 if native
  j = j + 4                                                                   'Bump to the next vtable entry
                                                                              '
  For i = 1 To 511                                                            'Loop through a 'sane' number of vtable entries
    GetMem4 j, n                                                              'Get the method pointer
                                                                              '
    If IsBadCodePtr(n) Then                                                   'If the method pointer is an invalid code address
      GoTo vTableEnd                                                          'We've reached the end of the vTable, exit the for loop
    End If                                                                    '
                                                                              '
    GetMem1 n, b                                                              'Get the first method byte
                                                                              '
    If b <> m Then                                                            'If the method byte doesn't matche the known good value
      GoTo vTableEnd                                                          'We've reached the end of the vTable, exit the for loop
    End If                                                                    '
                                                                              '
    j = j + 4                                                                 'Bump to the next vTable entry
  Next i                                                                      'Bump counter
  
  Debug.Assert False                                                          'Halt if running under the VB IDE
  Err.Raise vbObjectError + 4, SRC & "CallbackObj", "Ordinal not found"       'Raise error if running compiled
                                                                              '
vTableEnd:                                                                    'We've hit the end of the vTable
  GetMem4 j - (nOrdinal * 4), n                                               'Get the method pointer for the specified ordinal
                                                                              '
  CallbackObj = vCode + 128 + ((nIndex - 1) * 64)                             'Address of the callback wrapper. Pass this return value as the callback address parameter
                                                                              '
  PutMem8 CallbackObj + 0, 648518346342877.6073@                              'Callback wrapper machine code
  PutMem8 CallbackObj + 8, 9425443492.7235@                                   '
  PutMem8 CallbackObj + 16, -29652486425477.8624@                             '
  PutMem8 CallbackObj + 24, 614907631944580.0296@                             '
  PutMem8 CallbackObj + 32, -444355163233240.1323@                            '
  PutMem4 CallbackObj + 40, &H90900055                                        '
                                                                              '
  PutMem1 CallbackObj + &HD, nParams                                          'Patch the number of params
  PutMem4 CallbackObj + &H19, o                                               'Patch the callback object
  PutMem4 CallbackObj + &H1E, n - CallbackObj - (&H1E + 4)                    'Patch the callback call address
  PutMem1 CallbackObj + &H28, IIf(bCDECL, 0, nParams * 4)                     'Patch the stack correction
End Function                                                                  '
                                                                              
Public Property Get FastCall() As Boolean                                     'Get FastCall flag
  FastCall = m_FastCall                                                       '
End Property                                                                  '
                                                                              '
Public Property Let FastCall(ByVal bValue As Boolean)                         'Let Fastcall flag
  m_FastCall = bValue                                                         '
  PutMem2 vCode + &H11, IIf(m_FastCall, &H34EB, &H9090)                       'Patch the code as per FastCall status
End Property                                                                  '
                                                                              
'IsValidDll - return whether the passed dll [path\]name is valid
'
' strDLL - [path\]name of the DLL
Public Function IsValidDll(ByRef strDll As String)                            '
  Dim hMod As Long                                                            '
                                                                              '
  If bUnicode Then                                                            'If unicode
    hMod = LoadLibraryW(StrPtr(strDll & vbNullChar))                          'Load the module with the unicode version of LoadLibrary
  Else                                                                        '
    hMod = LoadLibraryA(strDll)                                               'Load the module with the ascii version of LoadLibrary
  End If                                                                      '
                                                                              '
  If hMod Then                                                                'If the library loaded okay
    FreeLibrary hMod                                                          'Free the library
    IsValidDll = True                                                         'Indicate success
  End If                                                                      '
End Function                                                                  '

'IsValidMethod - return whether the passed dll [path\]name / method name is valid
'
' strDLL   - [path\]name of the DLL
' strFunc  - Name of the function or it's ordinal value preceded by a '#' eg. "#2"
Public Function IsValidMethod(ByRef strDll As String, _
                              ByRef strFunc As String)                        '
  Dim hMod As Long                                                            '
                                                                              '
  If bUnicode Then                                                            'If unicode
    hMod = LoadLibraryW(StrPtr(strDll & vbNullChar))                          'Load the module with the unicode version of LoadLibrary
  Else                                                                        '
    hMod = LoadLibraryA(strDll)                                               'Load the module with the ascii version of LoadLibrary
  End If                                                                      '
                                                                              '
  If hMod Then                                                                'If the library loaded okay
    Dim nFuncAddr As Long                                                     'Function address
                                                                              '
    If Asc(strFunc) = 35 Then                                                 'If "#..." eg "#2", ordinal 2
      nFuncAddr = GetProcByOrdinal(hModule, CLng(Mid$(strFunc, 2)))           'Get the address of the function by ordinal
    Else                                                                      '
      nFuncAddr = GetProcByName(hModule, strFunc)                             'Get the address of the function by name
    End If                                                                    '
                                                                              '
    If nFuncAddr Then                                                         'If the function was found in the module
      IsValidMethod = True                                                    'Indicate success
    End If                                                                    '
                                                                              '
    FreeLibrary hMod                                                          'Free the library
  End If                                                                      '
End Function                                                                  '

Public Property Get LastError() As Long                                       'Get last error
  LastError = m_LastError                                                     '
End Property                                                                  '
                                                                              
'CallCommon: common CallFunc/CallPointer code
'
' RetType - Function return type
Private Function CallCommon(ByVal RetType As eReturnType) As Variant
  PutMem1 vCode + &H27, RetType                                               'Patch the return type jump
                                                                              '
  SetLastError 0                                                              'Clear the error code
                                                                              '
  'N.B. we patch the vTable on each call because there could be multiple
  'instances of this class. Multiple instances share the same code...
  'and would otherwise share the vCode of the last created instance.
  'So we re-patch the vTable on each call to ensure the entry is hooked
  'to the instance's vCode
  Select Case RetType                                                         'Select on return type
    Case eReturnType.retByte                                                  'Return a Byte
      PutMem4 vTable + (19 * 4), vCode                                        'Patch the z_CallFunc_i08 entry to point to vCode
      CallCommon = z_CallFunc_i08(nAddrPb)                                    'Call
                                                                              '
    Case eReturnType.retInteger                                               'Return an Integer
      PutMem4 vTable + (20 * 4), vCode                                        'Patch the z_CallFunc_i16 entry to point to vCode
      CallCommon = z_CallFunc_i16(nAddrPb)                                    'Call
                                                                              '
    Case eReturnType.retLong                                                  'Return a Long
      PutMem4 vTable + (21 * 4), vCode                                        'Patch the z_CallFunc_i32 entry to point to vCode
      CallCommon = z_CallFunc_i32(nAddrPb)                                    'Long
                                                                              '
    Case eReturnType.retInt64                                                 'Return 64bits (e.g. Currency)
      PutMem4 vTable + (22 * 4), vCode                                        'Patch the z_CallFunc_i64 entry to point to vCode
      CallCommon = z_CallFunc_i64(nAddrPb)                                    'Call
                                                                              '
    Case eReturnType.retSingle                                                'Return a Single
      PutMem4 vTable + (23 * 4), vCode                                        'Patch the z_CallFunc_Sng entry to point to vCode
      CallCommon = z_CallFunc_Sng(nAddrPb)                                    'Call
                                                                              '
    Case eReturnType.retDouble                                                'Return a Double
      PutMem4 vTable + (24 * 4), vCode                                        'Patch the z_CallFunc_Dbl entry to point to vCode
      CallCommon = z_CallFunc_Dbl(nAddrPb)                                    'Call
                                                                              '
    Case eReturnType.retSub                                                   'Subroutine, no return value
      PutMem4 vTable + (25 * 4), vCode                                        'Patch the z_CallFunc_Sub entry to point to vCode
      Call z_CallFunc_Sub(nAddrPb)                                            'Call
      
    Case Else                                                                 'Undefined return type
      Debug.Assert False                                                      'Halt if running under the VB IDE
      Err.Raise vbObjectError + 5, SRC & "CallCommon", "Unknown return type"  'Raise error if running compiled
  End Select                                                                  '
                                                                              '
  m_LastError = GetLastError()                                                'Get the error code
End Function

'Class_Initialize: initialize the cCallFunc2 instance
Private Sub Class_Initialize()                                                '
  vCode = VirtualAlloc(0, &H1000&, &H1000&, &H40&)                            'Allocate 4k of read/write/executable memory
                                                                              '
  PutMem8 vCode + 0, 695618785647368.6248@                                    'Universal function caller machine code
  PutMem8 vCode + 8, -208726556020175.3831@                                   '
  PutMem8 vCode + 16, -29652486425143.4233@                                   '
  PutMem8 vCode + 24, 614902794093417.828@                                    '
  PutMem8 vCode + 32, 193965741455568.6229@                                   '
  PutMem8 vCode + 40, -151277692825560.6392@                                  '
  PutMem8 vCode + 48, -857442152266638.7183@                                  '
  PutMem8 vCode + 56, 21029022751752.3025@                                    '
  PutMem8 vCode + 64, -7203916540378.4739@                                    '
  PutMem8 vCode + 72, -61276775362635.1564@                                   '
  PutMem8 vCode + 80, -454553025687766.4117@                                  '
                                                                              '
  GetMem4 ObjPtr(Me), vTable                                                  'Get the address of the class vTable
                                                                              '
  If GetProcByName(LoadLibraryA("user32"), "IsWindowUnicode") Then            'Is IsWindowUnicode present
    bUnicode = IsWindowUnicode(GetDesktopWindow())                            'Determine whether we'll use the unicode version of LoadLibrary
  End If                                                                      '
                                                                              '
  FastCall = False                                                            'Default to non-Fastcall
  nAddrPb = VarPtr(pb)                                                        'Address of the parameter block
End Sub                                                                       '
                                                                              '
'Class_Terminate: cleanup the cCallFunc2 instance
Private Sub Class_Terminate()                                                 '
  If hModule <> 0 Then                                                        'If a module is loaded
    FreeLibrary hModule                                                       'Free the loaded module
  End If                                                                      '
                                                                              '
  VirtualFree vCode, 0, &H8000&                                               'Free the allocated memory
End Sub
                                                                              
'*******************************************************************************************************************
' DO NOT MOVE THE FOLLOWING ROUTINES...
Private Function z_CallFunc_i08(ByVal nParmAddr As Long) As Byte:     End Function
Private Function z_CallFunc_i16(ByVal nParmAddr As Long) As Integer:  End Function
Private Function z_CallFunc_i32(ByVal nParmAddr As Long) As Long:     End Function
Private Function z_CallFunc_i64(ByVal nParmAddr As Long) As Currency: End Function
Private Function z_CallFunc_Sng(ByVal nParmAddr As Long) As Single:   End Function
Private Function z_CallFunc_Dbl(ByVal nParmAddr As Long) As Double:   End Function
     Private Sub z_CallFunc_Sub(ByVal nParmAddr As Long):             End Sub
