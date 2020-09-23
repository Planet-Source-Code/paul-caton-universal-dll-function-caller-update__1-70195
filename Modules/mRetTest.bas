Attribute VB_Name = "mRetTest"

'**********************************************************************************
' Super-simplistic return type demo... also shows how to pass 64bit variables
' (Double & Currency) ByVal
'

Option Explicit

'There aren't any parameters in the following dll calls, thus, despite the
'functions being nominally cdecl, we *can* call them from VB in the usual way.
'We do so here to ensure that the exact same values are returned as using the
'cCallFunc2 class

Private Declare Function GetByte Lib "RetTest.dll" () As Byte
Private Declare Function GetInteger Lib "RetTest.dll" () As Integer
Private Declare Function GetLong Lib "RetTest.dll" () As Long
Private Declare Function GetPointer Lib "RetTest.dll" () As Long
Private Declare Function GetInt64 Lib "RetTest.dll" () As Currency
Private Declare Function GetSingle Lib "RetTest.dll" () As Single
Private Declare Function GetDouble Lib "RetTest.dll" () As Double

Private Type FILETIME
  dwLowDateTime     As Long
  dwHighDateTime    As Long
End Type

Private Type WIN32_FIND_DATA
  dwFileAttributes  As Long
  ftCreationTime    As FILETIME
  ftLastAccessTime  As FILETIME
  ftLastWriteTime   As FILETIME
  nFileSizeHigh     As Long
  nFileSizeLow      As Long
  dwReserved0       As Long
  dwReserved1       As Long
  cFileName         As String * 260
  cAlternate        As String * 14
End Type

Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Sub PutMem8 Lib "msvbvm60" (ByVal Addr As Long, ByVal NewVal As Double)

'Type used to pass a 64bit parameter ByVal
Private Type tParam64
  PartA As Long
  PartB As Long
End Type

Private Sub Main()
  Dim cf2     As New cCallFunc2
  Dim strDll  As String
  Dim b       As Byte
  Dim i       As Integer
  Dim l       As Long
  Dim p       As Long
  Dim c       As Currency
  Dim s       As Single
  Dim d       As Double
  
  ChDrive App.Path        'Ensure we're running on the app's own drive
  ChDir App.Path          'Ensure we're running in the app's own directory
  
  strDll = "RetTest.dll"
  
  'If the Dll isn't present, extract it from the app resource
  If Not FileExists(strDll) Then ExtractDLL 101, strDll
  
  '********************************************************************************************************************************************************************
  'Test using the declared functions
  '********************************************************************************************************************************************************************
  b = GetByte()
  i = GetInteger()
  l = GetLong()
  p = GetPointer()
  c = GetInt64()
  s = GetSingle()
  d = GetDouble()
  Debug.Print "Byte: &H" & Hex(b) & ", Integer: &H" & Hex(i) & ", Long: &H" & Hex(l) & ", Pointer: &H" & Hex(p) & ", 64bit: " & c & ", Single: " & s & ", Double: " & d
  
  '********************************************************************************************************************************************************************
  'Test using the cCallFunc2 class
  '********************************************************************************************************************************************************************
  b = cf2.CallFunc(strDll, retByte, "GetByte")
  i = cf2.CallFunc(strDll, retInteger, "GetInteger")
  l = cf2.CallFunc(strDll, retLong, "GetLong")
  p = cf2.CallFunc(strDll, retLong, "GetPointer")
  c = cf2.CallFunc(strDll, retInt64, "GetInt64")
  s = cf2.CallFunc(strDll, retSingle, "GetSingle")
  d = cf2.CallFunc(strDll, retDouble, "GetDouble")
  Debug.Print "Byte: &H" & Hex(b) & ", Integer: &H" & Hex(i) & ", Long: &H" & Hex(l) & ", Pointer: &H" & Hex(p) & ", 64bit: " & c & ", Single: " & s & ", Double: " & d
  
  'No value returned... like a VB Sub
  cf2.CallFunc strDll, retSub, "Subroutine", &H12345678
  
  '********************************************************************************************************************************************************************
  'Demonstrate how to pass a Double parameter (or the 64bit Currency type) ByVal
  '
  'ByRef isn't an issue because the address of the variable is passed, not its value.
  'On a 32 bit OS (or emulation) the address is always 32 bits, no matter the size of the type referenced, and thus fits into a Long.
  '********************************************************************************************************************************************************************
  Dim dbl     As Double
  Dim Param64 As tParam64
 
  dbl = 3.14159 'whatever
   
  PutMem8 VarPtr(Param64), dbl
  
  d = cf2.CallFunc(strDll, retDouble, "GetDblInc", Param64.PartA, Param64.PartB)
  Debug.Print "dbl: " & dbl & ", GetDblInc: " & d
End Sub

'Extract the specified resource and save to disk
Private Sub ExtractDLL(ByVal nIndex As Long, ByRef strDll As String)
  Dim FileNum     As Integer
  Dim DataArray() As Byte
  
  DataArray = LoadResData(nIndex, "CUSTOM")
  
  FileNum = FreeFile
  Open strDll For Binary As #FileNum
  Put #FileNum, 1, DataArray()
  Close #FileNum
  
  Erase DataArray
End Sub

'Return whether the passed file exists
Private Function FileExists(ByRef strFile As String) As Boolean
  Dim WFD   As WIN32_FIND_DATA
  Dim hFile As Long
  
  hFile = FindFirstFile(strFile, WFD)
  
  If hFile <> -1 Then
    FileExists = True
    FindClose hFile
  End If
End Function
