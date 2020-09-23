VERSION 5.00
Begin VB.Form frmCallbackObj 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Callback sample"
   ClientHeight    =   8100
   ClientLeft      =   165
   ClientTop       =   480
   ClientWidth     =   8370
   ClipControls    =   0   'False
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   FontTransparent =   0   'False
   ForeColor       =   &H00800000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   540
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   558
   StartUpPosition =   2  'CenterScreen
   Begin VB.Menu mnuCallback 
      Caption         =   "&Callback"
      Begin VB.Menu mnuItem 
         Caption         =   "EnumWindows"
         Index           =   0
      End
      Begin VB.Menu mnuItem 
         Caption         =   "EnumFontFamilies"
         Index           =   1
      End
      Begin VB.Menu mnuItem 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuItem 
         Caption         =   "Subclass"
         Index           =   3
      End
      Begin VB.Menu mnuItem 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuItem 
         Caption         =   "Smooth scrolling"
         Index           =   5
      End
      Begin VB.Menu mnuItem 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuItem 
         Caption         =   "E&xit"
         Index           =   7
      End
   End
End
Attribute VB_Name = "frmCallbackObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'*************************************************************************************************
'** frmCallbackObj - Demonstrate object callbacks
'*************************************************************************************************

Option Explicit

Private Const LF_FACESIZE As Long = 32

Private Type LOGFONT
  lfHeight                As Long
  lfWidth                 As Long
  lfEscapement            As Long
  lfOrientation           As Long
  lfWeight                As Long
  lfItalic                As Byte
  lfUnderline             As Byte
  lfStrikeOut             As Byte
  lfCharSet               As Byte
  lfOutPrecision          As Byte
  lfClipPrecision         As Byte
  lfQuality               As Byte
  lfPitchAndFamily        As Byte
  lfFaceName(LF_FACESIZE) As Byte
End Type

Private Type NEWTEXTMETRIC
  tmHeight                As Long
  tmAscent                As Long
  tmDescent               As Long
  tmInternalLeading       As Long
  tmExternalLeading       As Long
  tmAveCharWidth          As Long
  tmMaxCharWidth          As Long
  tmWeight                As Long
  tmOverhang              As Long
  tmDigitizedAspectX      As Long
  tmDigitizedAspectY      As Long
  tmFirstChar             As Byte
  tmLastChar              As Byte
  tmDefaultChar           As Byte
  tmBreakChar             As Byte
  tmItalic                As Byte
  tmUnderlined            As Byte
  tmStruckOut             As Byte
  tmPitchAndFamily        As Byte
  tmCharSet               As Byte
  ntmFlags                As Long
  ntmSizeEM               As Long
  ntmCellHeight           As Long
  ntmAveWidth             As Long
End Type

Private Type RECT
  Left                    As Long
  Top                     As Long
  Right                   As Long
  Bottom                  As Long
End Type

Private Declare Function CallWindowProcA Lib "user32" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function EnumFontFamiliesA Lib "gdi32" (ByVal hDC As Long, ByVal lpszFamily As String, ByVal lpEnumFontFamProc As Long, lParam As Any) As Long
Private Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowTextA Lib "user32" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (cPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (cFrequency As Currency) As Long
Private Declare Function ScrollWindowEx Lib "user32" (ByVal hWnd As Long, ByVal dx As Long, ByVal dy As Long, lprcScroll As RECT, lprcClip As RECT, ByVal hrgnUpdate As Long, lprcUpdate As Any, ByVal fuScroll As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function UpdateWindow Lib "user32" (ByVal hWnd As Long) As Long

Private p                 As cCallFunc2
Private nOriginalWndProc  As Long                           'Original WndProc
Private nLineHeight       As Long                           'Height of a line of text
Private qpF               As Currency                       'Performance frequency
Private rc                As RECT                           'Scrolling rectangle

Private Sub Form_Load()
  Set p = New cCallFunc2                                    'Create cCallFunc2 class instance
  
  nLineHeight = Me.TextHeight("My")                         'Get the height in pixels of a line of text
  QueryPerformanceFrequency qpF                             'Get performance counter frequency
End Sub

Private Sub Form_Resize()
  rc.Right = Me.ScaleWidth
  rc.Bottom = Me.ScaleHeight
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If nOriginalWndProc Then
    Const WM_CLOSE = &H10

    SetWindowLongA Me.hWnd, -4, nOriginalWndProc            'Restore the original WndProc
    nOriginalWndProc = 0                                    'Indicate that we're not subclassed
    Cancel = True                                           'Cancel this unload, it isn't safe to quit yet as the pre-existing subclassing hasn't fully played out
    PostMessage Me.hWnd, WM_CLOSE, 0, 0                     'Initiate another unload
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set p = Nothing                                           'Destroy the cCallFunc2 class
End Sub

Private Sub mnuItem_Click(Index As Integer)
  Dim nReturn   As Long
  Dim nCallback As Long
  Dim qpC1      As Currency
  Dim qpC2      As Currency
  
  DoEvents
  
  Select Case Index
  Case 0 'EnumWindows
    nCallback = p.CallbackObj(objFrm, Me, 2, 1)   'Callback ordinal #1 with 2 parameters
    Display "*** EnumWindows"
    QueryPerformanceCounter qpC1
    nReturn = EnumWindows(nCallback, 123)
    QueryPerformanceCounter qpC2
    Display "*** EnumWindows returns: " & nReturn & ", time: " & Format$((qpC2 - qpC1) / qpF, "0.0000") & "s"
    
  Case 1 'EnumFontFamilies
    nCallback = p.CallbackObj(objFrm, Me, 4, 2) 'Callback ordinal #2 with 4 parameters
    Display "*** EnumFontFamilies"
    QueryPerformanceCounter qpC1
    nReturn = EnumFontFamiliesA(Me.hDC, vbNullString, nCallback, 0)
    QueryPerformanceCounter qpC2
    With Me
      .FontName = "Tahoma"
      .FontBold = False
      .FontItalic = False
      .FontSize = 10
      .FontStrikethru = False
      .FontUnderline = False
    End With
    nLineHeight = Me.TextHeight("My")
    Display "*** EnumFontFamilies returns: " & nReturn & ", time: " & Format$((qpC2 - qpC1) / qpF, "0.0000") & "s"

  Case 3 'Subclassing
    If mnuItem(Index).Checked Then
      SetWindowLongA Me.hWnd, -4, nOriginalWndProc
      nOriginalWndProc = 0
      mnuItem(Index).Checked = False
      DoEvents
      Display "*** Subclass ending."
    Else
      nCallback = p.CallbackObj(objFrm, Me, 4, 3, , 2) 'Callback ordinal #3 with 4 parameters, index 2 becuae it will operate in parallel with other callbacks
      Display "*** Subclass start. WARNING, the Subclass callback isn't 'End' safe."
      nOriginalWndProc = SetWindowLongA(Me.hWnd, -4, nCallback)
      mnuItem(Index).Checked = True
    End If
    
  Case 5
    mnuItem(Index).Checked = Not mnuItem(Index).Checked
  
  Case 7
    Unload Me
    
  End Select
End Sub

Private Sub Display(ByVal sText As String)
  Const SW_INVALIDATE As Long = &H2
  
  If Me.CurrentY + nLineHeight > Me.ScaleHeight Then
    If mnuItem(5).Checked Then
      Dim i As Long
      
      For i = 1 To nLineHeight
        ScrollWindowEx Me.hWnd, 0, -1, rc, rc, 0, ByVal 0&, SW_INVALIDATE
      Next i
    Else
      ScrollWindowEx Me.hWnd, 0, -nLineHeight, rc, rc, 0, ByVal 0&, SW_INVALIDATE
    End If
    
    UpdateWindow Me.hWnd
    Me.CurrentY = Me.ScaleHeight - nLineHeight
  End If
  
  Print " " & sText
End Sub

Private Function Hfmt(ByVal nValue As Long) As String
  Hfmt = Right$("0000000" & Hex$(nValue), 8)
End Function

'*************************************************************************************************
'** Callbacks - the final private routine is ordinal #1, second last is ordinal #2 etc
'**
'** These callback routines *must* be private and the final routines in the file
'*************************************************************************************************

'Callback ordinal 3
Private Function WndProc(ByVal lng_hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Const WM_PAINT = &HF
  
  WndProc = CallWindowProcA(nOriginalWndProc, lng_hWnd, uMsg, wParam, lParam)
  
  If uMsg <> WM_PAINT Then
    Display "hWnd: " & Hfmt(lng_hWnd) & ", " & _
            "uMsg: " & Hfmt(uMsg) & ", " & _
            "wParam: " & Hfmt(wParam) & ", " & _
            "lParam: " & Hfmt(lParam) & ", " & _
            "Return: " & Hfmt(WndProc)
  End If
End Function

'Callback ordinal 2
Private Function EnumFontFamProc(lpNLF As LOGFONT, lpNTM As NEWTEXTMETRIC, ByVal FontType As Long, ByVal lParam As Long) As Long
  Dim FaceName As String
  
  FaceName = StrConv(lpNLF.lfFaceName, vbUnicode)
  FaceName = Left$(FaceName, InStr(FaceName, vbNullChar) - 1)
  
  With Me
    .FontName = FaceName
    .FontBold = False
    .FontItalic = False
    .FontSize = 10
    .FontStrikethru = False
    .FontUnderline = False
  End With
  
  nLineHeight = Me.TextHeight("My")
  Display FaceName
  
  EnumFontFamProc = 1
End Function

'Callback ordinal 1
Private Function EnumWindowsProc(ByVal lng_hWnd As Long, ByVal lParam As Long) As Long
  
  If IsWindowVisible(lng_hWnd) Then
    Dim nLen     As Long
    Dim sCaption As String
    
    sCaption = Space$(256)
    nLen = GetWindowTextA(lng_hWnd, sCaption, 255)
    
    If nLen > 0 Then
      Display Left$(sCaption, nLen)
    End If
  End If
  
  EnumWindowsProc = 1
End Function
