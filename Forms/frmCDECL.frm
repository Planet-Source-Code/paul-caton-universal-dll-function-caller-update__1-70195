VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmCDECL 
   Caption         =   "CDECL Test"
   ClientHeight    =   5325
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4950
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5325
   ScaleWidth      =   4950
   Begin VB.OptionButton optCallback 
      Caption         =   "frm callback"
      Height          =   255
      Index           =   1
      Left            =   3480
      TabIndex        =   11
      Top             =   5040
      Width           =   1320
   End
   Begin VB.OptionButton optCallback 
      Caption         =   "bas callback"
      Height          =   255
      Index           =   0
      Left            =   3495
      TabIndex        =   10
      Top             =   4680
      Value           =   -1  'True
      Width           =   1320
   End
   Begin VB.CheckBox chkAscending 
      Caption         =   "Ascending"
      Height          =   210
      Left            =   3675
      TabIndex        =   3
      Top             =   1815
      Value           =   1  'Checked
      Width           =   1140
   End
   Begin VB.CommandButton cmdSort 
      Caption         =   "Sort"
      Height          =   390
      Left            =   3660
      TabIndex        =   2
      Top             =   1110
      Width           =   1320
   End
   Begin MSComctlLib.ListView lv 
      Height          =   2055
      Left            =   345
      TabIndex        =   1
      Top             =   480
      Width           =   2745
      _ExtentX        =   4842
      _ExtentY        =   3625
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      PictureAlignment=   5
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Index"
         Object.Width           =   1191
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Values"
         Object.Width           =   1508
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Sorted"
         Object.Width           =   1508
      EndProperty
      Picture         =   "frmCDECL.frx":0000
   End
   Begin VB.CommandButton cmdPopulate 
      Caption         =   "Populate"
      Height          =   390
      Left            =   3615
      TabIndex        =   0
      Top             =   525
      Width           =   1320
   End
   Begin VB.Label lblStatic 
      AutoSize        =   -1  'True
      Caption         =   "Comparisons:"
      Height          =   210
      Index           =   2
      Left            =   3630
      TabIndex        =   9
      Top             =   3930
      Width           =   1065
   End
   Begin VB.Label lblComp 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   3615
      TabIndex        =   8
      Top             =   4185
      Width           =   1320
   End
   Begin VB.Label lblStatic 
      AutoSize        =   -1  'True
      Caption         =   "Sort elements:"
      Height          =   210
      Index           =   0
      Left            =   3630
      TabIndex        =   7
      Top             =   2295
      Width           =   1215
   End
   Begin VB.Label lblElements 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   3615
      TabIndex        =   6
      Top             =   2550
      Width           =   1320
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   3660
      TabIndex        =   5
      Top             =   3405
      Width           =   1320
   End
   Begin VB.Label lblStatic 
      AutoSize        =   -1  'True
      Caption         =   "Sort time:"
      Height          =   210
      Index           =   1
      Left            =   3675
      TabIndex        =   4
      Top             =   3150
      Width           =   825
   End
End
Attribute VB_Name = "frmCDECL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'**********************************************************************************
'** Demonstrate how to call CDECL dll functions using the cCallFunc2 class.
'**********************************************************************************

Option Explicit

Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (cPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (cFrequency As Currency) As Long

Private Const ELEMENTS        As Long = 10000

Private arInt(1 To ELEMENTS)  As Integer
Private qpF                   As Currency
Private sDLL                  As String
Private oComparisons          As Long
Private p                     As cCallFunc2

Private Sub Form_Load()
  Dim i     As Long
  Dim r     As Long
  Dim sBuf  As String
  Dim sFmt  As String
  
  Set p = New cCallFunc2                                    'Create cCallFunc2 class instance
  sDLL = "MSVCRT20"                                         'DLL name (MicroSoft Visual C Run Time version 2.0)
  
  QueryPerformanceFrequency qpF                             'Get performance counter frequency
  lblElements.Caption = Format$(ELEMENTS, "#,###")          'Display the number of array elements
  
  cmdSort.Enabled = False
  chkAscending.Enabled = False
  
  r = p.CallFunc(sDLL, retLong, "time", 0)                  'Get the time (in secs) for use as a random number seed
  p.CallFunc sDLL, retSub, "srand", r                       'Seed the random number generator
  
  sBuf = String$(20, vbNullChar)                            'Make some space
  sFmt = "%6d" & vbNullChar                                 'The format...
  
  For i = 1 To ELEMENTS
    r = p.CallFunc(sDLL, retLong, "swprintf", StrPtr(sBuf), StrPtr(sFmt), i)
    lv.ListItems.Add , , StripNull(sBuf)                    'C strings are NULL terminated, strip it and add to the the list
  Next i
  
  Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 4
End Sub

Private Sub Form_Resize()
  On Error Resume Next
    cmdPopulate.Move ScaleWidth - 240 - cmdPopulate.Width, 240
    cmdSort.Move cmdPopulate.Left, cmdPopulate.Top + cmdPopulate.Height + 240, cmdPopulate.Width, cmdPopulate.Height
    chkAscending.Move cmdSort.Left, cmdSort.Top + cmdSort.Height + 240
    lblStatic(0).Move chkAscending.Left, chkAscending.Top + chkAscending.Height + 240
    lblElements.Move lblStatic(0).Left, lblStatic(0).Top + lblStatic(0).Height + 30, cmdSort.Width
    lblStatic(1).Move lblElements.Left, lblElements.Top + lblElements.Height + 240
    lblTime.Move lblStatic(1).Left, lblStatic(1).Top + lblStatic(1).Height + 30, cmdSort.Width
    lblStatic(2).Move lblTime.Left, lblTime.Top + lblTime.Height + 240
    lblComp.Move lblStatic(2).Left, lblStatic(2).Top + lblStatic(2).Height + 30, cmdSort.Width
    optCallback(0).Move lblComp.Left - 30, lblComp.Top + lblComp.Height + 240
    optCallback(1).Move optCallback(0).Left, optCallback(0).Top + optCallback(0).Height + 60
    lv.Move 240, 240, cmdPopulate.Left - 480, Me.ScaleHeight - 480
  On Error GoTo 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set p = Nothing                                           'Destroy the cCallFunc2 class
End Sub

'Populate the list with random integers
Private Sub cmdPopulate_Click()
  Dim i As Long
  
  cmdSort.Enabled = True
  chkAscending.Enabled = True
  
  LockWindowUpdate lv.hWnd
    For i = 1 To ELEMENTS
      arInt(i) = CInt(p.CallFunc(sDLL, retLong, "rand"))
      With lv.ListItems(i)
        .SubItems(1) = Format$(arInt(i), "#,##0")
        .SubItems(2) = vbNullString
      End With
    Next i
  LockWindowUpdate 0
End Sub

'Sort the list
Private Sub cmdSort_Click()
  Dim i             As Long
  Dim nCallbackAddr As Long
  Dim qpC1          As Currency
  Dim qpC2          As Currency
  
'The 'C' qsort (quicksort) library routine can sort anything provided the data, or
'pointers to the data, lie in contiguous memory. The qsort routine implements the
'quicksort algorithm logic, the user provides the comparison code in a callback routine.
'All that qsort requires is...
'   the address of the first element
'   the number of elements
'   the size in bytes of each element
'   the address of a user provided routine that will compare two elements and return
'     a numeric indication of greater or less than.
'
'The issue with the qsort routine is that it will make the callback as a cdecl routine. If we
'merely pass the address of a .bas module function it will crash because of the different stack
'correction conventions between cdecl and stdcall. What we need to do is create a wrapper around our
'bas module function that will perform the necessary stack correction.. we pass the address of
'the wrapper function instead of the .bas function
  
  If optCallback(0).Value Then 'bas callback
    mCallback.nComparisons = 0
    
    'Prepare the .bas module callback function wrapper
    If chkAscending.Value = vbChecked Then
      nCallbackAddr = p.CallbackCdecl(AddressOf mCallback.qsort_compare_up, 2)
     Else
      nCallbackAddr = p.CallbackCdecl(AddressOf mCallback.qsort_compare_dn, 2)
    End If
  
  Else 'frm Callback
    oComparisons = 0
  
    'Prepare the form object callback function wrapper
    If chkAscending.Value = vbChecked Then
      nCallbackAddr = p.CallbackObj(objFrm, Me, 2, 1, True)
     Else
      nCallbackAddr = p.CallbackObj(objFrm, Me, 2, 2, True)
    End If
  End If
  
  'Sort the array
  QueryPerformanceCounter qpC1
    p.CallFunc sDLL, retSub, "qsort", VarPtr(arInt(1)), ELEMENTS, 2, nCallbackAddr
  QueryPerformanceCounter qpC2
  
  'Fill the ListView with the sorted array
  LockWindowUpdate lv.hWnd
    For i = 1 To ELEMENTS
      lv.ListItems(i).SubItems(2) = Format$(arInt(i), "#,##0")
    Next i
  LockWindowUpdate 0
  
  lblTime.Caption = Format$((qpC2 - qpC1) / qpF, "0.0000") & "s"
  
  If optCallback(0).Value Then 'bas callback
    lblComp.Caption = Format$(mCallback.nComparisons, "#,###")
  Else 'frm callback
    lblComp.Caption = Format$(oComparisons, "#,###")
  End If
End Sub

'Strip any string terminating nulls
Private Function StripNull(s As String) As String
  Dim i As Long
  
  i = InStr(1, s, vbNullChar)
  
  If i > 0 Then
    StripNull = Left$(s, i - 1)
   Else
    StripNull = s
  End If
End Function

'*************************************************************************************************
'** Callbacks - the final private routine is ordinal #1, second last is ordinal #2 etc
'**
'** These callback routines *must* be private and the final routines in the file
'*************************************************************************************************

'Ordinal #2
Private Function qsort_compare_dn(ByRef arg1 As Integer, ByRef arg2 As Integer) As Long
  qsort_compare_dn = arg2 - arg1
  oComparisons = oComparisons + 1 'Just for info
End Function

'Ordinal #1
Private Function qsort_compare_up(ByRef arg1 As Integer, ByRef arg2 As Integer) As Long
  qsort_compare_up = arg1 - arg2
  oComparisons = oComparisons + 1 'Just for info
End Function
