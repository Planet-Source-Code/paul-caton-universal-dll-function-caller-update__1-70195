Attribute VB_Name = "mCallback"

'**********************************************************************************
'** mCallback.bas - qsort callback routines
'**********************************************************************************

Option Explicit

'Just for info... you can see the number of comparisons that were required in these different scenarios
'a) data is random
'b) data is sorted
'c) data is sorted but in the wrong sort order
Public nComparisons     As Long

'This is one way of doing it (ascending/descending) - the other (better) beneath...
'
'Public bAscending As Boolean
'
'Public Function qsort_compare(ByRef arg1 As Integer, ByRef arg2 As Integer) As Long
'  If bAscending Then
'    qsort_compare = arg1 - arg2
'  Else
'    qsort_compare = arg2 - arg1
'  End If
'
'  nComparisons = nComparisons + 1
'End Function

'This way we pick the compare routine at runtime... and save ~170,000 If/Else's
Public Function qsort_compare_dn(ByRef arg1 As Integer, ByRef arg2 As Integer) As Long
  qsort_compare_dn = arg2 - arg1
  nComparisons = nComparisons + 1
End Function

Public Function qsort_compare_up(ByRef arg1 As Integer, ByRef arg2 As Integer) As Long
  qsort_compare_up = arg1 - arg2
  nComparisons = nComparisons + 1
End Function
