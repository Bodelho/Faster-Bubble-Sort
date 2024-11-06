'
Option Explicit
'
'===============================================================================================================
'
'===============================================================================================================
'
Public Function Bubble_Sort_Vector(Target As Variant, _
                                   Optional ByVal Replace As Variant = True, _
                                   Optional ByVal Descending As Variant = False, _
                                   Optional ByVal Comparison As Variant = vbTextCompare) As Variant
  ' ------------------------------------------------------------------------------------------------------------
  ' Bubble_Sort_Vector() will sort the one-dimension array (vector) provided in Target(). By default,  the sort
  ' will be performed in ascending  order; setting optional argument Descending to True will  cause the sort to
  ' be performed in descending order.
  '
  ' Bubble_Sort_Vector() will return either an array when the sort operation succeeds,  or  an Error() when the
  ' sort operation fails - that is, testing  for Bubble_Sort_Vector()'s returned  data for "IsError()" suffices
  ' in order to find whether or not the sort operation finished successfully.
  '
  ' When   the   sort  operation  succeeds  and  optional   argument  Replace   is   set   to   True (default),
  ' Bubble_Sort_Vector() returns an  empty array.  In this case (success), the relative order of  Target() cell
  ' values will have been rearranged by the sort operation.
  '
  ' When  the sort  operation  is  successfully  completed  and  optional  argument  Replace  is set  to False,
  ' Bubble_Sort_Vector() returns a "LBound() = 0"  array having as  many cells  as  Target()  has; the returned
  ' array carries  subscripts to Target()  cells telling the order  in  which Target() cells must be fetched in
  ' order to have their  cell  values  in  the requested  sort order  (ascending  /  descending). In this case,
  ' Target()  cells  are NOT  rearranged  -  that  is,  Target()  remains  untouched.  For  instance,  be  "R =
  ' Bubble_Sort_Vector(Some_Array(),  False, ...)",  then after a successful sort  operation "Some_Array(R(0))"
  ' will have  "Some_Array()"'s smallest cell value found in "Some_Array()"'s cells - if the sort was performed
  ' in ascending order - or will have the largest such value if the sort was performed in descending order.
  '
  ' Target() must be a one-dimension array. Target() cells  must have a valid,  acceptable data type.  The sort
  ' operation  will fail if  ANY Target() cell  is found  to be  of any of the  following  data types: vbEmpty,
  ' vbNull, vbObject, vbError, vbDataObject, vbUserDefinedType, vbArray. All Target() cells must share the same
  ' base data type of the first Target() cell (whatever its subscript may be).
  '
  ' Target()'s LBound() and UBound() may be any as long as they are valid and as long as "LBound() <= UBound()"
  ' is True.
  '
  ' Optional argument Comparison (default =  vbTtextCompare) defines the  comparison method  to be used  in the
  ' sort operation: either vbTtextCompare or vbBinaryCompare. If neither of those  is  provided, vbTtextCompare
  ' will be used as comparison method (no error will be raised).
  '
  ' In case of errors, Error(N) will be returned having one of the following error codes:
  '  1  ...  Target is not an array
  '  2  ...  Target() is not a vector
  '  4  ...  sort key has an invalid data type
  '  5  ...  sort key has values of intermixed data types
  '  9  ...  fatal failure in sort operation
  ' 99  ...  internal error (an error message is displayed disclosing what the error is)
  ' ------------------------------------------------------------------------------------------------------------
  '
  On Error GoTo Bubble_Sort_Vector_Error
  '
  Dim SData() As Variant                                    ' temporary  array  for Bubble_Sort(), to-be-sorted
                                                            ' data
  Dim RSort() As Variant                                    ' array to be returned to caller when Replace is
                                                            ' ... False
  Dim OSort(0 To 0, 0 To 1)                                 ' sort options for Bubble_Sort()
  Dim Entries As Long                                       ' how many values will be sorted, short by one
                                                            '
                                                            '
  Dim LB_Target As Long                                     ' Target's LB
  Dim UB_Target As Long                                     ' Target's UB
  '
  Dim Fail As Variant                                       ' error?
  Dim I As Long
  Dim J As Long
  '
  ' ------------------------------------------------------------------------------------------------------------
  ' Initialization
  ' ------------------------------------------------------------------------------------------------------------
  '
  If Not IsArray(Target) Then                               ' check Target()
     '
     Bubble_Sort_Vector = CVErr(1)                          ' Target is not an array
     Exit Function
     '
  End If
  '
  On Error Resume Next
  Err.Clear
  LB_Target = LBound(Target, 2)
  If Err.Number = 0 Then
     '
     Bubble_Sort_Vector = CVErr(2)                          ' Target() is not a vector
     Exit Function
     '
  End If
  On Error GoTo Bubble_Sort_Vector_Error
  Err.Clear
  '
  LB_Target = LBound(Target, 1)
  UB_Target = UBound(Target, 1)
  '
  '
  '
  If LB_Target = UB_Target Then
     '
     ' nothing needs done
     '
     If Replace Then
        '
        Bubble_Sort_Vector = Array()
        '
     Else
        '
        Bubble_Sort_Vector = Array(LB_Target)
        '
     End If
     Exit Function
     '
  End If
  Entries = UB_Target - LB_Target                           ' how many values will be sorted short by one
  '
  '
  '
  If ((Comparison <> vbTextCompare) And (Comparison <> vbBinaryCompare)) Then Comparison = vbTextCompare
  '
  ' Bubble_Sort() will check data types in Target()
  '
  ' ************************************************************************************************************
  '   MAIN - MAIN - MAIN - MAIN - MAIN - MAIN - MAIN - MAIN - MAIN - MAIN - MAIN - MAIN - MAIN - MAIN - MAIN
  ' ************************************************************************************************************
  '
  ReDim SData(0 To Entries, 0 To 1)                         ' Bubble_Sort() array argument
  ReDim RSort(0 To Entries)                                 ' to-be-returned array
  '
  ' Set up to-be-sorted data for Bubble_Sort()
  '
  J = 0
  For I = LB_Target To UB_Target
      '
      SData(J, 0) = Target(I)
      SData(J, 1) = J
      J = J + 1
      '
  Next I
  '
  OSort(0, 0) = Descending
  OSort(0, 1) = Comparison
  '
  '
  Fail = Bubble_Sort(SData(), OSort())                      ' perform the sort operation
  '
  ' ERROR
  '
  If IsError(Fail) Then
     '
     Select _
       Case CInt(Fail)
            '
       Case 6
            Fail = CVErr(4)                                 ' sort key has an invalid data type
            '
       Case 7
            Fail = CVErr(5)                                 ' sort key has values of intermixed data types
            '
       Case Else
            Fail = CVErr(9)                                 ' shouldn't happen: internal error
            '
     End Select
     '
     Bubble_Sort_Vector = Fail                              ' throw the error
     Exit Function
     '
  End If
  '
  ' SUCCESS
  '
  If Replace Then
     '
     ' Replace Target() cells with ordered data
     '
     I = LB_Target
     '
     For J = 0 To Entries
         '
         Target(I + J) = SData(SData(J, 1), 0)
         '
     Next J
     Bubble_Sort_Vector = Array()                               ' success
     Exit Function
     '
  End If
  '
  ' Leave Target() cells untouched and return an array telling the order in which Target() cells must be
  ' retrieved in order to have Target() cell values fetched in the requested sort order
  '
  For I = 0 To Entries
      '
      RSort(I) = SData(I, 1) + LB_Target
      '
  Next
  Bubble_Sort_Vector = RSort                                    '
  Exit Function
  '
  '
Bubble_Sort_Vector_Error:
  '
  Dim S As String
  '
  S = ""
  S = S & "Bubble_Sort_Vector()" & Chr(13) & Chr(13)
  S = S & "Error (" & Err.Number & "): " & Err.Description
  '
  If Erl <> 0 Then S = S & " at line " & Erl
  S = S & Chr(13)
  '
  I = MsgBox(S, vbOKOnly Or vbCritical, "INTERNAL ERROR")
  '
  Bubble_Sort_Vector = CVErr(99)
  Exit Function
  '
  '
End Function
'
'===============================================================================================================
'
'===============================================================================================================
'
Public Function Bubble_Sort(ByRef SData() As Variant, ByRef Sort_Options() As Variant) As Variant
  '
  ' ------------------------------------------------------------------------------------------------------------
  ' SData  stands for  "Sort  Data". The SData() array hardly will  be the  actual Target  array  being sorted.
  ' SData() must have only the sort keys to be used in the sort operation, plus an additional bit  of data (see
  ' next). It is up to the caller to set up SData() this way.
  '
  ' All  sort keys must have a  valid,  acceptable  data type. The sort operation will  fail if any sort key is
  ' found  to be  of  any of  the  following data  types:  vbEmpty,  vbNull,  vbObject,  vbError, vbDataObject,
  ' vbUserDefinedType, vbArray.  ALL  sort  key VALUES must share a  same compatible  data type, else  the sort
  ' operation will fail. Data types "byte", "integer", "long", "single", "double", "decimal", and "boolean" are
  ' considered to be  "numeric  data types"  and they  CAN be intermixed in a  sort key. Other than those, data
  ' types "currency", "date", and "string" are allowed for sort  keys although they can NOT be  intermixed in a
  ' sort key.
  '
  ' SData() is  expected to be  a "(0 To  N-1,  0 To K)" array, where "N" is the number of  rows in  the actual
  ' to-be-sorted Target array.  SData()'s D1 (1st dimension) must have as many rows  as the actual to-be-sorted
  ' Target array has. The number of sort keys is equal to "K"; that is, "K" must be equal to the number of sort
  ' keys plus one because "LBound(SData, 2) = 0". There  is no clear upper  limit to  the number  of  sort keys
  ' (depends on the environment's resources, e.g. available memory). The 1st (main) sort key is expected  to be
  ' in  SData()'s  D2 subscript 0 - that is, SData(N,  0). The  2nd (secondary)  sort key  is expected to be in
  ' SData()'s D2 subscript 1  - that is, SData(N, 1). And so on, up to SData()'s "Kth" sort key, subscript "K -
  ' 1".
  '
  ' SData()'s  D2 subscript  "K" - that is,  SData(N, K) - is not used for sorting comparisons. It should store
  ' the actual Target array's subscript for  the array's row  from which the provided sort keys were retrieved.
  ' When the sort operation  is finished, SData(N, K) will tell the  order in which  the actual  Target array's
  ' rows must  be retrieved  in  order to have  the actual  Target  array's data retrieved in the required sort
  ' order. That  is, SData()'s to-be-sorted data (rows containing sort keys) NEVER is moved around in the array
  ' (which remains  untouched / unchanged), EXCEPT for the cells  / entries in SData(N, K)  which change places
  ' during the sort operation.
  '
  ' Sort_Options() must be an array  else  the sort operation will fail. The  Sort_Options()  array is meant to
  ' provide sorting options (ascending  /  descending sort  order + comparison method) for each  and every sort
  ' key. Is is expected to de  declared as "Sort_Options(0 To K -1, 0 To 1)". If LBound(Sort_Options, 1) is not
  ' equal to zero,  the  sort operation will fail.  Although  UBound(Sort_Options, 2) should be equal  to 1, no
  ' error will be raised when it is not. Array Sort_Options() should have an entry for each and  every sort key
  ' (e.g. Sort_Options(1) is expected to have  sort options for the  2nd sort key), but if  declaration of sort
  ' key options is missing from Sort_Options() for any sort key, then default options will be used (that is, an
  ' empty Sort_Options() array can be provided if default sort options suffice for the sort keys).
  '
  ' Sort_Options(X - 1, 0) defines the sort direction (ascending = False; descending = True) for the "Xth" sort
  ' key; if the  option is missing or Sort_Options(X - 1, 0) is not a boolean value, a default  "ascending sort
  ' direction" will be  used. Sort_Options(X - 1,  1) defines the comparison  method  for the "Xth" sort  key -
  ' either  binary comparison (vbBinaryCompare =  0) OR text comparison  (vbTextCompare = 1);  if the option is
  ' missing or  the option is neither vbBinaryCompare nor vbTextCompare, then a default "vbTextCompare" will be
  ' used.
  '
  ' When  the Bubble_Sort() function successfully finishes the sort  operation,  True will be  returned  to the
  ' caller, and SData(N, K)  will tell the order that shall be used to retrieve  the actual Target array's rows
  ' in a sorted / ordered way (ascending / descending  as per the provided sorting specs). That is, to retrieve
  ' data  from  the  Target  array  in  a  sorted  /  ordered  way, Target  array's rows  must be  addressed as
  ' Target(SData(N, K)), N ranging from LBound(Target, 1) to UBound(Target, 1).
  '
  ' In case of errors, Error(N) will be returned having one of the following error codes:
  ' 1  ...  SData is not an array
  ' 2  ...  SData()'s lower bound is not zero
  ' 3  ...  SData must be at least (N, 0 To 1)
  ' 4  ...  Sort_Options is not an array
  ' 5  ...  Sort_Options()'s lower bound is not zero
  ' 6  ...  a sort key has an invalid data type
  ' 7  ...  a sort key has values of intermixed data types
  ' 9  ...  internal error (an error message is displayed disclosing what the error is)
  ' ------------------------------------------------------------------------------------------------------------
  '
  Const cfg_Bubble_Sort_Alt = True                          ' if True, use the alternate Bubble Sort algotrithm
                                                            ' else use the standard one
  '
  ' ------------------------------------------------------------------------------------------------------------
  '
  On Error GoTo Bubble_Sort_Error                           ' enable default runtime error handling
  '
  Dim Keys As Integer                                       ' number of sort keys
  '
                                                            ' one of these for each sort key:
  Dim SO() As Variant                                       ' - SO(N,0): True if descendig sort order
                                                            ' - SO(N,1): compare method (vbCompareMethod)
                                                            ' - SO(N,2): True if sort key is numeric
  '
  Dim Base_Type As String                                   ' reference data type for a sort key
  Dim This_Type As String                                   ' data type for an instance of a sort key
  '
  Dim I As Long                                             '
  Dim X As Long                                             '
  Dim S As String                                           '
  '
  ' ------------------------------------------------------------------------------------------------------------
  ' Check arguments up
  ' ------------------------------------------------------------------------------------------------------------
  '
  If Not IsArray(SData) Then                                ' check SData()
     '
     Bubble_Sort = CVErr(1)
     Exit Function
     '
  End If
  '
  '
  I = LBound(SData, 1)
  If I <> 0 Then
     '
     Bubble_Sort = CVErr(2)                                 ' bad SData()'s D1 lower bound
     Exit Function
     '
  End If
  '
  I = UBound(SData, 2)
  If I < 1 Then
     '
     Bubble_Sort = CVErr(3)                                 ' bad SData()'s D1 upper bound
     Exit Function
     '
  End If
  '
  ' ------------------------------------------------------------------------------------------------------------
  '
  If Not IsArray(Sort_Options) Then                         ' check Sort_Options()
     '
     Bubble_Sort = CVErr(4)
     Exit Function
     '
  End If
  '
  '
  On Error Resume Next
  Err.Clear
  '
  I = LBound(Sort_Options, 1)
  If I <> 0 Then
     '
     Bubble_Sort = CVErr(5)                                 ' bad Sort_Options()'s lower bound
     Exit Function
     '
  End If
  On Error GoTo Bubble_Sort_Error                           ' enable default runtime error handling
  '
  ' ------------------------------------------------------------------------------------------------------------
  ' Initialization
  ' ------------------------------------------------------------------------------------------------------------
  '
  Keys = UBound(SData, 2) - 1                               ' number of sort keys short by one
  ReDim SO(0 To Keys, 0 To 2)                               ' Bubble_Sort_Worker()'s sort options array
  '
  On Error Resume Next                                      '
  Err.Clear                                                 '
  '
  For I = 0 To Keys
      '
      ' Get all sort keys' options & data type
      '
      SO(I, 0) = Sort_Options(I, 0)                         ' ascending / descending sort key
      If Err.Number <> 0 Then
         '
         ' Not found: use default
         '
         SO(I, 0) = False                                   ' ascending sort order
         '
      End If
      Err.Clear
      If IsEmpty(SO(I, 0)) Then SO(I, 0) = False
      '
      If VarType(SO(I, 0)) <> vbBoolean Then
         '
         ' Bad data type, not boolean: use default
         '
         SO(I, 0) = False                                   ' ascending sort order
         '
      End If
      '
      '
      '
      SO(I, 1) = Sort_Options(I, 1)                         ' comparison method
      If Err.Number <> 0 Then
         '
         ' Not found: use default
         '
         SO(I, 1) = vbTextCompare
         '
      End If
      Err.Clear
      If IsEmpty(SO(I, 1)) Then SO(I, 1) = vbTextCompare
      '
      If ((SO(I, 1) <> vbTextCompare) And (SO(I, 1) <> vbBinaryCompare)) Then
         '
         ' Bad compare method: use default
         '
         SO(I, 1) = vbTextCompare
         '
      End If
      On Error GoTo Bubble_Sort_Error                       ' enable default runtime error handling
      '
      '
      '
      Base_Type = LCase(TypeName(SData(0, I)))              ' base data type for current sort key
      '
      If InStr(Base_Type, "()") > 0 Then
         '
         ' Array
         '
         Bubble_Sort = CVErr(6)                             ' can not sort those
         Exit Function
         '
      End If
      '
      ' Check for allowed data types
      '
      Select _
        Case Base_Type
             '
        Case "byte", "integer", "long", "single", "double", "decimal", "boolean"
             '
             SO(I, 2) = True                                ' a numeric sort key
             Base_Type = "numeric"                          ' generic type name for all numeric sort keys
             '
        Case "currency", "date"
             '
             SO(I, 2) = True                                ' a numeric sort key
             '
        Case "string"
             '
             SO(I, 2) = False                               ' not a numeric sort key
             '
        Case Else
             '
             Bubble_Sort = CVErr(6)                         ' can not sort those
             Exit Function
             '
      End Select
      '
      '
      ' Variant allows mixed data types; make sure all key values for this sort key share the same data type
      '
      '
      For X = 1 To UBound(SData, 1)
          '
          This_Type = LCase(TypeName(SData(X, I)))          ' data type for this sort key instance
          '
          If InStr(This_Type, "()") > 0 Then
             '
             ' Array
             '
             Bubble_Sort = CVErr(7)                         ' sort key has values of mixed data types
             Exit Function
             '
          End If
          '
          '
          Select _
            Case This_Type
                 '
            Case "byte", "integer", "long", "single", "double", "decimal", "boolean"
                 '
                 This_Type = "numeric"                      ' generic type name for all numeric sort keys
                 '
            Case "currency", "date"
                 '
                 ' an allowed numeric sort key
                 '
            Case "string"
                 '
                 ' an allowed not numeric sort key
                 '
            Case Else
                 '
                 Bubble_Sort = CVErr(7)                     ' sort key has values of mixed data types
                 Exit Function
                 '
          End Select
          '
          '
          If This_Type <> Base_Type Then
             '
             Bubble_Sort = CVErr(7)                         ' sort key has values of mixed data types
             Exit Function
             '
          End If
          '
      Next X ' >>> For X = 1 To UBound(SData, 1)
      '
  Next I ' >>> For I = 0 To Keys
  '
  '
  ' ************************************************************************************************************
  '   MAIN - MAIN - MAIN - MAIN - MAIN - MAIN - MAIN - MAIN - MAIN - MAIN - MAIN - MAIN - MAIN - MAIN - MAIN
  ' ************************************************************************************************************
  '
  If cfg_Bubble_Sort_Alt Then
     '
     Bubble_Sort = Alt_Bubble_Sort_Worker(SData(), 0, UBound(SData, 1), SO())
     '
  Else
     '
     Bubble_Sort = Std_Bubble_Sort_Worker(SData(), 0, UBound(SData, 1), SO())
     '
  End If
  Exit Function
  '
  '
Bubble_Sort_Error:
  '
  S = ""
  S = S & "Bubble_Sort()" & Chr(13) & Chr(13)
  S = S & "Error (" & Err.Number & "): " & Err.Description
  '
  If Erl <> 0 Then S = S & " at line " & Erl
  S = S & Chr(13)
  '
  I = MsgBox(S, vbOKOnly Or vbCritical, "INTERNAL ERROR")
  '
  Bubble_Sort = CVErr(9)
  Exit Function
  '
End Function
'
'===============================================================================================================
'                                   ALTERNATE Bubble Sort algorithm
'
' This is an alternate implementation of the classic  "Bubble Sort"  algorithm. This alternate  algorithm works
' pretty much the same way the classic  "Bubble Sort" algorithm does. It aims  at improving the classic "Bubble
' Sort" algorithm performance; therefore from now on it is dubbed "Smart Bubble Sort".
'
' The classic "Bubble Sort" algorithm repeatedly goes FORWARD over the to-be-sorted array from start to finish,
' swapping adjacent array  elements when the values in those  adjacent  array elements  are "out of order". The
' sort task is over when the  to-be-sorted array  is  scanned from  start to  finish and  no array elements are
' swapped.
'
' The "Smart  Bubble Sort" algorithm  also  goes FORWARD  over the to-be-sorted array from start to finish, but
' just once. When the values  in back-to-back array elements are found to be "out of order", the array elements
' are swapped  in the  array; then  the algorithm  stops  moving  FORWARD over the  array  and it starts moving
' BACKWARDS, comparing and  swapping array  elements as required until the "swapped down" value is fit into its
' right ordering place among the preceding already sorted values.  As  soon as  the "swapped down" value is fit
' into its right ordering place, the algorithm  stops moving  BACKWARDS over the array and restarts the FORWARD
' move at the last "swapped up" array element. The sort task is over when the last array element is reached.
'
' Be the to-be-sorted array "Dim Data(0 To N)" where N = 6. Be "I" the subscript for an array element "Data(I)"
' that is being compared to its next neighbor element "Data(J)" (where "J = I + 1"). Be "D" the array subscript
' for a "swapped down" array element ("0 <=  D <  = N - 1"), and be "U" the array  subscript for a "swapped up"
' array element ("1 <= U <= N"). Be array Data() unsorted values
'
' +---+---+---+---+---+---+---+
' | 1 | 3 | 5 | 2 | 6 | 7 | 4 |   unsorted data to be ascendlingly sorted
' +---+---+---+---+---+---+---+
'   0   1   2   3   4   5   6
'
' The  algorithm goes  forward  from subscript "0" to subscript "2" and no array elements  swapping takes place
' because values "1", "3", and "5" are  already in an ascending order. When "Data(2)" is compared to "Data(3)",
' those  elements'  values are swapped  because "Data(2) >  Data(3)".  At this point: "D = 2", "U =  3",  and 3
' comparisons have been made.
'
' +---+---+---+---+---+---+---+
' | 1 | 3 | 2 | 5 | 6 | 7 | 4 |
' +---+---+---+---+---+---+---+  before: I = 0; J = 1
'   0   1   2   3   4   5   6    after:  I = 2; J = 3; D = 2; U = 3
'
' Now the  algorithm  starts  moving backwards comparing "Data(2)" to "Data(1)", and have them  swapped because
' "Data(1) > Data(2)".
'
' +---+---+---+---+---+---+---+
' | 1 | 2 | 3 | 5 | 6 | 7 | 4 |
' +---+---+---+---+---+---+---+  before: I = 1; J = 2
'   0   1   2   3   4   5   6    after:  D = 1; U = 3
'
' Next,  "Data(1)" is  compared to  "Data(0)" and NO swap occurs because "Data(1) > Data(0)",  meaning that the
' "swapped  down"  value  ("2")  has  been  moved  into  its  right spot  among  already sorted  data  (using 2
' comparisons).
'
' +---+---+---+---+---+---+---+
' | 1 | 2 | 3 | 5 | 6 | 7 | 4 |
' +---+---+---+---+---+---+---+  before: I = 0; J = 1
'   0   1   2   3   4   5   6    after:  D = 1; U = 3
'
' Then  the algorithm restarts moving forward at the  "swapped up" array element "Data(U)" (where "U  = 3"). No
' array elements swapping  occurs  because values  "5",  "6",  and  "7"  are  already in ascending  order. When
' "Data(5)" is compared to "Data(6)", those  elements'  values are swapped because "Data(5) > Data(6)". At this
' point: "D = 5", "U = 6", and 3 additional comparisons have been made.
'
' +---+---+---+---+---+---+---+
' | 1 | 2 | 3 | 5 | 6 | 4 | 7 |
' +---+---+---+---+---+---+---+  before: I = 3; J = I + 1 = 4
'   0   1   2   3   4   5   6    after:  I = 5; J = 6; D = 5; U = 6
'
' Now the  algorithm  starts  moving backwards comparing "Data(4)" to "Data(5)", and have them  swapped because
' "Data(4) > Data(5)".
'
' +---+---+---+---+---+---+---+
' | 1 | 2 | 3 | 5 | 4 | 6 | 7 |
' +---+---+---+---+---+---+---+  before: I = 4; J = 5
'   0   1   2   3   4   5   6    after:  D = 4; U = 6
'
' "Data(3)" is compared to "Data(4)", and the array elements are swapped because "Data(3) > Data(4)".
'
' +---+---+---+---+---+---+---+
' | 1 | 2 | 3 | 4 | 5 | 6 | 7 |
' +---+---+---+---+---+---+---+  before: I = 3; J = 4
'   0   1   2   3   4   5   6    after:  D = 3; U = 6
'
' Next,  "Data(2)" is  compared to  "Data(3)" and NO swap occurs because "Data(3) > Data(2)",  meaning that the
' "swapped  down"  value  ("4")  has  been  moved  into  its  right spot  among  already sorted  data  (using 3
' comparisons).
'
' +---+---+---+---+---+---+---+
' | 1 | 2 | 3 | 4 | 5 | 6 | 7 |
' +---+---+---+---+---+---+---+  before: I = 2; J = 3
'   0   1   2   3   4   5   6    after:  D = 3; U = 6
'
' Then   the algorithm restarts  moving forward at the  "swapped  up" array element "Data(U)" (where "U  = 6").
' "Data(6)" is  the  last array  element,  and the sort task is  finished using 11  comparisons ("3 forward + 2
' backwards + 3 forward + 3 backwards = 11").
'
' --------------------------------------------------------------------------------------------------------------
'
' If the classic "Bubble Sort" algorithm  were used  instead to sort  the  same  data  set, there  would  be 18
' comparisons to have the data sorted:
'
' +---+---+---+---+---+---+---+
' | 1 | 3 | 5 | 2 | 6 | 7 | 4 |   unsorted data to be ascendlingly sorted
' +---+---+---+---+---+---+---+
'   0   1   2   3   4   5   6
'
' +---+---+---+---+---+---+---+
' | 1 | 3 | 2 | 5 | 6 | 4 | 7 |   I = 0; J = 0 To (N - I - 1 = 5)
' +---+---+---+---+---+---+---+   swapped 2x3, 5x6: 6 comparisons
'   0   1   2   3   4   5   6
'
' +---+---+---+---+---+---+---+
' | 1 | 2 | 3 | 5 | 4 | 6 | 7 |   I = 1; J = 0 To (N - I - 1 = 4)
' +---+---+---+---+---+---+---+   swapped 4x5: 5 comparisons
'   0   1   2   3   4   5   6
'
' +---+---+---+---+---+---+---+
' | 1 | 2 | 3 | 4 | 5 | 6 | 7 |   I = 2; J = 0 To (N - I - 1 = 3)
' +---+---+---+---+---+---+---+   swapped 3x4: 4 comparisons
'   0   1   2   3   4   5   6
'
' +---+---+---+---+---+---+---+
' | 1 | 2 | 3 | 4 | 5 | 6 | 7 |   I = 3; J = 0 To (N - I - 1 = 2)
' +---+---+---+---+---+---+---+   NO SWAP / SORTED: 3 comparisons
'   0   1   2   3   4   5   6
'
' --------------------------------------------------------------------------------------------------------------
'
' BOTTOM LINE: according  to  some  benchmarking  carried  out  on   both  algorithms  uaing  randomly  ordered
'              to-be-sorted data  sets,  "Smart Bubble Sort" performs twice as fast than classic "Bubble Sort".
'              Performance of both is the same in a "best case scenario" (that is, already sorted data)  and in
'              a "worst case scenario" (that is, reversely ordered data).
'
'===============================================================================================================
'
Private Function Alt_Bubble_Sort_Worker(ByRef SData() As Variant, ByRef LB As Long, ByRef UB As Long, _
                                        ByRef SO() As Variant) As Variant
  '
  ' ------------------------------------------------------------------------------------------------------------
  ' In case of errors, Error(N) will be returned having one of the following error codes:
  ' 9  ...  internal error (an error message is displayed disclosing what the error is)
  ' ------------------------------------------------------------------------------------------------------------
  '
  On Error GoTo Alt_Bubble_Sort_Worker_Error                ' enable default runtime error handling
  '
  Dim KP As Long                                            ' SData()'s subscript for "data pointers"
  Dim XB As Long                                            ' SData()'s cursor subscript
  Dim RB As Long                                            ' SData()'s subscript
                                                            '
  Dim Retry As Boolean                                      '
  Dim Swap As Boolean                                       '
  Dim Backwards As Boolean                                  '
  '
  '
  Dim I As Long                                             '
  Dim K As Long                                             '
  Dim X As Long                                             '
  '
  ' ------------------------------------------------------------------------------------------------------------
  ' Quit right away if there's nothing to be worked on
  ' ------------------------------------------------------------------------------------------------------------
  '
  Alt_Bubble_Sort_Worker = True                             ' anything but Error()
  If LB = UB Then Exit Function                             ' nothing needs done
  If LB > UB Then Exit Function                             ' something is not quite right
  '
  ' ------------------------------------------------------------------------------------------------------------
  ' Initialization
  ' ------------------------------------------------------------------------------------------------------------
  '
  KP = UBound(SData, 2)                                     ' SData()'s subscript for "data pointers"; also the
                                                            ' number of sort keys plus one
  '
  ' ************************************************************************************************************
  '   MAIN - MAIN - MAIN - MAIN - MAIN - MAIN - MAIN - MAIN - MAIN - MAIN - MAIN - MAIN - MAIN - MAIN - MAIN
  ' ************************************************************************************************************
  '
  Backwards = False
  Retry = True
  XB = LB
  '
  Do While Retry
     '
     Retry = False
     '
     Do While ((XB < UB) And (Not Retry))
        '
        Swap = False
        K = 0                                               ' 1st sort key
        I = XB + 1
        '
        Do While (K < KP)
           '
           If SO(K, 0) Then
              '
              ' Descending order - ASSERT: I = XB + 1
              '
              Select _
                Case SO(K, 2)
                Case True
                     '
                     If SData(SData(I, KP), K) > SData(SData(XB, KP), K) Then
                        Swap = True
                     Else
                        If SData(SData(I, KP), K) < SData(SData(XB, KP), K) Then
                           K = KP - 1                       ' done with the current data
                        End If
                     End If
                     '
                Case Else
                     '
                     X = StrComp(SData(SData(I, KP), K), SData(SData(XB, KP), K))
                     If X > 0 Then
                        Swap = True
                     Else
                        If X < 0 Then
                           K = KP - 1                       ' done with the current data
                        End If
                     End If
                     '
              End Select
              '
           Else ' >>> If SO(K, 0)
              '
              ' Ascending order - ASSERT: I = XB + 1
              '
              Select _
                Case SO(K, 2)
                Case True
                     '
                     If SData(SData(XB, KP), K) > SData(SData(I, KP), K) Then
                        Swap = True
                     Else
                        If SData(SData(XB, KP), K) < SData(SData(I, KP), K) Then
                           K = KP - 1                       ' done with the current data
                        End If
                     End If
                     '
                Case Else
                     '
                     X = StrComp(SData(SData(XB, KP), K), SData(SData(I, KP), K))
                     If X > 0 Then
                        Swap = True
                     Else
                        If X < 0 Then
                           K = KP - 1                       ' done with the current data
                        End If
                     End If
                     '
              End Select
              '
           End If ' >>> Else SO(K, 0)
           '
           '
           '
           If Backwards Then
              '
              ' ------------------------------------------------------------------------------------------------
              ' Moving backward to sort data
              ' ------------------------------------------------------------------------------------------------
              '
              If Swap Then
                 '
                 For K = 0 To KP - 1
                     '
                     X = SData(I, KP)                       ' save to move to SData(XB)
                     SData(I, KP) = SData(XB, KP)           ' replace
                     SData(XB, KP) = X                      ' swap
                     '
                 Next K
                 K = KP                                     ' done with the current data
                 '
                 Retry = True                               ' array not sorted as yet; restart sorting
                 XB = XB - 1                                ' restart sorting  at the  entry  preceding the one
                                                            ' just swapped
                 If XB < LB Then
                    '
                    XB = RB                                 ' backwards move is over, restart moving forward
                    Backwards = False
                    '
                 End If
                 '
              Else ' >>> If Swap
                 '
                 K = K + 1                                  ' next sort key
                 '
                 If K = KP Then
                    '
                    XB = RB                                 ' backwards move is over, restart moving forward
                    Backwards = False
                    '
                 End If
                 '
              End If ' >>> Else Swap
              '
           Else ' >>> If Backwards
              '
              ' ------------------------------------------------------------------------------------------------
              ' Moving forward to sort data
              ' ------------------------------------------------------------------------------------------------
              '
              If Swap Then
                 '
                 For K = 0 To KP - 1
                     '
                     X = SData(I, KP)                       ' save to move to SData(XB)
                     SData(I, KP) = SData(XB, KP)           ' replace
                     SData(XB, KP) = X                      ' swap
                     '
                     RB = I                                 ' will restart moving forward here
                     Backwards = True                       ' move the just swapped data to its right place  in
                                                            ' the set of already sorted data
                     '
                 Next K
                 K = KP                                     ' done with the current data
                 '
                 Retry = True                               ' array not sorted as yet; restart sorting
                 XB = XB - 1                                ' restart sorting  at the  entry  preceding the one
                                                            ' just swapped
                 If XB < LB Then
                    '
                    XB = LB
                    Backwards = False                       ' in case it was set to True
                    '
                 End If
                 '
              Else ' >>> If Swap
                 '
                 K = K + 1                                  ' next sort key
                 '
                 If K = KP Then
                    '
                    XB = XB + 1
                    '
                 End If
                 '
              End If ' >>> Else Swap
              '
           End If ' >>> Else Backwards
           '
        Loop ' >>> Do While (K < KP)
        '
     Loop ' >>> Do While ((XB < UB) And (Not Retry))
     '
  Loop ' >>> Do While Retry
  '
  ' Done
  '
  Exit Function
  '
  '
Alt_Bubble_Sort_Worker_Error:
  '
  Dim S As String
  '
  S = ""
  S = S & "Bubble_Sort_Worker()" & Chr(13) & Chr(13)
  S = S & "Error (" & Err.Number & "): " & Err.Description
  '
  If Erl <> 0 Then S = S & " at line " & Erl
  S = S & Chr(13)
  '
  I = MsgBox(S, vbOKOnly Or vbCritical, "INTERNAL ERROR")
  Alt_Bubble_Sort_Worker = CVErr(9)                         ' unexpected internal error
  '
  Exit Function
  '
End Function
'
'===============================================================================================================
' STANDARD Bubble Sort algorithm: https://www.interviewkickstart.com/blogs/learn/bubble-sort
'===============================================================================================================
'
Private Function Std_Bubble_Sort_Worker(ByRef SData() As Variant, ByRef LB As Long, ByRef UB As Long, _
                                        ByRef SO() As Variant) As Variant
  '
  ' ------------------------------------------------------------------------------------------------------------
  ' In case of errors, Error(N) will be returned having one of the following error codes:
  ' 9  ...  internal error (an error message is displayed disclosing what the error is)
  ' ------------------------------------------------------------------------------------------------------------
  '
  On Error GoTo Std_Bubble_Sort_Worker_Error                ' enable default runtime error handling
  '
  Dim KP As Long                                            ' SData()'s subscript for "data pointers"
  Dim YB As Long                                            ' SData()'s subscript ceiling value
  Dim ZB As Long                                            ' SData()'s subscript ceiling value
                                                            '
  Dim Retry As Boolean                                      '
  Dim Swap As Boolean                                       '
  '
  '
  Dim I As Long                                             '
  Dim J As Long                                             '
  Dim K As Long                                             '
  Dim N As Long                                             '
  Dim X As Long                                             '
  '
  ' ------------------------------------------------------------------------------------------------------------
  ' Quit right away if there's nothing to be worked on
  ' ------------------------------------------------------------------------------------------------------------
  '
  Std_Bubble_Sort_Worker = True                             ' anything but Error()
  If LB = UB Then Exit Function                             ' nothing needs done
  If LB > UB Then Exit Function                             ' something is not quite right
  '
  ' ------------------------------------------------------------------------------------------------------------
  ' Initialization
  ' ------------------------------------------------------------------------------------------------------------
  '
  KP = UBound(SData, 2)                                     ' SData()'s subscript for "data pointers"; also the
                                                            ' number of sort keys plus one
  ZB = (UB - LB) + LB + 1                                   ' all SData()'s subscripts will be less than this
  '
  ' ************************************************************************************************************
  '   MAIN - MAIN - MAIN - MAIN - MAIN - MAIN - MAIN - MAIN - MAIN - MAIN - MAIN - MAIN - MAIN - MAIN - MAIN
  ' ************************************************************************************************************
  '
  Retry = True
  I = LB
  '
  Do While Retry
     '
     Do While ((I < ZB) And Retry)
        '
        J = LB
        YB = ZB - I - 1
        Retry = False
        '
        Do While (J < YB)
           '
           Swap = False
           K = 0                                            ' 1st sort key
           N = J + 1
           '
           Do While (K < KP)                                ' all sort keys
              '
              If SO(K, 0) Then
                 '
                 ' Descending order - ASSERT: N = J + 1
                 '
                 Select _
                   Case SO(K, 2)
                   Case True
                        '
                        If SData(SData(N, KP), K) > SData(SData(J, KP), K) Then
                           Swap = True
                        Else
                           If SData(SData(N, KP), K) < SData(SData(J, KP), K) Then
                              K = KP                        ' done with the current data
                           End If
                        End If
                        '
                   Case Else
                        '
                        X = StrComp(SData(SData(N, KP), K), SData(SData(J, KP), K))
                        If X > 0 Then
                           Swap = True
                        Else
                           If X < 0 Then
                              K = KP                        ' done with the current data
                           End If
                        End If
                        '
                 End Select
                 '
              Else ' >>> If SO(K, 0)
                 '
                 ' Ascending order - ASSERT: N = J + 1
                 '
                 Select _
                   Case SO(K, 2)
                   Case True
                        '
                        If SData(SData(J, KP), K) > SData(SData(N, KP), K) Then
                           Swap = True
                        Else
                           If SData(SData(J, KP), K) < SData(SData(N, KP), K) Then
                              K = KP                        ' done with the current data
                           End If
                        End If
                        '
                   Case Else
                        '
                        X = StrComp(SData(SData(J, KP), K), SData(SData(N, KP), K))
                        If X > 0 Then
                           Swap = True
                        Else
                           If X < 0 Then
                              K = KP                        ' done with the current data
                           End If
                        End If
                        '
                 End Select
                 '
              End If ' >>> Else SO(K, 0)
              '
              '
              '
              If Swap Then
                 '
                 For K = 0 To KP - 1
                     '
                     X = SData(N, KP)                       ' save to move to SData(XB)
                     SData(N, KP) = SData(J, KP)            ' replace
                     SData(J, KP) = X                       ' swap
                     '
                 Next K
                 Retry = True                               ' array not sorted as yet
                 '
              End If ' >>> If Swap
              '
              K = K + 1                                     ' next sort key
              '
           Loop ' >>> Do While (K < KP)
           '
           J = J + 1                                        ' done with the current data
           '
        Loop ' >>> Do While (J < YB)
        '
        I = I + 1                                           ' done with the current data
        '
     Loop ' >>> Do While ((I < ZB) And Retry)
     '
  Loop ' >>> Do While Retry
  '
  ' Done
  '
  Exit Function
  '
  '
Std_Bubble_Sort_Worker_Error:
  '
  Dim S As String
  '
  S = ""
  S = S & "Bubble_Sort_Worker()" & Chr(13) & Chr(13)
  S = S & "Error (" & Err.Number & "): " & Err.Description
  '
  If Erl <> 0 Then S = S & " at line " & Erl
  S = S & Chr(13)
  '
  I = MsgBox(S, vbOKOnly Or vbCritical, "INTERNAL ERROR")
  Std_Bubble_Sort_Worker = CVErr(9)                         ' unexpected internal error
  '
  Exit Function
  '
End Function
'
'===============================================================================================================
'
'===============================================================================================================
'
