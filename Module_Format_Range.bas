Attribute VB_Name = "Module_Format_Range"
'***************************************************************************************************
'* Module_Format_Range (Module)
'*
'* @index ----------------------------------------------------------------------------------------->
'*
'*  Public  Type        Min_And_Max     : User-defined type for storing minimum and maximum values.
'*  Public  Funtion     Get_Max         : Get the maximum value of a row or column.
'*  Public  Funtion     Is_Array_Empty  : Returns True if the array has no elements.
'*  Public  Sub         Set_Format      : Format according to array instructions.
'*  Private Sub         Set_Format_Code : Auxiliary processing of Set_Format.
'*
'* @history --------------------------------------------------------------------------------------->
'*
'* 2024/09/11 (Ver.1.0.0) Create as new module. [Kyosuke Homma, https://github.com/kyosukehomma]
'*
'***************************************************************************************************
Option Explicit
'===================================================================================================

'***************************************************************************************************
'* Min_And_Max (Type)
'*
'* @description : Variable type to store the minimum and maximum values (e.g., row/column indices)
'*                used for identifying the range in the worksheet.
'***************************************************************************************************
Public Type Min_And_Max
    Min As Long
    Max As Long
End Type

'***************************************************************************************************
'* Get_Max
'*
'* @description : Retrieves the maximum row or column index from the specified base cell. If
'*                Is_Matrix = True, it explores the non-specified direction (row or column) to find
'*                the maximum index in the specified direction.
'*
'* @return      : The maximum row or column index as a Long value.
'*
'* @argument    : Row_Or_Column = A String specifying "Row" or "Column" to define the exploration direction.
'*                Base_Cell     = Range object that acts as the starting point for the exploration.
'*                Is_Matrix     = (Optional) Boolean. If True, it performs a matrix search.
'***************************************************************************************************
Public Function Get_Max( _
    ByRef Row_Or_Column As String, _
    ByRef Base_Cell As Range, _
    Optional ByRef Is_Matrix As Boolean = False) As Long    ' This function can be quite useful on its own

    '===============================================================================================
    ' If Is_Matrix = True, starting from Base_Cell, it explores the non-specified direction
    ' to find the maximum index in the specified direction.
    '
    ' Specified Direction = The row or column specified by Row_Or_Column
    ' Non-Specified Direction = The row or column not specified by Row_Or_Column
    '===============================================================================================
    
    Dim max_value     As Long       ' Holds the maximum index
    Dim current_value As Long       ' Holds the current index being explored
    Dim i   As Long
    Dim rng As Range
    Dim ws  As Worksheet
    
    Set ws = Base_Cell.Parent
    
    If Is_Matrix = False Then       ' If not exploring the entire matrix (Is_Matrix = False)
        
        Select Case Row_Or_Column   ' Get the last row or column in the specified direction
            
            Case "Row"              ' If exploring rows, get the last row in the column
                i = Base_Cell.Column
                Set rng = ws.Cells(ws.Rows.Count, i)
                max_value = rng.End(xlUp).Row
            
            Case "Column"           ' If exploring columns, get the last column in the row
                i = Base_Cell.Row
                Set rng = ws.Cells(i, ws.Columns.Count)
                max_value = rng.End(xlToLeft).Column
        End Select
    
    Else                            ' If exploring the entire matrix (Is_Matrix = True)
        
        Select Case Row_Or_Column   ' Based on the specified direction,
                                    ' explore the non-specified direction to get the last row or column
            ' --------------------------------------------------------------------------------------
            Case "Row"      ' If exploring rows
            
                ' First get the last column of the base cell and check the last row of each column
                max_value = Get_Max("Column", Base_Cell, False)
                For i = 1 To max_value
                
                    ' Get the last row of each column
                    Set rng = ws.Cells(1, i)
                    current_value = Get_Max("Row", rng, False)
                    
                    ' Update if the obtained row index exceeds the maximum value
                    If current_value > max_value Then
                        max_value = current_value
                    End If
                Next i
            ' --------------------------------------------------------------------------------------
            Case "Column"   ' If exploring columns
            
                ' First get the last row of the base cell and check the last column of each row
                max_value = Get_Max("Row", Base_Cell, False)
                For i = 1 To max_value
                
                    ' Get the last column of each row
                    Set rng = ws.Cells(i, 1)
                    current_value = Get_Max("Column", rng, False)
                    
                    ' Update if the obtained column index exceeds the maximum value
                    If current_value > max_value Then
                        max_value = current_value
                    End If
                Next i
            ' --------------------------------------------------------------------------------------
        End Select
    End If
    
    Get_Max = max_value    ' Return the final maximum index
    Set rng = Nothing
    Set ws = Nothing
End Function

'***************************************************************************************************
'* Is_Array_Empty
'*
'* @description : Checks if the specified array is empty or if an error occurs when retrieving
'*                the upper bound of the array. This function handles errors to determine if the
'*                array is empty or not.
'*
'* @return      : True if the array is empty or an error occurs; False if the array is not empty.
'*
'* @argument    : arr        = The array to check.
'*                dimension  = (Optional) The dimension of the array to check. Default is 1 (the first dimension).
'***************************************************************************************************
Public Function Is_Array_Empty( _
    ByRef arr As Variant, _
    Optional ByRef dimension As Long = 1) As Boolean    ' This function can be quite useful on its own
    
    If Not IsArray(arr) Then
        Is_Array_Empty = True
        Exit Function
    End If

    On Error Resume Next
    If UBound(arr, dimension) < LBound(arr, dimension) Then
        Is_Array_Empty = True
    Else
        Is_Array_Empty = False
    End If
    
    On Error GoTo 0
End Function

'***************************************************************************************************
'* Set_Format
'*
'* @description : Applies format settings (number format, horizontal alignment, and indent level)
'*                to a specified range based on keywords found in the header row.
'*
'* @return      : None
'*
'* @argument    : Base_Cell    = Range object representing the starting cell for format application.
'*                Keywords     = Array of keywords to match against the table header.
'*                Format_Codes = Array of corresponding format codes to apply for each keyword.
'*                Horiz_Aligns = Array of horizontal alignment values for each keyword.
'*                Indent_Ctrls = Array of indent levels for each keyword.
'***************************************************************************************************
Public Sub Set_Format( _
    ByRef Base_Cell As Range, _
    ByRef Keywords() As String, _
    ByRef Format_Codes() As String, _
    ByRef Horiz_Aligns() As Long, _
    ByRef Indent_Ctrls() As Long)

    Dim i       As Long
    Dim j       As Long
    Dim ty_col  As Min_And_Max
    Dim ty_row  As Min_And_Max
    Dim thead   As Range
    Dim rng     As Range
    Dim str     As String
    Dim ws      As Worksheet
    Dim found   As Boolean
    
    Select Case True    ' If any of them is True, the process is terminated.
        Case _
            Is_Array_Empty(Keywords), _
            Is_Array_Empty(Format_Codes), _
            Is_Array_Empty(Horiz_Aligns), _
            Is_Array_Empty(Indent_Ctrls), _
            UBound(Keywords, 1) <> UBound(Format_Codes, 1)
            Exit Sub
    End Select
    
    Set ws = Base_Cell.Parent
    
    ty_col.Min = Base_Cell.Column
    ty_row.Min = Base_Cell.Row
    
    ty_col.Max = Get_Max("Column", Base_Cell)       ' Get the maximum column index from the base cell
    ty_row.Max = Get_Max("Row", Base_Cell, True)    ' Get the maximum row index from the column direction by exploring from the base cell
    
    For i = 1 To ty_col.Max
        Set thead = ws.Cells(ty_row.Min, i)
        Set rng = ws.Range(thead.Offset(1, 0), ws.Cells(ty_row.Max, i))
        Let str = thead.Value
        
        found = False       ' Check if there are any matches with the keywords
        For j = LBound(Keywords, 1) To UBound(Keywords, 1)

            If InStr(str, Keywords(j)) > 0 Then
                Call Set_Format_Code(rng, str, Keywords(j), Format_Codes(j), Horiz_Aligns(j), Indent_Ctrls(j))
                found = True
                Exit For
            End If
        Next j
        
        If Not found Then   ' Apply default format if no matches were found
            Call Set_Format_Code(rng, str)
        End If
    Next i
    
    Set rng = Nothing
    Set thead = Nothing
    Set ws = Nothing
End Sub

'***************************************************************************************************
'* Set_Format_Code
'*
'* @description : Applies a specific format (number format, horizontal alignment, and indent level)
'*                to the specified target range based on the table header and keyword.
'*
'* @return      : None
'*
'* @argument    : Target       = Range object to apply the formatting to.
'*                Table_Header = String representing the table header to match against the keyword.
'*                Keyword      = (Optional) String representing the keyword to match in the header.
'*                Format_Code  = (Optional) String representing the number format to apply.
'*                Horiz_Align  = (Optional) Long representing the horizontal alignment value.
'*                Indent_Ctrl  = (Optional) Long representing the indent level to apply.
'*
'* @caution     : If the system locale is not Japan, change the stipulated value of the Format_Code
'*                argument from ÅgG/ïWèÄÅh to ÅgGeneralÅh or something similar.
'*
'***************************************************************************************************
Private Sub Set_Format_Code( _
    ByRef Target As Range, _
    ByRef Table_Header As String, _
    Optional ByRef Keyword As String = "", _
    Optional ByRef Format_Code As String = "G/ïWèÄ", _
    Optional ByRef Horiz_Align As Long = 0, _
    Optional ByRef Indent_Ctrl As Long = 0)
    
    If Keyword <> "" Then
        If InStr(Table_Header, Keyword) > 0 Then
            Target.NumberFormatLocal = Format_Code
        End If
    Else
        Target.NumberFormatLocal = Format_Code
    End If

    ' Set horizontal alignment
    If Horiz_Align <> 0 Then
        Target.HorizontalAlignment = Horiz_Align
    End If
    
    ' Set indent level (only for left or right alignment)
    Select Case Horiz_Align
        Case xlLeft, xlRight
            Target.IndentLevel = Indent_Ctrl
    End Select
End Sub

'----------------------------------------<< End of Source >>----------------------------------------
