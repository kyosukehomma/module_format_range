Attribute VB_Name = "Sample"
'***************************************************************************************************
'* Sample (Module)
'*
'* @history --------------------------------------------------------------------------------------->
'* 2024/09/11 (Ver.1.0.0) Create as new module. [Kyosuke Homma, https://github.com/kyosukehomma]
'***************************************************************************************************
Option Explicit
'===================================================================================================

'***************************************************************************************************
'* Main         : How to use Mod_Format_Range.
'*
'* @description : Initializes test data such as target words, format codes, horizontal alignments,
'*                and indent controls, and applies the format settings to the specified range.
'* @return      : None
'* @argument    : None
'***************************************************************************************************
Public Sub Main()
    Dim rng As Range
    Dim words(0 To 3) As String
    Dim codes(0 To 3) As String
    Dim haligns(0 To 3) As Long
    Dim indents(0 To 3) As Long
    
    Call Main_sub(rng, words, codes, haligns, indents)      ' Cf. This Module
    
    Call Set_Format(rng, words, codes, haligns, indents)    ' Cf. Module_Format_Range Module
    
    Set rng = Nothing
    Erase words
    Erase codes
    Erase haligns
    Erase indents
End Sub

'***************************************************************************************************
'* Main_sub
'*
'* @description : Initializes test data.
'* @return      : None
'* @argument    : rng     = Range object representing the starting cell for format application.
'*                words   = Array of keywords to match against the table header.
'*                codes   = Array of corresponding format codes to apply for each keyword.
'*                haligns = Array of horizontal alignment values for each keyword.
'*                Indents = Array of indent levels for each keyword.
'***************************************************************************************************
Private Sub Main_sub( _
    ByRef rng As Range, _
    ByRef words() As String, _
    ByRef codes() As String, _
    ByRef haligns() As Long, _
    ByRef indents() As Long)

    Set rng = Sheet1.Range("A1")
    
    words(0) = "SALARY"
    words(1) = "BIRTHDAY"
    words(2) = "ID"
    words(3) = "NAME"
                
    codes(0) = "$#,##0"
    codes(1) = "yyyy/mm/dd"
    codes(2) = "0000"
    codes(3) = "@"
    
    haligns(0) = xlRight
    haligns(1) = xlCenter
    haligns(2) = xlCenter
    haligns(3) = xlLeft
                
    indents(0) = 1
    indents(1) = 0
    indents(2) = 0
    indents(3) = 1

End Sub

'***************************************************************************************************
'* Revert       : Revert to before the "Main" macro execution.
'*
'* @description : Initializes test data such as target words, format codes, horizontal alignments,
'*                and indent controls, and applies the format settings to the specified range.
'* @return      : None
'* @argument    : None
'***************************************************************************************************
Public Sub Revert()
    Dim rng As Range
    Dim words(0 To 3) As String
    Dim codes(0 To 3) As String
    Dim haligns(0 To 3) As Long
    Dim indents(0 To 3) As Long
    
    Call Revert_sub(rng, words, codes, haligns, indents)    ' Cf. This Module
    
    Call Set_Format(rng, words, codes, haligns, indents)    ' Cf. Module_Format_Range Module
    
    Set rng = Nothing
    Erase words
    Erase codes
    Erase haligns
    Erase indents
End Sub

'***************************************************************************************************
'* Revert_sub
'*
'* @description : Initializes test data.
'* @return      : None
'* @argument    : rng     = Range object representing the starting cell for format application.
'*                words   = Array of keywords to match against the table header.
'*                codes   = Array of corresponding format codes to apply for each keyword.
'*                haligns = Array of horizontal alignment values for each keyword.
'*                Indents = Array of indent levels for each keyword.
'***************************************************************************************************
Private Sub Revert_sub( _
    ByRef rng As Range, _
    ByRef words() As String, _
    ByRef codes() As String, _
    ByRef haligns() As Long, _
    ByRef indents() As Long)

    Set rng = Sheet1.Range("A1")
    
    words(0) = "SALARY"
    words(1) = "BIRTHDAY"
    words(2) = "ID"
    words(3) = "NAME"
                
    codes(0) = "G/標準"
    codes(1) = "G/標準"
    codes(2) = "G/標準"
    codes(3) = "G/標準"
    
    haligns(0) = xlLeft
    haligns(1) = xlLeft
    haligns(2) = xlLeft
    haligns(3) = xlLeft
                
    indents(0) = 0
    indents(1) = 0
    indents(2) = 0
    indents(3) = 0

End Sub
'----------------------------------------<< End of Source >>----------------------------------------
