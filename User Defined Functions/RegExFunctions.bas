' Create 5 UDF functions for RegEx
' Gilson, Kevin

'=VB_RegExMatch(
'    Text:           String to look within
'   ,Pattern:        Pattern to look for
'   ,[Ignore Case]): Ignore case (Optional, False by default)

'=VB_RegExCount(
'    Text:           String to look within
'   ,Pattern:        Pattern to look for
'   ,[Ignore Case]): Ignore case (Optional, False by default)

'=VB_RegExExtract(
'    Text:              String to look within
'   ,Pattern:           Pattern to look for
'   ,[Ignore Case]:     Ignore case (Optional, False by default)
'   ,[Match Index]:     Return Nth Match, start at 1 (Optional)
'   ,[SubMatch Index]): Return Nth SubMatch, start at 1 (Optional)

'=VB_RegExFormat(
'    Text:           String to look within
'   ,Pattern:        Pattern to look for
'   ,[Format]:       RegEx format used to display the extacted Nth groups,
'                    under the form $N (Optional, $0 by default)
'   ,[Ignore Case]:  Ignore case (Optional, False by default)
'   ,[Match Index]): Return Nth Match, start at 1 (Optional)

'=VB_RegExReplace(
'    Text:           String to look within
'   ,Pattern:        Pattern to look for
'   ,Replacement:    Text to replace Pattern with
'   ,[Ignore Case]:  Ignore case (Optional, False by default)
'   ,[Replace All]): Replace all occurences (Optional, True by default)

Option Explicit

Dim wb As Workbook

Private Sub SetVar()
    Set wb = Application.ThisWorkbook
End Sub

'RegEx Match - Boolean
'Look for a pattern within a given string, return TRUE or FALSE
'str_Text: String to look in
'str_Pattern: RegEx pattern
'bl_IgnoreCase: Ignore case - Optional (False by default)
Public Function VB_RegExMatch(ByVal str_Text As String, _
ByVal str_Pattern As String, _
Optional ByVal bl_IgnoreCase As Boolean = False) As Boolean
    On Error GoTo reMatch_ErrVal

    Dim obj_RegEx, obj_Matches As Object
    Set obj_RegEx = CreateObject("VBScript.RegExp")
    Set obj_Matches = Nothing

    Dim bl_Result As Boolean

    With obj_RegEx
        .Global = True
        .MultiLine = True
        .IgnoreCase = bl_IgnoreCase
        .Pattern = str_Pattern
    End With

    bl_Result = IIf(obj_RegEx.Test(str_Text),True,False)

    Set obj_Matches = Nothing
    Set obj_RegEx = Nothing
    VB_RegExMatch = bl_Result
    Exit Function

reMatch_ErrVal:
    Set obj_Matches = Nothing
    Set obj_RegEx = Nothing
    VB_RegExMatch = CVErr(xlErrValue)
End Function

'RegEx Count - Long
'Look for a pattern within a given string, return the number of matches
'str_Text: String to look in
'str_Pattern: RegEx pattern
'bl_IgnoreCase: Ignore case - Optional (False by default)
Public Function VB_RegExCount(ByVal str_Text As String, _
ByVal str_Pattern As String, _
Optional ByVal bl_IgnoreCase As Boolean = False) As Long
    On Error GoTo reCount_ErrVal

    Dim obj_RegEx, obj_Matches As Object
    Set obj_RegEx = CreateObject("VBScript.RegExp")
    Set obj_Matches = Nothing

    Dim lng_Result

    With obj_RegEx
        .Global = True
        .MultiLine = True
        .IgnoreCase = bl_IgnoreCase
        .Pattern = str_Pattern
    End With

    If obj_RegEx.Test(str_Text) Then
        Set obj_Matches = obj_RegEx.Execute(str_Text)
        lng_Result = VBA.CLng(obj_Matches.Count)
    Else
        lng_Result = 0
    End If

    Set obj_Matches = Nothing
    Set obj_RegEx = Nothing
    VB_RegExCount = lng_Result
    Exit Function

reCount_ErrVal:
    Set obj_Matches = Nothing
    Set obj_RegEx = Nothing
    VB_RegExCount = CVErr(xlErrValue)
End Function

'RegEx Extract - String
'Look for a pattern within a given string, extract it
'str_Text: String to look in
'str_Pattern: RegEx pattern
'bl_IgnoreCase: Ignore case - Optional (False by default)
'lng_MatchIndex: Match index - Optional (0 by default, set to -1 for testing)
'lng_SubMatchIndex: SubMatch index - Optional (None by default, set to -1 for testing)
Public Function VB_RegExExtract(ByVal str_Text As String, _
ByVal str_Pattern As String, _
Optional ByVal bl_IgnoreCase As Boolean = False, _
Optional ByVal lng_MatchIndex As Long = -1, _
Optional ByVal lng_SubMatchIndex As Long = -1) As String
    On Error GoTo reExtract_ErrVal

    Dim obj_RegEx, obj_Matches As Object
    Set obj_RegEx = CreateObject("VBScript.RegExp")
    Set obj_Matches = Nothing

    Dim str_Result As String

    With obj_RegEx
        .Global = True
        .MultiLine = True
        .IgnoreCase = bl_IgnoreCase
        .Pattern = str_Pattern
    End With

    If obj_regEx.Test(str_Text) Then
        Set obj_Matches = obj_RegEx.Execute(str_Text)
        If (lng_MatchIndex <> -1) Then
            If (lng_SubMatchIndex <> -1) Then
                str_Result = obj_Matches(lng_MatchIndex - 1).SubMatches(lng_SubMatchIndex - 1)
            Else
                str_Result = obj_Matches(lng_MatchIndex - 1)
            End If
        Else
            str_Result = obj_Matches(0)
        End If
    Else
        GoTo reExtract_ErrNa
    End If

    Set obj_Matches = Nothing
    Set obj_RegEx = Nothing
    VB_RegExExtract = str_Result
    Exit Function

reExtract_ErrNa:
    Set obj_Matches = Nothing
    Set obj_RegEx = Nothing
    VB_RegExExtract = CVErr(xlErrNA)
    Exit Function

reExtract_ErrVal:
    Set obj_Matches = Nothing
    Set obj_RegEx = Nothing
    VB_RegExExtract = CVErr(xlErrValue)
End Function

'RegEx Format - String
'Look for a pattern within a given string, extract it and display it based on a given pattern
'str_Text: String to look in
'str_Pattern: RegEx pattern
'str_Format: Format used to display result - Optional ($0 by default, for the entire match)
'bl_IgnoreCase: Ignore case - Optional (False by default)
'lng_MatchIndex: Match index - Optional (0 by default, set to -1 for testing)
Public Function VB_RegExFormat(ByVal str_Text As String, _
ByVal str_Pattern As String, _
Optional ByVal str_Format As String = "$0", _
Optional ByVal bl_IgnoreCase As Boolean = False, _
Optional ByVal lng_MatchIndex As Long = -1) As String
    On Error GoTo reFormat_ErrVal

    Dim obj_RegEx, obj_Matches As Object
    Set obj_RegEx = CreateObject("VBScript.RegExp")
    Set obj_Matches = Nothing

    Dim str_Result As String

    With obj_RegEx
        .Global = True
        .MultiLine = True
        .IgnoreCase = bl_IgnoreCase
        .Pattern = str_Pattern
    End With

    Dim obj_RegExFormat, obj_RegExReplace As Object
    Set obj_RegExFormat = CreateObject("VBScript.RegExp")
    Set obj_RegExReplace = CreateObject("VBScript.RegExp")

    With obj_RegExFormat
        .Global = True
        .Multiline = True
        .IgnoreCase = False
        .Pattern = "\$(\d+)"
    End With
    With obj_RegExReplace
        .Global = True
        .Multiline = True
        .IgnoreCase = False
    End With

    If (lng_MatchIndex <> -1) Then
        lng_MatchIndex = lng_MatchIndex - 1
    Else
        lng_MatchIndex = 0
    End If

    If obj_regEx.Test(str_Text) Then
        Set obj_Matches = obj_RegEx.Execute(str_Text)

        Dim obj_ReplaceMatches As Object
        Dim vr_ReplaceMatch As Variant

        Set obj_ReplaceMatches = obj_RegExFormat.Execute(str_Format)
        For Each vr_ReplaceMatch In obj_ReplaceMatches
            Dim lng_ReplaceNumber As Long
            lng_ReplaceNumber = vr_ReplaceMatch.SubMatches(0)
            obj_RegExReplace.Pattern = "\$" & lng_ReplaceNumber

            If lng_ReplaceNumber > obj_Matches(lng_MatchIndex).SubMatches.Count Then
                GoTo reFormat_ErrVal
            ElseIf lng_ReplaceNumber = 0 Then
                str_Format = obj_RegExReplace.Replace(str_Format, obj_Matches(lng_MatchIndex).Value)
            Else
                str_Format = obj_RegExReplace.Replace(str_Format, _
                    obj_Matches(lng_MatchIndex).SubMatches(lng_ReplaceNumber - 1))
            End If
        Next vr_ReplaceMatch

        str_Result = str_Format
    Else
        GoTo reFormat_ErrNa
    End If

    Set obj_Matches = Nothing
    Set obj_ReplaceMatches = Nothing
    Set obj_RegEx = Nothing
    Set obj_RegExFormat = Nothing
    Set obj_RegExReplace = Nothing
    VB_RegExFormat = str_Result
    Exit Function

reFormat_ErrNa:
    Set obj_Matches = Nothing
    Set obj_ReplaceMatches = Nothing
    Set obj_RegEx = Nothing
    Set obj_RegExFormat = Nothing
    Set obj_RegExReplace = Nothing
    VB_RegExFormat = CVErr(xlErrNA)
    Exit Function

reFormat_ErrVal:
    Set obj_Matches = Nothing
    Set obj_ReplaceMatches = Nothing
    Set obj_RegEx = Nothing
    Set obj_RegExFormat = Nothing
    Set obj_RegExReplace = Nothing
    VB_RegExFormat = CVErr(xlErrValue)
End Function

'RegEx Replace - String
'Look for a pattern within a given string, replace it with another string
'str_Text: String to look in
'str_Pattern: RegEx pattern
'str_Replace: String to replace the pattern with
'bl_IgnoreCase: Ignore case - Optional (False by default)
'bl_ReplaceAll: Replace all matches - Optional (True by default)
Public Function VB_RegExReplace(ByVal str_Text As String, _
ByVal str_Pattern As String, _
ByVal str_Replace As String, _
Optional ByVal bl_IgnoreCase As Boolean = False, _
Optional ByVal bl_ReplaceAll As Boolean = True) As String
    On Error GoTo reReplace_Err

    Dim obj_RegEx, obj_Matches As Object
    Set obj_RegEx = CreateObject("VBScript.RegExp")
    Set obj_Matches = Nothing

    Dim str_Result As String

    With obj_RegEx
        .Global = bl_ReplaceAll
        .MultiLine = True
        .IgnoreCase = bl_IgnoreCase
        .Pattern = str_Pattern
    End With

    If obj_RegEx.Test(str_Text) Then
        str_Result = obj_RegEx.Replace(str_Text, str_Replace)
    Else
        str_Result = str_Text
    End If

    Set obj_Matches = Nothing
    Set obj_RegEx = Nothing
    VB_RegExReplace = str_Result
    Exit Function

reReplace_Err:
    Set obj_Matches = Nothing
    Set obj_RegEx = Nothing
    VB_RegExReplace = CVErr(xlErrValue)
End Function