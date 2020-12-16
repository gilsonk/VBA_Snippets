' Compare two workbooks with the same structure, and sort differences and missing entries
' Useful for DB extracts
' Gilson, Kevin

' Need 3 worksheets:
' * Equal
' * Differences
' * Missing
Option Explicit

Dim wb As Workbook
Dim ws_Eq As Worksheet
Dim ws_Df As Worksheet
Dim ws_Ms As Worksheet
Dim str_TmpSht

Private Type FullFile
    Wkb As Workbook
    Sht As Worksheet
    rng As Range
    Hnd As Boolean
    Rows As Long
    Cols As Long
End Type

Private Sub SetVar()
    Set wb = Application.ThisWorkbook
    Set ws_Eq = wb.Worksheets("Equal")
    Set ws_Df = wb.Worksheets("Differences")
    Set ws_Ms = wb.Worksheets("Missing")
    str_TmpSht = vbNull
End Sub

Sub CleanAll()
    Call SetVar
    ws_Eq.Cells.ClearContents
    ws_Eq.Cells.ClearFormats
    ws_Df.Cells.ClearContents
    ws_Df.Cells.ClearFormats
    ws_Ms.Cells.ClearContents
    ws_Ms.Cells.ClearFormats
End Sub

Private Sub SetSheet()
    str_TmpSht = Application.CommandBars.ActionControl.Tag
End Sub

Private Function OpenWb(str_Title As String) As FullFile
    Call SetVar
    Dim ff_File As FullFile
    
    Dim str_FPath As String
    Dim wb_Opn As Workbook
    Dim ws_Slt As Worksheet
    Dim rng_Slt As Range
    Dim bl_Head As Boolean
    Dim lng_Rows As Long
    Dim lng_Cols As Long
    
    Dim cb_CmdBar As CommandBar
    Dim bb_BarBtn As CommandBarButton
    
    'Ask to choose file
    With Application.FileDialog(msoFileDialogFilePicker)
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx; *.xlsm; *.xls; *.xlsb", 1
        .AllowMultiSelect = False
        .Title = str_Title
        If .Show = True Then
            str_FPath = .SelectedItems.Item(1)
        End If
    End With
    
    'Open file
    Set wb_Opn = Application.Workbooks.Open(str_FPath)
    
    'Ask to choose Worksheet
    On Error Resume Next
        Application.CommandBars("Register").Delete
    On Error GoTo 0
    
    Set cb_CmdBar = Application.CommandBars.Add("Register", msoBarPopup)
    For Each ws_Slt In wb_Opn.Worksheets
        Set bb_BarBtn = cb_CmdBar.Controls.Add
        With bb_BarBtn
            .Caption = ws_Slt.Name
            .Tag = ws_Slt.Name
            .Style = msoButtonCaption
            .OnAction = "SetSheet"
        End With
    Next ws_Slt
    cb_CmdBar.ShowPopup
    Set ws_Slt = wb_Opn.Worksheets(str_TmpSht)
    
    On Error Resume Next
        Application.CommandBars("Register").Delete
    On Error GoTo 0
    
    'Ask which range of Columns should for the Look-up key
    Set rng_Slt = Application.InputBox("Please select the columns used as a look-up key", "Range", Type:=8)
    
    'Ask whether there is a header or not
    If MsgBox("Should the first row be considered as a header?", vbYesNo, "Header") = vbYes Then
        bl_Head = True
    Else
        bl_Head = False
    End If
    
    'Get height
    lng_Rows = ws_Slt.Range("A:A").Cells.SpecialCells(xlCellTypeConstants).Count
    
    'Get width
    lng_Cols = ws_Slt.Range("1:1").Cells.SpecialCells(xlCellTypeConstants).Count
    
    'Return info
    With ff_File
        Set .Wkb = wb_Opn
        Set .Sht = ws_Slt
        Set .rng = rng_Slt
        .Hnd = bl_Head
        .Rows = lng_Rows
        .Cols = lng_Cols
    End With
    OpenWb = ff_File
End Function

Private Sub CloseWb(wb_Del As Workbook)
    Application.DisplayAlerts = False
        wb_Del.Close savechanges:=False
    Application.DisplayAlerts = True
End Sub

Private Function LookupKey(rng As Range, ws As Worksheet, rw As Long) As String
    Call SetVar
    Dim i As Long
    Dim str_Key As String
    
    For i = rng.Column To rng.Columns(rng.Columns.Count).Column
        If i <> rng.Column Then
            str_Key = str_Key & "_"
        End If
        str_Key = str_Key & ws.Cells(rw, i).Value
    Next i
    LookupKey = str_Key
End Function

Sub Main()
    Call SetVar
    Dim ff_One As FullFile
    Dim ff_Two As FullFile
    
    ff_One = OpenWb("Please select the first file for the comparison")
    ff_Two = OpenWb("Please select the second file for the comparison")
    
    MsgBox "Both files will be compared based on their lookup keys." & _
        vbNewLine & "Files should have the same width, and columns should be in the same order"
    'Test width
    If ff_One.Cols <> ff_Two.Cols Then
        MsgBox "Error! Files doesn't have the same width."
    'Test lookup-key size
    ElseIf ff_One.rng.Columns.Count <> ff_Two.rng.Columns.Count Then
        MsgBox "Error! Lookup-keys doesn't have the same size."
    Else
        Application.ScreenUpdating = False
        
        'Clear and initialize worksheets
        Dim lng_RowEq As Long
        Dim lng_RowDf As Long
        Dim lng_RowMs As Long
        lng_RowEq = 1
        lng_RowDf = 1
        lng_RowMs = 1
        
        Call CleanAll
        
        'Generate Keys
        Dim i As Long
        Dim j As Long
        Dim str_Key As String
        Dim dic_LkpOne As Object
        Dim dic_LkpTwo As Object
        
        Set dic_LkpOne = CreateObject("Scripting.Dictionary")
        Set dic_LkpTwo = CreateObject("Scripting.Dictionary")
        
        str_Key = vbNull
        
        'First File
        For i = 1 To ff_One.Rows
            If Not ((ff_One.Hnd = True) And (i = 1)) Then
                str_Key = LookupKey(ff_One.rng, ff_One.Sht, i)
                On Error GoTo err_Dic
                    dic_LkpOne.Add str_Key, i
                On Error GoTo 0
            End If
        Next i
        
        'Second File
        For i = 1 To ff_Two.Rows
            If Not ((ff_Two.Hnd = True) And (i = 1)) Then
                str_Key = LookupKey(ff_Two.rng, ff_Two.Sht, i)
                On Error GoTo err_Dic
                    dic_LkpTwo.Add str_Key, i
                On Error GoTo 0
            End If
        Next i
        
        'Keep everything or only show differences
        Dim bl_All As Boolean
        If MsgBox("Should we keep every data?" & vbNewLine & "If no, only differences will be kept.", vbYesNo, "Header") = vbYes Then
            bl_All = True
        Else
            bl_All = False
        End If
        
        'Add Header
        If ff_One.Hnd Then
            ws_Ms.Cells(lng_RowMs, 1).Value = "Workbook"
            For i = 1 To ff_One.Cols
                ws_Eq.Cells(lng_RowEq, i).Value = ff_One.Sht.Cells(1, i).Value
                ws_Df.Cells(lng_RowDf, i).Value = ff_One.Sht.Cells(1, i).Value
                ws_Df.Cells(lng_RowDf, i + ff_One.Cols - ff_One.Rng.Columns.Count).Value = ff_One.Sht.Cells(1, i).Value
                ws_Ms.Cells(lng_RowMs, i + 1).Value = ff_One.Sht.Cells(1, i).Value
            Next i
            lng_RowEq = lng_RowEq + 1
            lng_RowDf = lng_RowDf + 1
            lng_RowMs = lng_RowMs + 1
        ElseIf ff_Two.Hnd Then
            ws_Ms.Cells(lng_RowMs, 1).Value = "Workbook"
            For i = 1 To ff_Two.Cols
                ws_Eq.Cells(lng_RowEq, i).Value = ff_Two.Sht.Cells(1, i).Value
                ws_Df.Cells(lng_RowDf, i).Value = ff_Two.Sht.Cells(1, i).Value
                ws_Df.Cells(lng_RowDf, i + ff_Two.Cols - ff_Two.Rng.Columns.Count).Value = ff_Two.Sht.Cells(1, i).Value
                ws_Ms.Cells(lng_RowMs, i + 1).Value = ff_Two.Sht.Cells(1, i).Value
            Next i
            lng_RowEq = lng_RowEq + 1
            lng_RowDf = lng_RowDf + 1
            lng_RowMs = lng_RowMs + 1
        End If
        
        'Iterate through Dictionary #1
        Dim vr_One As Variant
        For Each vr_One In dic_LkpOne
            'Comparison
            If dic_LkpTwo.Exists(vr_One) Then
                Dim bl_Diff As Boolean
                bl_Diff = False
                'Test for differences
                For i = 1 To ff_One.Cols
                    If (ff_One.Sht.Cells(dic_LkpOne(vr_One), i).Value <> ff_Two.Sht.Cells(dic_LkpTwo(vr_One), i).Value) Then
                        bl_Diff = True
                    End If
                Next i
                
                'Differences
                If bl_Diff Then
                    For i = 1 To ff_One.Cols
                        ws_Df.Cells(lng_RowDf, i).Value = ff_One.Sht.Cells(dic_LkpOne(vr_One), i).Value
                    Next i
                    For i = ff_One.Cols + 1 To ff_One.Cols + ff_Two.Cols
                        ws_Df.Cells(lng_RowDf, i).Value = ff_Two.Sht.Cells(dic_LkpTwo(vr_One), i - ff_Two.Cols + ff_Two.rng.Columns.Count).Value
                    Next i
                    
                    lng_RowDf = lng_RowDf + 1
                Else
                    If bl_All Then
                        For i = 1 To ff_One.Cols
                            ws_Eq.Cells(lng_RowEq, i).Value = ff_One.Sht.Cells(dic_LkpOne(vr_One), i).Value
                        Next i
                    End If
                    
                    lng_RowEq = lng_RowEq + 1
                End If
            'Missing
            Else
                ws_Ms.Cells(lng_RowMs, 1).Value = ff_One.Wkb.Name
                For i = 1 To ff_One.Cols
                    ws_Ms.Cells(lng_RowMs, i + 1).Value = ff_One.Sht.Cells(dic_LkpOne(vr_One), i).Value
                Next i
            
                lng_RowMs = lng_RowMs + 1
            End If
        Next vr_One
        
        'Iterate through Dictionary #2 for missings only
        Dim vr_Two As Variant
        For Each vr_Two In dic_LkpTwo
            If Not (dic_LkpOne.Exists(vr_Two)) Then
                ws_Ms.Cells(lng_RowMs, 1).Value = ff_Two.Wkb.Name
                For i = 1 To ff_Two.Cols
                    ws_Ms.Cells(lng_RowMs, i + 1).Value = ff_Two.Sht.Cells(dic_LkpTwo(vr_Two), i).Value
                Next i
                
                lng_RowMs = lng_RowMs + 1
            End If
        Next vr_Two
        
        'Close files
        Call CloseWb(ff_One.Wkb)
        Call CloseWb(ff_Two.Wkb)
        
        Application.ScreenUpdating = True
    End If
    
    Exit Sub

'Error handling
err_Dic:
    MsgBox "Error! Duplicate lookup key." & vbNewLine & "Key: " & str_Key
    Application.ScreenUpdating = True
End Sub