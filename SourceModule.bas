Attribute VB_Name = "SearchWKBooks"
Sub SearchWKBooks()
On Error Resume Next
' Variable declarations
Dim oWS_Result As Worksheet, oWS_Source As Worksheet
Dim sPath_DIR As String
Dim rngSearch As Range: Set rngSearch = ThisWorkbook.Worksheets("Initialization").Range("C3:C22")
Dim iCount As Single, iLoadingDone As Single, iFileCount As Single, iFileCountCurrent As Single
' Get dir with excel files
With Application.FileDialog(msoFileDialogFolderPicker)
    .InitialFileName = ""
    Dim iReturn As Single: iReturn = .Show
    If iReturn = -1 Then
        ' Dir selected, start processing
        OptimizeVBA True
        ufProgress.LabelProgress.Width = 0
        ufProgress.Show
        sPath_DIR = .SelectedItems(1) & "\"
        iFileCount = CountFilesInFolder(sPath_DIR, "*xl??")
    Else
        ' Nothing selected, stop processing
        Unload ufProgress
        OptimizeVBA False
        Exit Sub
    End If
End With
' Delete previous result sheet
Dim sThisWS As Worksheet
For Each sThisWS In Application.ActiveWorkbook.Worksheets
    If sThisWS.Name <> "Initialization" Then
        sThisWS.Delete
    End If
Next
' Create new result sheet
Set oWS_Result = Sheets.Add
oWS_Result.Name = "Rezultat"
oWS_Result.Range("A1") = "File"
oWS_Result.Range("B1") = "Sheet"
oWS_Result.Range("C1") = "Cell Address"
oWS_Result.Range("D1") = "Link"
oWS_Result.Range("E1") = "Value"
With oWS_Result.Range("A1:E1")
    .AutoFilter
    .Font.Bold = True
    .Font.Size = 16
End With
' Reset counters
iCount = 0
iFileCountCurrent = 0
' Iterate through excel files inside the selected dir
sPath_File = Dir(sPath_DIR)
Do Until sPath_File = ""
    ' Skip system files
    If sPath_File = "." Or sPath_File = ".." Then
    Else
        ' Process only xls, xlsx or xlsm files
        If Right(sPath_File, 3) = "xls" Or Right(sPath_File, 4) = "xlsx" Or Right(sPath_File, 4) = "xlsm" Then
            On Error Resume Next
            iFileCountCurrent = iFileCountCurrent + 1
            ' Open workbook
            Workbooks.Open Filename:=sPath_DIR & sPath_File, UpdateLinks:=False, ReadOnly:=True, IgnoreReadOnlyRecommended:=True, CorruptLoad:=xlExtractData
            If Err.Number > 0 Then
                oWS_Result.Range("A2").Offset(iCount, 0).Value = sPath_File
                oWS_Result.Range("B2").Offset(iCount, 0).Value = "Nu s-a putut deschide fisierul"
                iCount = iCount + 1
            Else
                On Error GoTo 0
                ' Iterate through each sheet in the opened workbook
                For Each oWS_Source In ActiveWorkbook.Worksheets
                    ' Iterate through each search criteria
                    For Each oCell In rngSearch
                        If oCell.Value <> "" Then
                            ' Execute the Find method
                            Set oFindResult = oWS_Source.Cells.Find(oCell, LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext)
                            If Not oFindResult Is Nothing Then
                                Dim sAddress As String: sAddress = oFindResult.Address
                                Do ' Set findings into the result sheet
                                    oWS_Result.Range("A2").Offset(iCount, 0).Value = sPath_File
                                    oWS_Result.Range("B2").Offset(iCount, 0).Value = oWS_Source.Name
                                    oWS_Result.Range("C2").Offset(iCount, 0).Value = oFindResult.Address
                                    oWS_Result.Range("E2").Offset(iCount, 0).Value = oFindResult.Value
                                    oWS_Result.Hyperlinks.Add _
                                        Anchor:=oWS_Result.Range("D2").Offset(iCount, 0), _
                                        Address:=sPath_DIR & sPath_File, _
                                        SubAddress:=oWS_Source.Name & "!" & oFindResult.Address, TextToDisplay:="Link"
                                    iCount = iCount + 1
                                    Set oFindResult = oWS_Source.Cells.FindNext(oFindResult)
                                Loop While Not oFindResult Is Nothing And oFindResult.Address <> sAddress
                            End If
                        End If
                    Next oCell
                Next oWS_Source
            End If
            ' Close workbook
            Workbooks(sPath_File).Close False
            On Error GoTo 0
        End If
    End If
    sPath_File = Dir
    ' Set loading progress bar
    iLoadingDone = iFileCountCurrent / iFileCount
    With ufProgress
        .LabelCaption.Caption = "Progress: " & iFileCountCurrent & " / " & iFileCount
        .LabelProgress.Width = iLoadingDone * (.FrameProgress.Width)
    End With
    ufProgress.Repaint
Loop
' Refresh the result sheet and end processing
Cells.EntireColumn.AutoFit
Unload ufProgress
OptimizeVBA False
End Sub

Sub OptimizeVBA(bMode As Boolean)
    Dim oWS As Worksheet
    With Application
        If bMode Then
            .Calculation = xlCalculationManual
            .ScreenUpdating = False
            .EnableAnimations = False
            .DisplayAlerts = False
            For Each oWS In ActiveWorkbook.Worksheets
                oWS.DisplayPageBreaks = False
            Next oWS
        Else
            .Calculation = xlCalculationAutomatic
            .ScreenUpdating = True
            .EnableAnimations = True
            .DisplayAlerts = True
        End If
    End With
End Sub

Private Function CountFilesInFolder(sDIR As String, Optional sType As String)
    Dim oFile As Variant
    Dim iCount As Integer
    If Right(sDIR, 1) <> "\" Then sDIR = sDIR & "\"
    oFile = Dir(sDIR & sType)
    While (oFile <> "")
        iCount = iCount + 1
        oFile = Dir
    Wend
    CountFilesInFolder = iCount
End Function
