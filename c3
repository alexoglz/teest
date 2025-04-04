Sub ComparePBOMtoEBOM()
    Dim selectedPBOM As String
    Dim selectedEBOM As String
    Dim wbPBOM As Workbook, wbEBOM As Workbook
    Dim wsPBOM As Worksheet, wsEBOM As Worksheet
    Dim lastRowPBOM As Long, lastRowEBOM As Long
    Dim i As Long, j As Long
    Dim foundMatch As Boolean

    Dim pbomEID As String, ebomEID As String
    Dim pbomPN As String, pbomProposedPN As String, ebomPN As String
    Dim pbomPhase As String, pbomProposedPhase As String, ebomPhase As String
    Dim pbomQty As Variant, pbomProposedQty As Variant, ebomQty As Variant
    Dim pbomStatus As String, pbomProposedStatus As String, ebomStatus As String

    ' Ask user to select the PBOM file (from PBOMCheck)
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Select New_Unit_Release File (PBOM Output)"
        .Filters.Add "Excel Files", "*.xls; *.xlsx; *.xlsm", 1
        If .Show <> -1 Then Exit Sub
        selectedPBOM = .SelectedItems(1)
        UserForm2.Hide
    End With

    ' Ask user to select EBOM file
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Select EBOM File"
        .Filters.Add "Excel Files", "*.xls; *.xlsx; *.xlsm", 1
        If .Show <> -1 Then Exit Sub
        selectedEBOM = .SelectedItems(1)
    End With

    Set wbPBOM = Workbooks.Open(selectedPBOM)
    Set wsPBOM = wbPBOM.Sheets(1)
    Set wbEBOM = Workbooks.Open(selectedEBOM, ReadOnly:=True)
    Set wsEBOM = wbEBOM.Sheets(1)

    lastRowPBOM = wsPBOM.Cells(wsPBOM.Rows.Count, "B").End(xlUp).Row
    lastRowEBOM = wsEBOM.Cells(wsEBOM.Rows.Count, "B").End(xlUp).Row

    Application.ScreenUpdating = False

    ' EBOM Header setup
    wsPBOM.Range("AZ3:BD3").Merge
    With wsPBOM.Range("AZ3")
        .Value = "EBOM"
        .HorizontalAlignment = xlCenter
        .Font.Bold = True
    End With
    wsPBOM.Range("AZ4").Value = "EID"
    wsPBOM.Range("BA4").Value = "PN"
    wsPBOM.Range("BB4").Value = "PHASE"
    wsPBOM.Range("BC4").Value = "QTY"
    wsPBOM.Range("BD4").Value = "STATUS"

    ' Add black separator at column AY
    With wsPBOM.Columns("AY")
        .ColumnWidth = 2
        .Interior.Color = RGB(0, 0, 0)
    End With

    ' Start comparisons row by row
    For i = 6 To lastRowPBOM
        pbomEID = Trim(wsPBOM.Cells(i, "B").Value)
        pbomPN = Trim(wsPBOM.Cells(i, "D").Value)
        pbomProposedPN = Trim(wsPBOM.Cells(i, "M").Value)
        pbomPhase = Trim(wsPBOM.Cells(i, "G").Value)
        pbomProposedPhase = Trim(wsPBOM.Cells(i, "P").Value)
        pbomQty = wsPBOM.Cells(i, "H").Value
        pbomProposedQty = wsPBOM.Cells(i, "Q").Value
        pbomStatus = Trim(wsPBOM.Cells(i, "K").Value)
        pbomProposedStatus = Trim(wsPBOM.Cells(i, "O").Value)

        foundMatch = False

        For j = 5 To lastRowEBOM
            ebomEID = Trim(wsEBOM.Cells(j, "B").Value)

            If pbomEID = ebomEID Then
                foundMatch = True

                ' AZ - EID Match
                wsPBOM.Cells(i, "AZ").Value = pbomEID

                ' BA - Part Number Match
                ebomPN = Trim(wsEBOM.Cells(j, "C").Value)
                If pbomStatus = "Replace" Then
                    If pbomProposedPN = ebomPN Then
                        wsPBOM.Cells(i, "BA").Value = pbomProposedPN
                    Else
                        wsPBOM.Cells(i, "BA").Value = "Mismatch"
                        wsPBOM.Cells(i, "BA").Interior.Color = RGB(255, 199, 206)
                    End If
                ElseIf pbomStatus = "Accepted" Then
                    If pbomPN = ebomPN Then
                        wsPBOM.Cells(i, "BA").Value = pbomPN
                    Else
                        wsPBOM.Cells(i, "BA").Value = "Mismatch"
                        wsPBOM.Cells(i, "BA").Interior.Color = RGB(255, 199, 206)
                    End If
                ElseIf pbomStatus = "CM Remove Level" Then
                    wsPBOM.Cells(i, "BA").Value = "N/A"
                End If

                ' BB - Phase Match
                ebomPhase = Trim(wsEBOM.Cells(j, "G").Value)
                If pbomStatus = "Replace" Then
                    If pbomProposedPhase = ebomPhase Then
                        wsPBOM.Cells(i, "BB").Value = pbomProposedPhase
                    Else
                        wsPBOM.Cells(i, "BB").Value = "Mismatch"
                        wsPBOM.Cells(i, "BB").Interior.Color = RGB(255, 199, 206)
                    End If
                ElseIf pbomStatus = "Accepted" Then
                    If pbomPhase = ebomPhase Then
                        wsPBOM.Cells(i, "BB").Value = pbomPhase
                    Else
                        wsPBOM.Cells(i, "BB").Value = "Mismatch"
                        wsPBOM.Cells(i, "BB").Interior.Color = RGB(255, 199, 206)
                    End If
                End If

                ' BC - Quantity Match
                ebomQty = wsEBOM.Cells(j, "I").Value
                If pbomQty = ebomQty Then
                    wsPBOM.Cells(i, "BC").Value = pbomQty
                ElseIf pbomProposedQty = ebomQty And pbomProposedQty <> "" Then
                    wsPBOM.Cells(i, "BC").Value = pbomProposedQty
                ElseIf pbomQty = "" And pbomProposedQty = "" Then
                    wsPBOM.Cells(i, "BC").Value = "N/A"
                Else
                    wsPBOM.Cells(i, "BC").Value = "Mismatch"
                    wsPBOM.Cells(i, "BC").Interior.Color = RGB(255, 199, 206)
                End If

                ' BD - Status Match
                ebomStatus = Trim(wsEBOM.Cells(j, "H").Value)
                If pbomStatus = "Replace" Then
                    If pbomProposedStatus = ebomStatus Then
                        wsPBOM.Cells(i, "BD").Value = pbomProposedStatus
                    Else
                        wsPBOM.Cells(i, "BD").Value = "Mismatch"
                        wsPBOM.Cells(i, "BD").Interior.Color = RGB(255, 199, 206)
                    End If
                ElseIf pbomStatus = "Released" Then
                    If pbomStatus = ebomStatus Then
                        wsPBOM.Cells(i, "BD").Value = pbomStatus
                    Else
                        wsPBOM.Cells(i, "BD").Value = "Mismatch"
                        wsPBOM.Cells(i, "BD").Interior.Color = RGB(255, 199, 206)
                    End If
                ElseIf pbomStatus = "CM Remove Level" Then
                    wsPBOM.Cells(i, "BD").Value = "N/A"
                End If

                Exit For
            End If
        Next j

        ' No match in EBOM
        If Not foundMatch Then
            wsPBOM.Cells(i, "AZ").Resize(1, 5).Value = "Mismatch"
            wsPBOM.Range(wsPBOM.Cells(i, "AZ"), wsPBOM.Cells(i, "BD")).Interior.Color = RGB(255, 199, 206)
        End If
    Next i

    wbEBOM.Close SaveChanges:=False
    wbPBOM.Save

    Application.ScreenUpdating = True
    MsgBox "PBOM vs EBOM comparison completed and saved.", vbInformation
End Sub
