' This macro compares PBOM against EBOM based on Display Name (column B)
' and validates EID, Part Number, Phase, Quantity, and Status.

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

    ' Ask user to select PBOM file
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Select PBOM File"
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

                ' BF - EID Match
                wsPBOM.Cells(i, "BF").Value = pbomEID
                wsPBOM.Cells(i, "BF").Interior.ColorIndex = xlNone

                ' BG - Part Number Match
                ebomPN = Trim(wsEBOM.Cells(j, "C").Value)
                If pbomStatus = "Replace" Then
                    If pbomProposedPN = ebomPN Then
                        wsPBOM.Cells(i, "BG").Value = pbomProposedPN
                        wsPBOM.Cells(i, "BG").Interior.ColorIndex = xlNone
                    Else
                        wsPBOM.Cells(i, "BG").Value = "Mismatch"
                        wsPBOM.Cells(i, "BG").Interior.Color = RGB(255, 199, 206)
                    End If
                ElseIf pbomStatus = "Accepted" Then
                    If pbomPN = ebomPN Then
                        wsPBOM.Cells(i, "BG").Value = pbomPN
                        wsPBOM.Cells(i, "BG").Interior.ColorIndex = xlNone
                    Else
                        wsPBOM.Cells(i, "BG").Value = "Mismatch"
                        wsPBOM.Cells(i, "BG").Interior.Color = RGB(255, 199, 206)
                    End If
                ElseIf pbomStatus = "CM Remove Level" Then
                    wsPBOM.Cells(i, "BG").Value = "N/A"
                    wsPBOM.Cells(i, "BG").Interior.ColorIndex = xlNone
                End If

                ' BH - Phase Match
                ebomPhase = Trim(wsEBOM.Cells(j, "G").Value)
                If pbomStatus = "Replace" Then
                    If pbomProposedPhase = ebomPhase Then
                        wsPBOM.Cells(i, "BH").Value = pbomProposedPhase
                        wsPBOM.Cells(i, "BH").Interior.ColorIndex = xlNone
                    Else
                        wsPBOM.Cells(i, "BH").Value = "Mismatch"
                        wsPBOM.Cells(i, "BH").Interior.Color = RGB(255, 199, 206)
                    End If
                ElseIf pbomStatus = "Accepted" Then
                    If pbomPhase = ebomPhase Then
                        wsPBOM.Cells(i, "BH").Value = pbomPhase
                        wsPBOM.Cells(i, "BH").Interior.ColorIndex = xlNone
                    Else
                        wsPBOM.Cells(i, "BH").Value = "Mismatch"
                        wsPBOM.Cells(i, "BH").Interior.Color = RGB(255, 199, 206)
                    End If
                End If

                ' BI - Quantity Match
                ebomQty = wsEBOM.Cells(j, "I").Value
                If pbomQty = ebomQty Then
                    wsPBOM.Cells(i, "BI").Value = pbomQty
                    wsPBOM.Cells(i, "BI").Interior.ColorIndex = xlNone
                ElseIf pbomProposedQty = ebomQty And pbomProposedQty <> "" Then
                    wsPBOM.Cells(i, "BI").Value = pbomProposedQty
                    wsPBOM.Cells(i, "BI").Interior.ColorIndex = xlNone
                ElseIf pbomQty = "" And pbomProposedQty = "" Then
                    wsPBOM.Cells(i, "BI").Value = "N/A"
                    wsPBOM.Cells(i, "BI").Interior.ColorIndex = xlNone
                Else
                    wsPBOM.Cells(i, "BI").Value = "Mismatch"
                    wsPBOM.Cells(i, "BI").Interior.Color = RGB(255, 199, 206)
                End If

                ' BJ - Status Match
                ebomStatus = Trim(wsEBOM.Cells(j, "H").Value)
                If pbomStatus = "Replace" Then
                    If pbomProposedStatus = ebomStatus Then
                        wsPBOM.Cells(i, "BJ").Value = pbomProposedStatus
                        wsPBOM.Cells(i, "BJ").Interior.ColorIndex = xlNone
                    Else
                        wsPBOM.Cells(i, "BJ").Value = "Mismatch"
                        wsPBOM.Cells(i, "BJ").Interior.Color = RGB(255, 199, 206)
                    End If
                ElseIf pbomStatus = "Released" Then
                    If pbomStatus = ebomStatus Then
                        wsPBOM.Cells(i, "BJ").Value = pbomStatus
                        wsPBOM.Cells(i, "BJ").Interior.ColorIndex = xlNone
                    Else
                        wsPBOM.Cells(i, "BJ").Value = "Mismatch"
                        wsPBOM.Cells(i, "BJ").Interior.Color = RGB(255, 199, 206)
                    End If
                ElseIf pbomStatus = "CM Remove Level" Then
                    wsPBOM.Cells(i, "BJ").Value = "N/A"
                    wsPBOM.Cells(i, "BJ").Interior.ColorIndex = xlNone
                End If

                Exit For
            End If
        Next j

        If Not foundMatch Then
            wsPBOM.Cells(i, "BF").Value = "Mismatch"
            wsPBOM.Cells(i, "BF").Interior.Color = RGB(255, 199, 206)
            wsPBOM.Cells(i, "BG").Value = "Mismatch"
            wsPBOM.Cells(i, "BG").Interior.Color = RGB(255, 199, 206)
            wsPBOM.Cells(i, "BH").Value = "Mismatch"
            wsPBOM.Cells(i, "BH").Interior.Color = RGB(255, 199, 206)
            wsPBOM.Cells(i, "BI").Value = "Mismatch"
            wsPBOM.Cells(i, "BI").Interior.Color = RGB(255, 199, 206)
            wsPBOM.Cells(i, "BJ").Value = "Mismatch"
            wsPBOM.Cells(i, "BJ").Interior.Color = RGB(255, 199, 206)
        End If
    Next i

    wbEBOM.Close SaveChanges:=False
    Application.ScreenUpdating = True

    MsgBox "PBOM vs EBOM comparison completed.", vbInformation
End Sub
