Attribute VB_Name = "Module1"
Sub btnlaunch_Click()
UserForm1.Show
End Sub

Sub LineVoltage()

Application.ScreenUpdating = False
Application.EnableEvents = False
Application.Calculation = xlCalculationManual

Worksheets("Cover").Activate

Dim name As String: name = CStr(Cells(5, "L").Value)

Worksheets("Alarm").Activate

Dim kV As String: kV = CStr(Cells(11, "G").Value)

Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")

Dim Fileout As Object: Set Fileout = fso.CreateTextFile("C:\Users\" + Environ("Username") + "\Desktop\scaDAbuilder\LinekV\" + kV + "_" + name + ".txt", True, True)
        Fileout.Writeline name
        Fileout.Close

Application.ScreenUpdating = True
Application.EnableEvents = True
Application.Calculation = xlCalculationAutomatic

End Sub

Sub AORSort()

Application.ScreenUpdating = False
Application.EnableEvents = False
Application.Calculation = xlCalculationManual

Worksheets("Cover").Activate

Dim AOR As String: AOR = CStr(Cells(10, "D").Value)

ActiveWorkbook.SaveAs Filename:="C:\Users\" + Environ("Username") + "\Desktop\scaDAbuilder\AOR\" + AOR + "\" + ActiveWorkbook.name + ".xlsm" _
                                , FileFormat:=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False

Application.ScreenUpdating = True
Application.EnableEvents = True
Application.Calculation = xlCalculationAutomatic

End Sub

Sub ToDo(i As Long)

Application.ScreenUpdating = False
Application.EnableEvents = False
Application.Calculation = xlCalculationManual

Worksheets("Cover").Activate

Dim AOR As String: AOR = CStr(Cells(10, "D").Value)
Dim RTU As String: RTU = CStr(Cells(5, "L").Value)
Dim devtype As String: devtype = CStr(Cells(4, "L").Value)
Dim system As String

    If AOR = "DART" Then
        system = "DART"
    Else
        system = "PROD"
    End If

If Dir("C:\Users\" + Environ("Username") + "\Desktop\scaDAbuilder\To Do List\To Do List.xlsx") <> "" Then
    
    Workbooks.Open Filename:="C:\Users\" + Environ("Username") + "\Desktop\scaDAbuilder\To Do List\To Do List.xlsx"
    
    Cells(2 + i, "A").Value = RTU
    Cells(2 + i, "B").Value = devtype
    Cells(2 + i, "D").Value = system
    Cells(2 + i, "F").Value = "Not Started"
    Cells(2 + i, "J").Value = "TRUE"
    Cells(2 + i, "O").Value = AOR
    Cells(2 + i, "P").Value = "Item"

End If

If Dir("C:\Users\" + Environ("Username") + "\Desktop\scaDAbuilder\To Do List\To Do List.xlsx") = "" Then

    Workbooks.Add
    
    Cells(1, "A").Value = "Device Id"
    Cells(1, "B").Value = "Device Type"
    Cells(1, "C").Value = "Description"
    Cells(1, "D").Value = "System"
    Cells(1, "E").Value = "SNOW Ticket"
    Cells(1, "F").Value = "Status"
    Cells(1, "G").Value = "Modeler"
    Cells(1, "H").Value = "Release Date"
    Cells(1, "I").Value = "Checkout Date"
    Cells(1, "J").Value = "EditSheet Available"
    Cells(1, "K").Value = "Created"
    Cells(1, "L").Value = "Created By"
    Cells(1, "M").Value = "Modified"
    Cells(1, "N").Value = "Modified By"
    Cells(1, "O").Value = "AOR"
    Cells(1, "P").Value = "Item Type"
    Cells(1, "Q").Value = "Path"
    
    Range("A1:Q1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 12611584
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    
    Columns("A:Q").Select
    Columns("A:Q").EntireColumn.AutoFit
    
    Cells(2 + i, "A").Value = RTU
    Cells(2 + i, "B").Value = devtype
    Cells(2 + i, "D").Value = system
    Cells(2 + i, "F").Value = "Not Started"
    Cells(2 + i, "J").Value = "TRUE"
    Cells(2 + i, "O").Value = AOR
    Cells(2 + i, "P").Value = "Item"

    ChDir "C:\Users\mcourte\Desktop\scaDAbuilder\To Do List"
    ActiveWorkbook.SaveAs Filename:= _
        "C:\Users\mcourte\Desktop\scaDAbuilder\To Do List\To Do List.xlsx", FileFormat _
        :=xlOpenXMLWorkbook, CreateBackup:=False
    ActiveWindow.Close
    
End If

Application.ScreenUpdating = True
Application.EnableEvents = True
Application.Calculation = xlCalculationAutomatic

End Sub

Sub alarmlocation(i As Long)

Application.ScreenUpdating = False
Application.EnableEvents = False
Application.Calculation = xlCalculationManual

Worksheets("Cover").Activate

Dim RTU As String: RTU = CStr(Cells(5, "L").Value)

If Dir("C:\Users\" + Environ("Username") + "\Desktop\scaDAbuilder\Project Files\T&D\AlarmLocation.xlsm") <> "" Then
    Workbooks.Open Filename:="C:\Users\" + Environ("Username") + "\Desktop\scaDAbuilder\Project Files\T&D\AlarmLocation.xlsm"
    Worksheets("AlarmLocation").Activate
    Numrows = Range("B1", Range("B1").End(xlDown)).Rows.Count + 2
    Cells(Numrows + i, "B").Value = RTU
    Cells(Numrows + i, "K").Value = "display /app=scada/viewport=alarm_oneline %LOCID%"
    ActiveWorkbook.Save
    ActiveWindow.Close
End If

If Dir("C:\Users\" + Environ("Username") + "\Desktop\scaDAbuilder\Project Files\DA\AlarmLocation.xlsm") <> "" Then
    Workbooks.Open Filename:="C:\Users\" + Environ("Username") + "\Desktop\scaDAbuilder\Project Files\DA\AlarmLocation.xlsm"
    Worksheets("AlarmLocation").Activate
    Numrows = Range("B1", Range("B1").End(xlDown)).Rows.Count + 2
    Cells(Numrows + i, "B").Value = RTU
    Cells(Numrows + i, "K").Value = "display /app=scada/viewport=alarm_oneline %LOCID%"
    ActiveWorkbook.Save
    ActiveWindow.Close
End If

Application.ScreenUpdating = True
Application.EnableEvents = True
Application.Calculation = xlCalculationAutomatic

End Sub

Sub FGDIS(j As Long)

Application.ScreenUpdating = False
Application.EnableEvents = False
Application.Calculation = xlCalculationManual

Worksheets("Cover").Activate

Dim RTU As String: RTU = CStr(Cells(5, "L").Value)

If Dir("C:\Users\" + Environ("Username") + "\Desktop\scaDAbuilder\Project Files\DA\FullGraphicsDisplayRecords.xlsm") <> "" Then
    Workbooks.Open Filename:="C:\Users\" + Environ("Username") + "\Desktop\scaDAbuilder\Project Files\DA\FullGraphicsDisplayRecords.xlsm"
    Worksheets("FullGraphicsDisplayRecords").Activate
    Numrows = Range("B1", Range("B1").End(xlDown)).Rows.Count + 2
    Cells(Numrows + j, "B").Value = RTU
    Cells(Numrows + j, "C").Value = RTU
    ActiveWorkbook.Save
    ActiveWindow.Close
End If

If Dir("C:\Users\" + Environ("Username") + "\Desktop\scaDAbuilder\Project Files\T&D\FullGraphicsDisplayRecords.xlsm") <> "" Then
    Workbooks.Open Filename:="C:\Users\" + Environ("Username") + "\Desktop\scaDAbuilder\Project Files\T&D\FullGraphicsDisplayRecords.xlsm"
    Worksheets("FullGraphicsDisplayRecords").Activate
    Numrows = Range("B1", Range("B1").End(xlDown)).Rows.Count + 2
    Cells(Numrows + j, "B").Value = RTU
    Cells(Numrows + j, "C").Value = RTU
    ActiveWorkbook.Save
    ActiveWindow.Close
End If

Application.ScreenUpdating = True
Application.EnableEvents = True
Application.Calculation = xlCalculationAutomatic

End Sub

Sub commstate(k As Long, FE As String, DNP As String)

Application.ScreenUpdating = False
Application.EnableEvents = False
Application.Calculation = xlCalculationManual

Dim x As Integer
Dim y As Integer

Worksheets("Cover").Activate
    
Dim RTU As String: RTU = CStr(Cells(5, "L").Value)
Dim AOR As String: AOR = CStr(Cells(10, "D").Value)

If Dir("C:\Users\" + Environ("Username") + "\Desktop\scaDAbuilder\Project Files\DA\SCADA SCDA COMMS - Substation Hierarchy.xlsm") <> "" Then

    Workbooks.Open Filename:="C:\Users\" + Environ("Username") + "\Desktop\scaDAbuilder\Project Files\DA\SCADA SCDA COMMS - Substation Hierarchy.xlsm"
    Worksheets("Command").Activate
    Numrows = Range("B1", Range("B1").End(xlDown)).Rows.Count
    x = Numrows - 1
    y = x + 3 + k
    Rows(x).Select
    Selection.Copy
    Cells(y, "A").Select
    Cells(y, "A").PasteSpecial
    Application.CutCopyMode = False
    
    Cells(Numrows + 2 + k, "A").Value = ""
    Cells(Numrows + 2 + k, "B").Value = "GenericEquipment COMMS RTU " + RTU + " STAT ENABLE"
    Cells(Numrows + 2 + k, "D").Value = ""
    Cells(Numrows + 2 + k, "H").Value = "GenericEquipment COMMS RTU " + RTU + " STAT"
    Cells(Numrows + 2 + k, "AH").Value = "LDAS" + FE
    Cells(Numrows + 2 + k, "AK").Value = "COMMS.RTU." + RTU + ".STAT.ENABLE"
    Cells(Numrows + 2 + k, "AL").Value = "JDAS" + FE

    Worksheets("Discrete").Activate
    Numrows = Range("B1", Range("B1").End(xlDown)).Rows.Count
    x = Numrows - 1
    y = x + 3 + k
    Rows(x).Select
    Selection.Copy
    Cells(y, "A").Select
    Cells(y, "A").PasteSpecial
    Application.CutCopyMode = False
    
    Cells(Numrows + 2 + k, "A").Value = ""
    Cells(Numrows + 2 + k, "B").Value = "GenericEquipment COMMS RTU " + RTU + " STAT"
    Cells(Numrows + 2 + k, "J").Value = AOR
    Cells(Numrows + 2 + k, "R").Value = ""
    Cells(Numrows + 2 + k, "DS").Value = "COMMS RTU " + RTU
    Cells(Numrows + 2 + k, "FJ").Value = "COMMS.RTU." + RTU + ".STAT"
    Cells(Numrows + 2 + k, "FL").Value = "JDAS" + FE
    Cells(Numrows + 2 + k, "FP").Value = "LDAS" + FE

    Worksheets("GenericEquipment").Activate
    Numrows = Range("B1", Range("B1").End(xlDown)).Rows.Count
    x = Numrows - 1
    y = x + 3 + k
    Rows(x).Select
    Selection.Copy
    Cells(y, "A").Select
    Cells(y, "A").PasteSpecial
    Application.CutCopyMode = False
    
    Cells(Numrows + 2 + k, "A").Value = ""
    Cells(Numrows + 2 + k, "B").Value = "COMMS RTU " + RTU
    Cells(Numrows + 2 + k, "T").Value = ""
    Cells(Numrows + 2 + k, "Z").Value = RTU
    Cells(Numrows + 2 + k, "BP").Value = AOR
    Cells(Numrows + 2 + k, "BU").Value = "COMMS.RTU." + RTU

End If

If Dir("C:\Users\" + Environ("Username") + "\Desktop\scaDAbuilder\Project Files\T&D\COMMS RTU_DA - EquipmentGroup Hierarchy.xlsm") <> "" Then

    Workbooks.Open Filename:="C:\Users\" + Environ("Username") + "\Desktop\scaDAbuilder\Project Files\T&D\COMMS RTU_DA - EquipmentGroup Hierarchy.xlsm"
    Worksheets("Command").Activate
    Numrows = Range("B1", Range("B1").End(xlDown)).Rows.Count
    x = Numrows - 1
    y = x + 3 + k
    Rows(x).Select
    Selection.Copy
    Cells(y, "A").Select
    Cells(y, "A").PasteSpecial
    Application.CutCopyMode = False

    
    Cells(Numrows + 2 + k, "A").Value = ""
    Cells(Numrows + 2 + k, "B").Value = "GenericEquipment COMMS RTU_DA " + RTU + " STAT ENABLE"
    Cells(Numrows + 2 + k, "D").Value = ""
    Cells(Numrows + 2 + k, "H").Value = "GenericEquipment COMMS RTU_DA " + RTU + " STAT"
    Cells(Numrows + 2 + k, "AK").Value = "COMMS.RTU_DA." + RTU + ".STAT.ENABLE"

    Worksheets("Discrete").Activate
    Numrows = Range("B1", Range("B1").End(xlDown)).Rows.Count
    x = Numrows - 1
    y = x + 3 + k
    Rows(x).Select
    Selection.Copy
    Cells(y, "A").Select
    Cells(y, "A").PasteSpecial
    Application.CutCopyMode = False

    
    Cells(Numrows + 2 + k, "A").Value = ""
    Cells(Numrows + 2 + k, "B").Value = "GenericEquipment COMMS RTU_DA " + RTU + " STAT"
    Cells(Numrows + 2 + k, "J").Value = AOR
    Cells(Numrows + 2 + k, "R").Value = ""
    Cells(Numrows + 2 + k, "DR").Value = "COMMS RTU_DA " + RTU
    Cells(Numrows + 2 + k, "FI").Value = "COMMS.RTU_DA." + RTU + ".STAT"

    Worksheets("GenericEquipment").Activate
    Numrows = Range("B1", Range("B1").End(xlDown)).Rows.Count
    x = Numrows - 1
    y = x + 3 + k
    Rows(x).Select
    Selection.Copy
    Cells(y, "A").Select
    Cells(y, "A").PasteSpecial
    Application.CutCopyMode = False

    
    Cells(Numrows + 2 + k, "A").Value = ""
    Cells(Numrows + 2 + k, "B").Value = "COMMS RTU_DA " + RTU
    Cells(Numrows + 2 + k, "T").Value = ""
    Cells(Numrows + 2 + k, "Z").Value = RTU
    Cells(Numrows + 2 + k, "BP").Value = AOR
    Cells(Numrows + 2 + k, "BU").Value = "COMMS.RTU_DA." + RTU

    Worksheets("InterSiteAliasName").Activate
    Numrows = Range("B1", Range("B1").End(xlDown)).Rows.Count
    x = Numrows - 1
    y = x + 3 + k
    Rows(x).Select
    Selection.Copy
    Cells(y, "A").Select
    Cells(y, "A").PasteSpecial
    Application.CutCopyMode = False

    
    Cells(Numrows + 2 + k, "A").Value = ""
    Cells(Numrows + 2 + k, "B").Value = "GenericEquipment COMMS RTU_DA " + RTU + " STAT STAT DASCADA"
    Cells(Numrows + 2 + k, "D").Value = ""
    Cells(Numrows + 2 + k, "H").Value = RTU
    Cells(Numrows + 2 + k, "M").Value = "GenericEquipment COMMS RTU_DA " + RTU + " STAT"
    Cells(Numrows + 2 + k, "R").Value = "(POINT) COMMS.RTU_DA." + RTU + ".STAT.STAT (DASCADA)"

End If

Application.ScreenUpdating = True
Application.EnableEvents = True
Application.Calculation = xlCalculationAutomatic

End Sub

'Michael Courteaux
'6/18/2020
'Edit Sheet Format Macro V3

Sub AlmCatCheck()

Dim i As Integer
Dim n As Integer
Dim j As Integer
Dim openPos1 As Integer
Dim openPos2 As Integer
Dim closePos1 As Integer
Dim closePos2 As Integer
Dim midBit1 As String
Dim midBit2 As String
Dim str As String

Application.ScreenUpdating = False
Application.EnableEvents = False
Application.Calculation = xlCalculationManual

Worksheets("Analog").Activate

n = 1

Numrows = Range("A10", Range("A10").End(xlDown)).Rows.Count

If Numrows < 4 Or Numrows > 100 Then
    Cells(10, "Z").Interior.ColorIndex = 6
    Cells(10, "AP").Value = "GenericEquipment " + Cells(3, "D").Value + " " + Cells(10, "F").Value + " " + Cells(10, "G").Value + " " + Cells(10, "O").Value
Else
    Do While n < 10
        For i = 10 To 9 + Numrows
            If Cells(i, "Z").Value = "Y" Then
                Cells(i, "Z").Interior.ColorIndex = 6
            End If
        Next
        For i = 10 To 9 + Numrows
            If Cells(i, "AO").Value = "Y" Then
                Cells(i, "AO").Interior.ColorIndex = 6
            End If
        Next

'************GENERATE GENERIC EQUIPMENT LINKAGES FOR ANALOG POINTS****************
        For i = 10 To 9 + Numrows
            Cells(i, "AP").Value = "GenericEquipment " + CStr(Cells(3, "D").Value) + " " + CStr(Cells(i, "F").Value) + " " + CStr(Cells(i, "G").Value) + " " + CStr(Cells(i, "O").Value)
        Next
        n = n + 1
    Loop
End If

Worksheets("Alarm").Activate

n = 1

Numrows = Range("B11", Range("B11").End(xlDown)).Rows.Count

Do While n < 2
    For i = 11 To 10 + Numrows
        If (Cells(i, "N").Value = "RCBL" Or (Cells(i, "N").Value = "RCLS")) And Cells(i, "M").Value <> "AUTO" Then
            Cells(i, "M").Value = "AUTO"
            Cells(i, "M").Interior.ColorIndex = 6
            Cells(i, "M").Font.ColorIndex = 3
       ElseIf (Cells(i, "N").Value = "STTS" Or Cells(i, "N").Value = "STTA" Or Cells(i, "N").Value = "STTB" Or Cells(i, "N").Value = "STTC") And Cells(i, "P").Value <> "FW" Then
            Cells(i, "P").Value = "FW"
            Cells(i, "P").Interior.ColorIndex = 6
            Cells(i, "P").Font.ColorIndex = 3
        ElseIf Cells(i, "K").Value = Cells(i, "M").Value And InStr(Cells(i, "P").Value, "C") > 0 And ((Cells(i, "N").Value <> "STTS" Or Cells(i, "N").Value <> "STTA" Or Cells(i, "N").Value <> "STTB" Or Cells(i, "N").Value <> "STTC") And Cells(i, "P").Value <> "FW") Then
            Cells(i, "P").Value = Replace(Cells(i, "P").Value, "C", "O")
            Cells(i, "P").Interior.ColorIndex = 6
            Cells(i, "P").Font.ColorIndex = 3
        ElseIf Cells(i, "L").Value = Cells(i, "M").Value And InStr(Cells(i, "P").Value, "O") > 0 And ((Cells(i, "N").Value <> "STTS" Or Cells(i, "N").Value <> "STTA" Or Cells(i, "N").Value <> "STTB" Or Cells(i, "N").Value <> "STTC") And Cells(i, "P").Value <> "FW") Then
            Cells(i, "P").Value = Replace(Cells(i, "P").Value, "O", "C")
            Cells(i, "P").Interior.ColorIndex = 6
            Cells(i, "P").Font.ColorIndex = 3
        End If
    Next
    For i = 11 To 10 + Numrows
        If Cells(i, "Q").Value = "Y" Then
            Cells(i, "Q").Interior.ColorIndex = 6
        End If
    Next
    For i = 11 To 10 + Numrows
        If Cells(i, "X").Value = "Y" Then
            Cells(i, "X").Interior.ColorIndex = 6
        End If
    Next
    
    '************GENERATE GENERIC EQUIPMENT LINKAGES FOR STATUS POINTS****************
    For i = 11 To 10 + Numrows
        Cells(i, "Y").Value = "GenericEquipment " + CStr(Cells(4, "D").Value) + " " + CStr(Cells(i, "H").Value) + " " + CStr(Cells(i, "I").Value) + " " + CStr(Cells(i, "N").Value)
    Next

    n = n + 1
Loop

Worksheets("Control").Activate

n = 1
j = 11
Numrows = Range("A11", Range("A11").End(xlDown)).Rows.Count

Do While n < 10
    For i = 10 To 10 + Numrows
        If (Cells(i, "G").Value = "RCBL" Or Cells(i, "G").Value = "RCLS") And Cells(i, "P").Value <> "AUTO" Then
            Cells(i, "P").Value = "AUTO"
            Cells(i, "P").Interior.ColorIndex = 6
            Cells(i, "P").Font.ColorIndex = 3
        ElseIf Cells(i, "Q").Value = "Y" Then
            Cells(i, "Q").Interior.ColorIndex = 6
      End If
    Next
    
'************GENERATE GENERIC EQUIPMENT LINKAGES FOR CONTROL POINTS****************
    For i = 11 To 10 + Numrows
    
        str = CStr(Cells(i, "M").Value)
        openPos1 = InStr(str, "(")
        closePos1 = InStr(str, "/")
        midBit1 = Mid(str, openPos1 + 1, closePos1 - openPos1 - 1)
        openPos2 = InStr(str, "/")
        closePos2 = InStr(str, ")")
        midBit2 = Mid(str, openPos2 + 1, closePos2 - openPos2 - 1)
        
        If midBit1 <> "SPARE" Then
            Cells(j, "R").Value = "GenericEquipment " + CStr(Cells(4, "E").Value) + " " + CStr(Cells(i, "K").Value) + " " + CStr(Cells(i, "L").Value) + " " + CStr(Cells(i, "G").Value) + " " + midBit1
        ElseIf midBit1 = CStr(LCase("spare")) Then
            Cells(j, "R").Value = Null
        End If
        
        If midBit2 <> "SPARE" Then
            Cells(j + 1, "R").Value = "GenericEquipment " + CStr(Cells(4, "E").Value) + " " + CStr(Cells(i, "K").Value) + " " + CStr(Cells(i, "L").Value) + " " + CStr(Cells(i, "G").Value) + " " + midBit2
        ElseIf midBit2 = CStr(LCase("spare")) Then
            Cells(j + 1, "R").Value = Null
        End If
        
        j = j + 2
        
    Next
    
    For i = 11 + (2 * Numrows) To 30 * Numrows
        Cells(i, "R") = Null
    Next
    
    n = n + 1
Loop

For i = 1 To Worksheets.Count
    If Worksheets(i).name = "Manual Points" Then
        exists = True
    End If
Next i
If exists = True Then
    Worksheets("Manual Points").Activate

n = 1

Numrows = Range("A11", Range("A11").End(xlDown)).Rows.Count

Do While n < 10
    For i = 10 To 10 + Numrows
        If Cells(i, "Q").Value = "Y" Then
            Cells(i, "Q").Interior.ColorIndex = 6
      End If
    Next
    For i = 10 To 10 + Numrows
        If Cells(i, "X").Value = "Y" Then
            Cells(i, "X").Interior.ColorIndex = 6
        End If
    Next
    For i = 10 To 10 + Numrows
        If Cells(i, "N").Value <> "STTS" Then
            Cells(i, "N").Value = "STTS"
        End If
    Next

    n = n + 1
Loop
End If

Worksheets("Cover").Activate

    Range("D4,D6,D8,D10,H6,L4,L5,L7,L8,L9,L10").Select
    Range("L9").Activate
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Selection.Font.Bold = True
    With Selection.Font
        .Color = -16776961
        .TintAndShade = 0
    End With

'***************THIS SECTION HIDES UNNEEDED COLUMNS***************
    Sheets("Analog").Select
    Columns("AA:AN").Select
    Selection.EntireColumn.Hidden = True
    Columns("AP:AP").EntireColumn.AutoFit
    Sheets("Control").Select
    Columns("R:R").EntireColumn.AutoFit
    Columns("S:S").EntireColumn.AutoFit
    Sheets("Alarm").Select
    Columns("R:W").Select
    Selection.EntireColumn.Hidden = True
    Columns("Y:Y").EntireColumn.AutoFit
    For i = 1 To Worksheets.Count
    If Worksheets(i).name = "Manual Points" Then
        exists = True
    End If
Next i

If exists = True Then
    Sheets("Manual Points").Select
    Columns("R:W").Select
    Selection.EntireColumn.Hidden = True
End If
    Sheets("Cover").Select
'***************SECTION END***************

Application.ScreenUpdating = True
Application.EnableEvents = True
Application.Calculation = xlCalculationAutomatic

End Sub


