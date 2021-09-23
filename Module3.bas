Attribute VB_Name = "Module3"
Sub Linker(title As String)

Application.ScreenUpdating = False
Application.EnableEvents = False
Application.Calculation = xlCalculationManual

Worksheets("Cover").Activate

Dim RTU As String: RTU = CStr(Cells(5, "L").Value)
Dim devtype As String: devtype = CStr(Cells(4, "L").Value)
Dim dispver As String
Dim scadardisp As String
Dim linkage As Range
Dim cl As Range
Dim wb1 As Workbook
Dim wb2 As Workbook

If devtype = "IntelliRupter" Then
    devtype = "IR"
End If

Worksheets("Display").Activate

If Cells(43, "A").Value = "" And Cells(46, "A").Value = "" And Cells(25, "A").Value = "" Then
    dispver = CStr(Cells(45, "A").Value)
ElseIf Cells(45, "A").Value = "" And Cells(46, "A").Value = "" And Cells(25, "A").Value = "" Then
    dispver = CStr(Cells(43, "A").Value)
ElseIf Cells(45, "A").Value = "" And Cells(43, "A").Value = "" And Cells(25, "A").Value = "" Then
    dispver = CStr(Cells(46, "A").Value)
ElseIf Cells(45, "A").Value = "" And Cells(43, "A").Value = "" And Cells(46, "A").Value = "" Then
    dispver = CStr(Cells(25, "A").Value)
End If

Select Case True
    Case (title Like "*_D1_*")
        scadardisp = "D1"
    Case (title Like "*_D2_*")
        scadardisp = "D2"
    Case (title Like "*_D3_*")
        scadardisp = "D3"
    Case (title Like "*_D4_*")
        scadardisp = "D4"
    Case (title Like "*_D5_*")
        scadardisp = "D5"
    Case (title Like "*_D6_*")
        scadardisp = "D6"
    Case (title Like "*_D7_*")
        scadardisp = "D7"
    Case (title Like "*_D8_*")
        scadardisp = "D8"
    Case (title Like "*_D9_*")
        scadardisp = "D9"
    Case (title Like "*_D10_*")
        scadardisp = "D10"
    Case (title Like "*_D11_*")
        scadardisp = "D11"
    Case (title Like "*_D12_*")
        scadardisp = "D12"
    Case (title Like "*_D13_*")
        scadardisp = "D13"
    Case (title Like "*_D14_*")
        scadardisp = "D14"
    Case (title Like "*_D15_*")
        scadardisp = "D15"
    Case (title Like "*_D16_*")
        scadardisp = "D16"
    Case (title Like "*_D17_*")
        scadardisp = "D17"
End Select

If dispver <> scadardisp Then
    dispver = scadardisp
End If

Dim scadar As String: scadar = devtype & dispver

If scadar = "351PD11" Then

    Workbooks.Open Filename:="C:\Users\" + Environ("Username") + "\Desktop\scaDAbuilder\Project Files\Templates\Link Template.xlsm"
    Set wb1 = ActiveWorkbook
    wb1.Worksheets("351PD11").Range("A1:F52").Copy 'Change worksheet
    
    'Open workbook 2
    Set wb2 = Workbooks.Open("C:\Users\" + Environ("Username") + "\Desktop\scaDAbuilder\Project Files\DA\DataItem.xlsm")

    With Worksheets("DataItem").Cells
        Set cl = .Find(RTU & " ANLG IED 0000", After:=.Range("B2"), LookIn:=xlValues)
        If Not cl Is Nothing Then
            cl.Select
        End If
    End With
    
    a = ActiveCell.Row
    
Application.DisplayAlerts = False
wb2.Sheets("DataItem").Cells(a, "AQ").PasteSpecial
Application.CutCopyMode = False
Range("A1").Select
Cells.Replace What:="XXXX", Replacement:=RTU, LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    
ElseIf scadar = "351RD11" Then

    Workbooks.Open Filename:="C:\Users\" + Environ("Username") + "\Desktop\scaDAbuilder\Project Files\Templates\Link Template.xlsm"
    Set wb1 = ActiveWorkbook
    wb1.Worksheets("351RD11").Range("A1:F53").Copy 'Change worksheet
    
    'Open workbook 2
    Set wb2 = Workbooks.Open("C:\Users\" + Environ("Username") + "\Desktop\scaDAbuilder\Project Files\DA\DataItem.xlsm")

    With Worksheets("DataItem").Cells
        Set cl = .Find(RTU & " ANLG IED 0000", After:=.Range("B2"), LookIn:=xlValues)
        If Not cl Is Nothing Then
            cl.Select
        End If
    End With
    
    a = ActiveCell.Row
    
Application.DisplayAlerts = False
wb2.Sheets("DataItem").Cells(a, "AQ").PasteSpecial
Application.CutCopyMode = False
Range("A1").Select
Cells.Replace What:="XXXX", Replacement:=RTU, LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2

ElseIf scadar = "351RD12" Then

    Workbooks.Open Filename:="C:\Users\" + Environ("Username") + "\Desktop\scaDAbuilder\Project Files\Templates\Link Template.xlsm"
    Set wb1 = ActiveWorkbook
    wb1.Worksheets("351RD12").Range("A1:F44").Copy 'Change worksheet
    
    'Open workbook 2
    Set wb2 = Workbooks.Open("C:\Users\" + Environ("Username") + "\Desktop\scaDAbuilder\Project Files\DA\DataItem.xlsm")

    With Worksheets("DataItem").Cells
        Set cl = .Find(RTU & " ANLG IED 0000", After:=.Range("B2"), LookIn:=xlValues)
        If Not cl Is Nothing Then
            cl.Select
        End If
    End With
    
    a = ActiveCell.Row
    
Application.DisplayAlerts = False
wb2.Sheets("DataItem").Cells(a, "AQ").PasteSpecial
Application.CutCopyMode = False
Range("A1").Select
Cells.Replace What:="XXXX", Replacement:=RTU, LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2

ElseIf scadar = "351RSD13" Then

    Workbooks.Open Filename:="C:\Users\" + Environ("Username") + "\Desktop\scaDAbuilder\Project Files\Templates\Link Template.xlsm"
    Set wb1 = ActiveWorkbook
    wb1.Worksheets("351RSD13").Range("A1:F31").Copy 'Change worksheet
    
    'Open workbook 2
    Set wb2 = Workbooks.Open("C:\Users\" + Environ("Username") + "\Desktop\scaDAbuilder\Project Files\DA\DataItem.xlsm")

    With Worksheets("DataItem").Cells
        Set cl = .Find(RTU & " ANLG IED 0000", After:=.Range("B2"), LookIn:=xlValues)
        If Not cl Is Nothing Then
            cl.Select
        End If
    End With
    
    a = ActiveCell.Row
    
Application.DisplayAlerts = False
wb2.Sheets("DataItem").Cells(a, "AQ").PasteSpecial
Application.CutCopyMode = False
Range("A1").Select
Cells.Replace What:="XXXX", Replacement:=RTU, LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    
ElseIf scadar = "651R2D1" Then

    Workbooks.Open Filename:="C:\Users\" + Environ("Username") + "\Desktop\scaDAbuilder\Project Files\Templates\Link Template.xlsm"
    Set wb1 = ActiveWorkbook
    wb1.Worksheets("651R2D1").Range("A1:F62").Copy 'Change worksheet
    
    'Open workbook 2
    Set wb2 = Workbooks.Open("C:\Users\" + Environ("Username") + "\Desktop\scaDAbuilder\Project Files\DA\DataItem.xlsm")

    With Worksheets("DataItem").Cells
        Set cl = .Find(RTU & " ANLG IED 0000", After:=.Range("B2"), LookIn:=xlValues)
        If Not cl Is Nothing Then
            cl.Select
        End If
    End With
    
    a = ActiveCell.Row
    
Application.DisplayAlerts = False
wb2.Sheets("DataItem").Cells(a, "AQ").PasteSpecial
Application.CutCopyMode = False
Range("A1").Select
Cells.Replace What:="XXXX", Replacement:=RTU, LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2

ElseIf scadar = "651RAD4" Then

    Workbooks.Open Filename:="C:\Users\" + Environ("Username") + "\Desktop\scaDAbuilder\Project Files\Templates\Link Template.xlsm"
    Set wb1 = ActiveWorkbook
    wb1.Worksheets("651RAD4").Range("A1:F43").Copy 'Change worksheet
    
    'Open workbook 2
    Set wb2 = Workbooks.Open("C:\Users\" + Environ("Username") + "\Desktop\scaDAbuilder\Project Files\DA\DataItem.xlsm")

    With Worksheets("DataItem").Cells
        Set cl = .Find(RTU & " ANLG IED 0000", After:=.Range("B2"), LookIn:=xlValues)
        If Not cl Is Nothing Then
            cl.Select
        End If
    End With
    
    a = ActiveCell.Row
    
Application.DisplayAlerts = False
wb2.Sheets("DataItem").Cells(a, "AQ").PasteSpecial
Application.CutCopyMode = False
Range("A1").Select
Cells.Replace What:="XXXX", Replacement:=RTU, LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2

ElseIf scadar = "DACD5" Then

    Workbooks.Open Filename:="C:\Users\" + Environ("Username") + "\Desktop\scaDAbuilder\Project Files\Templates\Link Template.xlsm"
    Set wb1 = ActiveWorkbook
    wb1.Worksheets("DACD5").Range("A1:F18").Copy 'Change worksheet
    
    'Open workbook 2
    Set wb2 = Workbooks.Open("C:\Users\" + Environ("Username") + "\Desktop\scaDAbuilder\Project Files\DA\DataItem.xlsm")

    With Worksheets("DataItem").Cells
        Set cl = .Find(RTU & " ANLG IED 0000", After:=.Range("B2"), LookIn:=xlValues)
        If Not cl Is Nothing Then
            cl.Select
        End If
    End With
    
    a = ActiveCell.Row
    
Application.DisplayAlerts = False
wb2.Sheets("DataItem").Cells(a, "AQ").PasteSpecial
Application.CutCopyMode = False
Range("A1").Select
Cells.Replace What:="XXXX", Replacement:=RTU, LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2

ElseIf scadar = "IRD2" Then

    Workbooks.Open Filename:="C:\Users\" + Environ("Username") + "\Desktop\scaDAbuilder\Project Files\Templates\Link Template.xlsm"
    Set wb1 = ActiveWorkbook
    wb1.Worksheets("IRD2").Range("A1:F63").Copy 'Change worksheet
    
    'Open workbook 2
    Set wb2 = Workbooks.Open("C:\Users\" + Environ("Username") + "\Desktop\scaDAbuilder\Project Files\DA\DataItem.xlsm")

    With Worksheets("DataItem").Cells
        Set cl = .Find(RTU & " ANLG IED 0000", After:=.Range("B2"), LookIn:=xlValues)
        If Not cl Is Nothing Then
            cl.Select
        End If
    End With
    
    a = ActiveCell.Row
    
Application.DisplayAlerts = False
wb2.Sheets("DataItem").Cells(a, "AQ").PasteSpecial
Application.CutCopyMode = False
Range("A1").Select
Cells.Replace What:="XXXX", Replacement:=RTU, LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2

ElseIf scadar = "IRD17" Then

    Workbooks.Open Filename:="C:\Users\" + Environ("Username") + "\Desktop\scaDAbuilder\Project Files\Templates\Link Template.xlsm"
    Set wb1 = ActiveWorkbook
    wb1.Worksheets("IRD17").Range("A1:F55").Copy 'Change worksheet
    
    'Open workbook 2
    Set wb2 = Workbooks.Open("C:\Users\" + Environ("Username") + "\Desktop\scaDAbuilder\Project Files\DA\DataItem.xlsm")

    With Worksheets("DataItem").Cells
        Set cl = .Find(RTU & " ANLG IED 0000", After:=.Range("B2"), LookIn:=xlValues)
        If Not cl Is Nothing Then
            cl.Select
        End If
    End With
    
    a = ActiveCell.Row
    
Application.DisplayAlerts = False
wb2.Sheets("DataItem").Cells(a, "AQ").PasteSpecial
Application.CutCopyMode = False
Range("A1").Select
Cells.Replace What:="XXXX", Replacement:=RTU, LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2


End If


Application.ScreenUpdating = True
Application.EnableEvents = True
Application.Calculation = xlCalculationAutomatic

End Sub


