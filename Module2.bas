Attribute VB_Name = "Module2"
Sub DisplayMaker(title As String)

Application.ScreenUpdating = False
Application.EnableEvents = False
Application.Calculation = xlCalculationManual

Worksheets("Cover").Activate

Dim scadar As String: scadar = CStr(Cells(4, "H").Value)
Dim AOR As String: AOR = CStr(Cells(10, "D").Value)
Dim jurisdiction As String: jurisdiction = CStr(Cells(4, "D").Value)
Dim RTU As String: RTU = CStr(Cells(5, "L").Value)
Dim devtype As String: devtype = CStr(Cells(4, "L").Value)
Dim dispver As String
Dim devkv As String
Dim menu As String
Dim scadardisp As String
Dim PathCrnt As String

If devtype = "IntelliRupter" Then
    devtype = "IR"
End If

If jurisdiction = "EAL" Then
    jurisdiction = "EAI"
    menu = "_AR"
ElseIf jurisdiction = "EML" Then
    jurisdiction = "EMI"
    menu = "_MPL"
ElseIf jurisdiction = "ETI" Then
    jurisdiction = "ETI"
    menu = "_ETI"
ElseIf AOR = "DOCNL" Then
    jurisdiction = "ELLN"
    menu = "_NLA"
ElseIf AOR = "DOCSL" Or AOR = "DOCSE" Then
    jurisdiction = "ELLS"
    menu = "_SLA"
ElseIf AOR = "DOCNO" Then
    jurisdiction = "ENOI"
    menu = "_SLA"
ElseIf AOR = "DOCWL" Or AOR = "DOCEL" Then
    jurisdiction = "EGSL"
    menu = "_EGSL"
End If

Worksheets("Analog").Activate

Dim linekv As Integer: linekv = Cells(10, "E").Value

If linekv < 5 Then
    devkv = "_4KV"
ElseIf linekv > 15 And linekv < 30 Then
    devkv = "_25KV"
ElseIf linekv > 30 Then
    devkv = "_34KV"
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

Worksheets("Cover").Activate

DisplayGen devkv, RTU, jurisdiction, devtype, dispver, menu

ChDir ("C:\Users\" & Environ("Username") & "\Desktop\scaDAbuilder\Displays\" & jurisdiction & "\")

Call Shell("C:\Users\" & Environ("Username") & "\Desktop\scaDAbuilder\Displays\" & jurisdiction & "\" & "Compile " & jurisdiction & " Displays.bat")

Application.ScreenUpdating = True
Application.EnableEvents = True
Application.Calculation = xlCalculationAutomatic

End Sub

Sub DisplayGen(devkv As String, RTU As String, jurisdiction As String, devtype As String, dispver As String, menu As String)

Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")

Dim Fileout As Object: Set Fileout = fso.CreateTextFile _
                        ("C:\Users\" + Environ("Username") + "\Desktop\scaDAbuilder\Displays\" + _
                        jurisdiction + "\" + jurisdiction + "_Display_" + RTU + ".txt", True, True)

Fileout.Writeline ""
Fileout.Writeline "     display " + """" + RTU + """"
Fileout.Writeline "     ("
Fileout.Writeline "         title(localize " + """" + "%DIS% [%DISAPP%][%DISFAM%][%HOST%]   (%VP%) %REF%" + """" + ")"
Fileout.Writeline ""
Fileout.Writeline "         application " + """" + "RECON" + """"
Fileout.Writeline "         ("
Fileout.Writeline "             color" + "(" + """" + "0,0,0" + """" + ")"
Fileout.Writeline "         )"
Fileout.Writeline "         application " + """" + "SCADA" + """"
Fileout.Writeline "         ("
Fileout.Writeline "             color" + "(" + """" + "0,0,0" + """" + ")"
Fileout.Writeline "         )"
Fileout.Writeline "         color" + "(" + """" + "0,0,0" + """" + ")"
Fileout.Writeline "         scale_to_fit_style(XY)"
Fileout.Writeline "         menu_bar_item " + """" + "SCADA_RELATED_DISPLAYS_MENU" + """" + "("
Fileout.Writeline "         label(localize " + """" + "Related Displays" + """" + ")"
Fileout.Writeline "         set(" + """" + "ONELINES" + """" + ") )"
Fileout.Writeline "         menu_bar_item " + """" + "ONELINES" + menu + """" + "("
Fileout.Writeline "         label(localize " + """" + "Onelines" + """" + ")"
Fileout.Writeline "         set(" + """" + "ONELINES_MENU" + """" + ") )"
Fileout.Writeline "         permitted_if"
Fileout.Writeline "         ("
Fileout.Writeline "             one_of("
Fileout.Writeline "             class("
Fileout.Writeline "             " + """" + "DSPTRWEA" + """" + ") )"
Fileout.Writeline "         )"
Fileout.Writeline "         horizontal_unit(10)"
Fileout.Writeline "         vertical_unit(10)"
Fileout.Writeline "         horizontal_page(50)"
Fileout.Writeline "         vertical_page(50)"
Fileout.Writeline "         refresh(4)"
Fileout.Writeline "         not locked_in_viewport"
Fileout.Writeline "         horizontal_scroll_bar"
Fileout.Writeline "         vertical_scroll_bar"
Fileout.Writeline "         std_menu_bar"
Fileout.Writeline "         not command_window"
Fileout.Writeline "         not on_top"
Fileout.Writeline "         not ret_last_tab_pnum"
Fileout.Writeline "         default_zoom(1.0000000)"
Fileout.Writeline "         simple_layer " + """" + "DEFAULT" + """" + ""
Fileout.Writeline "         ("
Fileout.Writeline "             not clip_to_regions"
Fileout.Writeline "             picture " + """" + "SCADA_BANNER_TO_TABULAR" + """" + ""
Fileout.Writeline "             ("
Fileout.Writeline "                 set(" + """" + "ONELINES" + """" + ")"
Fileout.Writeline "                 origin(0 0)"
Fileout.Writeline "                 xlocked"
Fileout.Writeline "                 ylocked"
Fileout.Writeline "             )"
Fileout.Writeline "             picture " + """" + "RTU_BANNER_RTUSTATE" + """" + ""
Fileout.Writeline "             ("
Fileout.Writeline "                 set(" + """" + "ONELINES" + """" + ")"
Fileout.Writeline "                 origin(978 2)"
Fileout.Writeline "                 xlocked"
Fileout.Writeline "                 ylocked"
Fileout.Writeline "             )"
Fileout.Writeline "             picture " + """" + "TO_RTU_8CHAR" + """" + ""
Fileout.Writeline "             ("
Fileout.Writeline "                 set(" + """" + "ONELINES" + """" + ")"
Fileout.Writeline "                 origin(994 8)"
Fileout.Writeline "                 xlocked"
Fileout.Writeline "                 ylocked"
Fileout.Writeline "                 composite_key"
Fileout.Writeline "                 ("
Fileout.Writeline "                     record(" + """" + "SUBSTN" + """" + ") record_key(" + """" + "COMMS" + """" + ")"
Fileout.Writeline "                     record(" + """" + "DEVTYP" + """" + ") record_key(" + """" + "RTU" + """" + ")"
Fileout.Writeline "                     record(" + """" + "DEVICE" + """" + ") record_key(" + """" + RTU + """" + ")"
Fileout.Writeline "                     record(" + """" + "POINT" + """" + ") record_key(" + """" + "STAT" + """" + ")"
Fileout.Writeline "                 )"
Fileout.Writeline "             )"
Fileout.Writeline "             picture " + """" + "DA_" + devtype + "_" + dispver + devkv + "" + """" + ""
Fileout.Writeline "             ("
Fileout.Writeline "                 set(" + """" + "ONELINES_DA" + """" + ")"
Fileout.Writeline "                 origin(306 62)"
Fileout.Writeline "                 composite_key"
Fileout.Writeline "                 ("
Fileout.Writeline "                     record(" + """" + "SUBSTN" + """" + ") record_key(" + """" + RTU + """" + ")"
Fileout.Writeline "                     record(" + """" + "DEVTYP" + """" + ") record_key(" + """" + "RECL" + """" + ")"
Fileout.Writeline "                     record(" + """" + "DEVICE" + """" + ") record_key(" + """" + RTU + """" + ")"
Fileout.Writeline "                     partial_key"
Fileout.Writeline "                 )"
Fileout.Writeline "             )"
Fileout.Writeline "             picture " + """" + "MAN_IN_STATION_DOC" + """" + ""
Fileout.Writeline "             ("
Fileout.Writeline "                 set(" + """" + "ONELINES" + """" + ")"
Fileout.Writeline "                 origin(340 0)"
Fileout.Writeline "                 xlocked"
Fileout.Writeline "                 ylocked"
Fileout.Writeline "                 composite_key"
Fileout.Writeline "                 ("
Fileout.Writeline "                     record(" + """" + "SUBSTN" + """" + ") record_key(" + """" + RTU + """" + ")"
Fileout.Writeline "                     record(" + """" + "DEVTYP" + """" + ") record_key(" + """" + "STN" + """" + ")"
Fileout.Writeline "                     record(" + """" + "DEVICE" + """" + ") record_key(" + """" + "DOC" + """" + ")"
Fileout.Writeline "                     record(" + """" + "POINT" + """" + ") record_key(" + """" + "MANS" + """" + ")"
Fileout.Writeline "                 )"
Fileout.Writeline "             )"
Fileout.Writeline "             text"
Fileout.Writeline "             ("
Fileout.Writeline "                 gab " + """" + "TEXT_TITLE" + """"
Fileout.Writeline "                 set(" + """" + "ONELINES" + """" + ")"
Fileout.Writeline "                 origin(524 5)"
Fileout.Writeline "                 xlocked"
Fileout.Writeline "                 ylocked"
Fileout.Writeline "                 localize " + """" + RTU + """" + ""
Fileout.Writeline "             )"
Fileout.Writeline "         )"
Fileout.Writeline "     );"
Fileout.Close

End Sub



