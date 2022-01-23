# VBA-RFQ_Follow_Up_Process
A VBA script that runs everyday in order to send out automatic follow-ups if the RfQs are missing quotations from suppliers. 



![image](https://user-images.githubusercontent.com/50633734/150686507-42dbd3c6-c1b9-4a73-9eb4-f0770a6df19e.png)




Sub Rfq_claims()
'
'
' Rfq_claims Macro
'

Application.DisplayAlerts = False

Set Connection = GetObject("SAPGUI").GetScriptingEngine.Children(0)
If Not IsObject(session) Then
   Set SAPsession = Connection.Children(0)
End If

Today = Format(Sheets("Variables").Range("a2"), "DD.MM.YYYY")

With SAPsession
        .findById("wnd[0]/tbar[0]/okcd").Text = "/n/BASF/Tbox_toolbox"
        .findById("wnd[0]").sendVKey 0
        

        .findById("wnd[0]/tbar[1]/btn[17]").press
        .findById("wnd[1]/usr/txtV-LOW").Text = "CLAIM RFQS"
        .findById("wnd[1]/usr/txtENAME-LOW").Text = ""
        .findById("wnd[1]/usr/txtV-LOW").caretPosition = 8
        .findById("wnd[1]").sendVKey 8
        .findById("wnd[0]/tbar[1]/btn[8]").press



'Jump to "In process" Tab & Copy the data to the clipboard
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC5").Select
        
        'Choose Layout

        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC5/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0505/cntlCUSTOM_505/shellcont/shell").pressToolbarContextButton "&MB_VARIANT"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC5/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0505/cntlCUSTOM_505/shellcont/shell").selectContextMenuItem "&LOAD"
        .findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").currentCellRow = 2
        .findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "2"
        .findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").contextMenu
        .findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectContextMenuItem "&FILTER"
        .findById("wnd[2]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").Text = "RFQCLAIM"
        .findById("wnd[2]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").caretPosition = 9
        .findById("wnd[2]").sendVKey 0
        .findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").currentCellColumn = "TEXT"
        .findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "0"
        .findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell
        
        
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC5/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0505/cntlCUSTOM_505/shellcont/shell").pressToolbarButton "EXPA"

        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC5/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0505/cntlCUSTOM_505/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC5/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0505/cntlCUSTOM_505/shellcont/shell").selectContextMenuItem "&PC"
        .findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").Select
        .findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").SetFocus
        .findById("wnd[1]/tbar[0]/btn[0]").press

'Paste information into Excel
            Sheets("Data").Activate
            Sheets("Data").Select
            Sheets("Data").Cells.ClearContents
            
            Sheets("Data").Select
            Sheets("Data").Range("A1").Select
            ActiveSheet.Paste
            
            FinalRow = Sheets("Data").Cells(Rows.Count, 1).End(xlUp).Row
        
            Sheets("Data").Range("A1:A" & FinalRow).Select
        
'Changes the format Text to Columns
            Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
                TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
                Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
                :="|", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, _
                1), Array(6, 1), Array(7, 1), Array(8, 1), Array(9, 1), Array(10, 1), Array(11, 1), Array(12 _
                , 1), Array(13, 1), Array(14, 1), Array(15, 1), Array(16, 1), Array(17, 1), Array(18, 1), _
                Array(19, 1), Array(20, 1), Array(21, 1), Array(22, 1), Array(23, 1)), _
                TrailingMinusNumbers:=True
            Application.DisplayAlerts = False
            Selection.Delete Shift:=xlToLeft
    
'Deletes uneccesary rows
            Sheets("Data").Range("1:3").Delete
            Rows(2).EntireRow.Delete
            
    
    
    
'Formating
            Range("O1").Select
            ActiveCell.FormulaR1C1 = "Date Formating"
            Range("P1").Select
            ActiveCell.FormulaR1C1 = "CopyPasta"
            Range("Q1").Select
            ActiveCell.FormulaR1C1 = "Status"
            Range("R1").Select
            ActiveCell.FormulaR1C1 = "Claim Date"
            Range("S1").Select
            ActiveCell.FormulaR1C1 = "Comment"
            Range("O2").Select
            Range("T1").Select
            ActiveCell.FormulaR1C1 = "TASK ID"
            Range("U1").Select
            ActiveCell.FormulaR1C1 = "Language RFQ"
            Range("V1").Select
            ActiveCell.FormulaR1C1 = "Date RFQ"
            Range("W1").Select
            ActiveCell.FormulaR1C1 = "Purch Group"
            Range("X1").Select
            ActiveCell.FormulaR1C1 = "Email"
            Range("O2").Select
            
            Columns("P:P").Select
            Selection.Delete Shift:=xlToLeft
            
            
    Columns("M:M").Select
    Selection.Replace What:="                    ", Replacement:="", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
            
            
'Date formatting column
                Range("O2").Select
                ActiveCell.FormulaR1C1 = "=IF(RC[-2]=0,(MID(RC[-4],(LEN(RC[-4]))-6,(LEN(RC[-4]))-7)&LEFT(RC[-4],(LEN(RC[-4]))-8)&RIGHT(RC[-4],(LEN(RC[-4]))-5)),(LEFT(RC[-2],5)&"".2021""))"
                Range("O3").Select
                
'Apply to all column



    FinalRow = Sheets("Data").Cells(Rows.Count, 1).End(xlUp).Row
    Range("O2").Select
    Selection.AutoFill Destination:=Range("O2:O" & FinalRow)



'Set date column as value
Dim rng As Range
For Each rng In Range("O:O")
    If rng.HasFormula Then
        rng.Formula = rng.Value
        
    End If
Next rng

'CopyPasta Column

'Range("O:O").Select
'                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False
 '               Application.CutCopyMode = False
'                Selection.TextToColumns Destination:=Range("P2"), DataType:=xlFixedWidth, _
                    OtherChar:="|", FieldInfo:=Array(0, 3), TrailingMinusNumbers:=True
                    
                    
'If condition to check if RFQ needs to be claimed
'Status Column
Range("P2").Select
            ActiveCell.FormulaR1C1 = _
                "=IF(DAYS(TODAY(),RC[-1])>4,""CLAIM"",""DON'T CLAIM"")"

            Selection.AutoFill Destination:=Range("P2:P" & FinalRow)
            
           ' Range("Q2:Q31").Select
            'Columns("Q:Q").Select
           ' Selection.NumberFormat = "General"
        
    Range("Q2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-1]=""CLAIM"",TODAY(),""NO CLAIM"")"
    Range("Q2").Select
    Selection.AutoFill Destination:=Range("Q2:Q" & FinalRow)
    
    
'Comment Column


    Range("R2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(MONTH(RC[-1])<10,""0""&MONTH(RC[-1]),MONTH(RC[-1]))&"".""&IF(DAY(RC[-1])<10,""0""&DAY(RC[-1]),DAY(RC[-1]))&"" ""&""RFQ CLAIMED BY BOT""&"" TASK ID ""&RC[1]"
    
        
'Apply to all cells
        
FinalRow = Sheets("Data").Cells(Rows.Count, 1).End(xlUp).Row
Range("R2").Select
Selection.AutoFill Destination:=Range("R2:R" & FinalRow)



'# TO ADD: Column O to date format#

    Columns("O:O").Select
    Selection.TextToColumns Destination:=Range("O1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="|", FieldInfo:=Array(1, 3), TrailingMinusNumbers:=True
    Columns("R:R").ColumnWidth = 42.57



'Copy-Paste Task ID Column
Range("S2").Select
    ActiveCell.FormulaR1C1 = "= ""Z"" &RANDBETWEEN(700,10000)"
    Range("s2").Select
    Selection.AutoFill Destination:=Range("S2:S" & FinalRow)
    Range("s:s").Select
    
    

For Each rng In Range("S:S")
    If rng.HasFormula Then
        rng.Formula = rng.Value
        
    End If
Next rng

For Each rng In Range("R:R")
    If rng.HasFormula Then
        rng.Formula = rng.Value
        
    End If
Next rng
End With
Call Get_data_RFQ
Call Send_emails



End Sub

Sub Get_data_RFQ()


'Loop To get RFQ's Language, Vendor's Email, purch group, rfq date


'Convert formulas to values

Dim rng As Range
For Each rng In Range("Q:Q")
    If rng.HasFormula Then
        rng.Formula = rng.Value
        
    End If
Next rng

Columns("G:G").Select
Selection.NumberFormat = "General"

lastitem = Sheets("Data").Cells(Rows.Count, 1).End(xlUp).Row


Set Connection = GetObject("SAPGUI").GetScriptingEngine.Children(0)
If Not IsObject(session) Then
   Set SAPsession = Connection.Children(0)
End If


With SAPsession
        .findById("wnd[0]/tbar[0]/okcd").Text = "/nme43"
        .findById("wnd[0]").sendVKey 0
End With
        

For e = 2 To lastitem
    RFQ = Sheets("Data").Range("G" & e).Value



'Parameter Check if Cell "STATUS" has "Claim"
    If Sheets("Data").Range("P" & e).Text = "CLAIM" Then
    
    'Loop
        With SAPsession
        
        
            .findById("wnd[0]/usr/ctxtRM06E-ANFNR").Text = RFQ
            
            .findById("wnd[0]/tbar[1]/btn[6]").press
            Sheets("Data").Range("U" & e).Value = SAPsession.findById("wnd[0]/usr/ctxtEKKO-SPRAS").Text
            .findById("wnd[0]/tbar[1]/btn[7]").press
        
            Sheets("Data").Range("V" & e).Value = SAPsession.findById("wnd[0]/usr/ctxtEKKO-BEDAT").Text
            Sheets("Data").Range("W" & e).Value = SAPsession.findById("wnd[0]/usr/ctxtEKKO-EKGRP").Text
            
            
            .findById("wnd[0]/usr/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtSZA1_D0100-SMTP_ADDR").SetFocus
            .findById("wnd[0]/usr/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtSZA1_D0100-SMTP_ADDR").caretPosition = 16
            
    'Email
            Sheets("Data").Range("X" & e).Value = SAPsession.findById("wnd[0]/usr/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtSZA1_D0100-SMTP_ADDR").Text
            .findById("wnd[0]/tbar[0]/btn[3]").press
            
            
        End With
    End If
    
Continue:
Next e
End Sub

Sub Getdata_RFQ()
'This Sub is not being used

'Loop To get RFQ's Language & Vendor's Email


'Convert formulas to values

Dim rng As Range
For Each rng In Range("Q:Q")
    If rng.HasFormula Then
        rng.Formula = rng.Value
        
    End If
Next rng

Columns("G:G").Select
Selection.NumberFormat = "General"
    
lastitem = Sheets("Data").Cells(Rows.Count, 1).End(xlUp).Row


Set Connection = GetObject("SAPGUI").GetScriptingEngine.Children(0)
If Not IsObject(session) Then
   Set SAPsession = Connection.Children(0)
End If


With SAPsession
        .findById("wnd[0]/tbar[0]/okcd").Text = "/nme43"
        .findById("wnd[0]").sendVKey 0
End With
        

For e = 2 To lastitem
    RFQ = Sheets("Data").Range("G" & e).Value



'Parameter Check if Cell "STATUS" has "Claim"
    If Sheets("Data").Range("P" & e).Text = "CLAIM" Then
    
    'Loop
        With SAPsession
        
        
            .findById("wnd[0]/usr/ctxtRM06E-ANFNR").Text = RFQ
            
            .findById("wnd[0]/tbar[1]/btn[6]").press
            Sheets("Data").Range("U" & e).Value = SAPsession.findById("wnd[0]/usr/ctxtEKKO-SPRAS").Text
            .findById("wnd[0]/tbar[1]/btn[7]").press
        
            Sheets("Data").Range("V" & e).Value = SAPsession.findById("wnd[0]/usr/ctxtEKKO-BEDAT").Text
            Sheets("Data").Range("W" & e).Value = SAPsession.findById("wnd[0]/usr/ctxtEKKO-EKGRP").Text
            
            
            .findById("wnd[0]/usr/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtSZA1_D0100-SMTP_ADDR").SetFocus
            .findById("wnd[0]/usr/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtSZA1_D0100-SMTP_ADDR").caretPosition = 16
            
    'Email
            Sheets("Data").Range("X" & e).Value = SAPsession.findById("wnd[0]/usr/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtSZA1_D0100-SMTP_ADDR").Text
            .findById("wnd[0]/tbar[0]/btn[3]").press
            
            
        End With
    End If
    
Continue:
Next e


End Sub

Sub Send_emails()
'Loop to send out email from Toolbox


'Variable last row
lastitem = Sheets("Data").Cells(Rows.Count, 1).End(xlUp).Row


'SAP Connection
Set Connection = GetObject("SAPGUI").GetScriptingEngine.Children(0)
If Not IsObject(session) Then
   Set SAPsession = Connection.Children(0)
End If

'Open Toolbox
With SAPsession
        .findById("wnd[0]/tbar[0]/okcd").Text = "/n/BASF/Tbox_toolbox"
        .findById("wnd[0]").sendVKey 0
End With
        

'For Loop Starts
For e = 2 To lastitem
'Variables
    RFQ = Sheets("Data").Range("G" & e).Value
    SolPe = Sheets("Data").Range("E" & e).Value
    PurchGrp = Sheets("Data").Range("W" & e).Value
    Email = Sheets("Data").Range("X" & e).Value
    Language = Sheets("Data").Range("U" & e).Value
    Comment = Sheets("Data").Range("R" & e).Value
    
'Parameter Check if Cell "STATUS" has "Claim"
    If Sheets("Data").Range("P" & e).Text = "CLAIM" Then
    
    'Loop
        With SAPsession
        
        
        .findById("wnd[0]/tbar[1]/btn[17]").press
        .findById("wnd[1]/usr/txtV-LOW").Text = "CLAIM RFQS"
        .findById("wnd[1]/usr/txtV-LOW").caretPosition = 10
        .findById("wnd[1]").sendVKey 0
        .findById("wnd[1]/usr/txtENAME-LOW").Text = ""
        .findById("wnd[1]/usr/txtENAME-LOW").SetFocus
        .findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 0
        .findById("wnd[1]").sendVKey 0
        .findById("wnd[1]/tbar[0]/btn[8]").press
        .findById("wnd[0]/usr/ctxtS_BANFN-LOW").Text = SolPe
        .findById("wnd[0]/tbar[1]/btn[8]").press
        
        'Choose other tab
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC5").Select
        

        
        
        'Open all RFQs
       .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC5/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0505/cntlCUSTOM_505/shellcont/shell").pressToolbarButton "EXPA"


        'Filter the RFQ

        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC5/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0505/cntlCUSTOM_505/shellcont/shell").pressToolbarContextButton "&MB_FILTER"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC5/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0505/cntlCUSTOM_505/shellcont/shell").selectContextMenuItem "&DELETE_FILTER"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC5/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0505/cntlCUSTOM_505/shellcont/shell").setCurrentCell -1, "ANFNR"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC5/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0505/cntlCUSTOM_505/shellcont/shell").selectColumn "ANFNR"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC5/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0505/cntlCUSTOM_505/shellcont/shell").pressToolbarButton "&MB_FILTER"
        .findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").Text = RFQ
        .findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").caretPosition = 10
        .findById("wnd[1]/tbar[0]/btn[0]").press



        
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC5").Select
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC5/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0505/cntlCUSTOM_505/shellcont/shell").modifyCell 0, "ZACTION", "Mail a proveedor"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC5/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0505/cntlCUSTOM_505/shellcont/shell").currentCellColumn = "ZACTION"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC5/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0505/cntlCUSTOM_505/shellcont/shell").firstVisibleColumn = "EKGRP"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC5/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0505/cntlCUSTOM_505/shellcont/shell").pressEnter
        
        'Nested IF Statemnt depending on the Language
        
        
        
'This is hard codded depening on the language :(
        'RFQ sent in Spanish
        If Language = "ES" Then
            
            .findById("wnd[1]/usr/subSENDSCREEN:/BASF/SAPLTBOX_MAIL:1020/subOBJECT:/BASF/SAPLTBOX_MAIL:2300/txtSOS33-OBJDES").Text = "Estado Peticion de Oferta " & RFQ & " " & "[" & PurchGrp & "]"
            .findById("wnd[1]/usr/subSENDSCREEN:/BASF/SAPLTBOX_MAIL:1020/subOBJECT:/BASF/SAPLTBOX_MAIL:2300/txtSOS33-OBJDES").SetFocus
            .findById("wnd[1]/usr/subSENDSCREEN:/BASF/SAPLTBOX_MAIL:1020/subOBJECT:/BASF/SAPLTBOX_MAIL:2300/txtSOS33-OBJDES").caretPosition = 26
            .findById("wnd[1]").sendVKey 0


        'RFQ sent in English
        ElseIf Language = "EN" Then
        
            'Choose English Template
            .findById("wnd[1]/usr/subSENDSCREEN:/BASF/SAPLTBOX_MAIL:1020/subOBJECT:/BASF/SAPLTBOX_MAIL:2300/cmbGV_SPRAS").SetFocus
            .findById("wnd[1]/usr/subSENDSCREEN:/BASF/SAPLTBOX_MAIL:1020/subOBJECT:/BASF/SAPLTBOX_MAIL:2300/cmbGV_SPRAS").Key = "InglÃ©s"
            
            'Subject
             .findById("wnd[1]/usr/subSENDSCREEN:/BASF/SAPLTBOX_MAIL:1020/subOBJECT:/BASF/SAPLTBOX_MAIL:2300/txtSOS33-OBJDES").Text = "Status of RFQ " & RFQ & " " & "[" & PurchGrp & "]"
        
    
            'Remove the CC
            .findById("wnd[1]/usr/subSENDSCREEN:/BASF/SAPLTBOX_MAIL:1020/subOBJECT:/BASF/SAPLTBOX_MAIL:2300/radGV_CC_NO").SetFocus
            .findById("wnd[1]/usr/subSENDSCREEN:/BASF/SAPLTBOX_MAIL:1020/subOBJECT:/BASF/SAPLTBOX_MAIL:2300/radGV_CC_NO").Select

        
        
        'RFQ sent in Italian
        ElseIf Language = "IT" Then
        
            'Italian Template
            .findById("wnd[1]/usr/subSENDSCREEN:/BASF/SAPLTBOX_MAIL:1020/subOBJECT:/BASF/SAPLTBOX_MAIL:2300/cmbGV_SPRAS").Key = "Italiano"
            
            
            'Subject
            .findById("wnd[1]/usr/subSENDSCREEN:/BASF/SAPLTBOX_MAIL:1020/subOBJECT:/BASF/SAPLTBOX_MAIL:2300/txtSOS33-OBJDES").Text = "Stato della Richiesta di Offerta " & RFQ & " [" & PurchGrp & "]"
            
            'Body
            .findById("wnd[1]/usr/subSENDSCREEN:/BASF/SAPLTBOX_MAIL:1020/subOBJECT:/BASF/SAPLTBOX_MAIL:2300/tabsSO33_TAB1/tabpTAB1/ssubSUB1:/BASF/SAPLTBOX_MAIL:2100/cntlEDITOR/shellcont/shell").Text = "Gentile fornitore," & vbCr & "" & vbCr & "Potrebbe cortesemente aggiornarci sullo stato della richiesta di offerta " & RFQ & "." & vbCr & "" & vbCr & "Grazie e Cordiali saluti," & vbCr & "Support Team" & vbCr & "" & vbCr & "" & vbCr & "" & vbCr & "" & vbCr & ""
            
            'Remove the CC
            .findById("wnd[1]/usr/subSENDSCREEN:/BASF/SAPLTBOX_MAIL:1020/subOBJECT:/BASF/SAPLTBOX_MAIL:2300/radGV_CC_NO").SetFocus
            .findById("wnd[1]/usr/subSENDSCREEN:/BASF/SAPLTBOX_MAIL:1020/subOBJECT:/BASF/SAPLTBOX_MAIL:2300/radGV_CC_NO").Select
            
        
        
        'RFQ sent in other language
        ElseIf Language = "DE" Or Languge = "FR" Then
        
            'Choose English Template
            .findById("wnd[1]/usr/subSENDSCREEN:/BASF/SAPLTBOX_MAIL:1020/subOBJECT:/BASF/SAPLTBOX_MAIL:2300/cmbGV_SPRAS").SetFocus
            On Error Resume Next
            .findById("wnd[1]/usr/subSENDSCREEN:/BASF/SAPLTBOX_MAIL:1020/subOBJECT:/BASF/SAPLTBOX_MAIL:2300/cmbGV_SPRAS").Key = "InglÃ©s"
    
            'Subject
             .findById("wnd[1]/usr/subSENDSCREEN:/BASF/SAPLTBOX_MAIL:1020/subOBJECT:/BASF/SAPLTBOX_MAIL:2300/txtSOS33-OBJDES").Text = "Status of RFQ " & RFQ & " " & "[" & PurchGrp & "]"
        
            'Remove the CC
            .findById("wnd[1]/usr/subSENDSCREEN:/BASF/SAPLTBOX_MAIL:1020/subOBJECT:/BASF/SAPLTBOX_MAIL:2300/radGV_CC_NO").SetFocus
            .findById("wnd[1]/usr/subSENDSCREEN:/BASF/SAPLTBOX_MAIL:1020/subOBJECT:/BASF/SAPLTBOX_MAIL:2300/radGV_CC_NO").Select

        
        End If


'TO ADD: DUTCH & GERMAN FRENCH-BELGIUM LANGUAGE


'SEND EMAIL

        'Add email
        .findById("wnd[1]/usr/subSENDSCREEN:/BASF/SAPLTBOX_MAIL:1020/subRECLIST:/BASF/SAPLTBOX_MAIL:0103/tabsTAB_CONTROL/tabpREC/ssubSUB1:/BASF/SAPLTBOX_MAIL:0150/tbl/BASF/SAPLTBOX_MAILREC_CONTROL/txtSOS04-L_ADR_NAME[0,0]").Text = Email
        .findById("wnd[1]/usr/subSENDSCREEN:/BASF/SAPLTBOX_MAIL:1020/subRECLIST:/BASF/SAPLTBOX_MAIL:0103/tabsTAB_CONTROL/tabpREC/ssubSUB1:/BASF/SAPLTBOX_MAIL:0150/tbl/BASF/SAPLTBOX_MAILREC_CONTROL/txtSOS04-L_ADR_NAME[0,0]").caretPosition = 29
        
        .findById("wnd[1]").sendVKey 0
        
        'Remove CC
        .findById("wnd[1]/usr/subSENDSCREEN:/BASF/SAPLTBOX_MAIL:1020/subOBJECT:/BASF/SAPLTBOX_MAIL:2300/radGV_CC_NO").SetFocus
        .findById("wnd[1]/usr/subSENDSCREEN:/BASF/SAPLTBOX_MAIL:1020/subOBJECT:/BASF/SAPLTBOX_MAIL:2300/radGV_CC_NO").Select
        
        .findById("wnd[1]/tbar[0]/btn[20]").press




'Change Status to "REMINDER SENT"

        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC5/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0505/cntlCUSTOM_505/shellcont/shell").modifyCell 0, "RFQ_AN_STATUS", "Reminder sent"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC5/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0505/cntlCUSTOM_505/shellcont/shell").currentCellColumn = "RFQ_AN_STATUS"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC5/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0505/cntlCUSTOM_505/shellcont/shell").pressEnter
           
            
'Add Commment

        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC5/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0505/cntlCUSTOM_505/shellcont/shell").currentCellColumn = "RTEXT"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC5/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0505/cntlCUSTOM_505/shellcont/shell").doubleClickCurrentCell
        .findById("wnd[1]/usr/subSUB1:SAPLYMM1_UI_COMPONENTS:0100/cntlTEXTEDIT/shellcont/shell").Text = Comment
        .findById("wnd[1]/usr/subSUB1:SAPLYMM1_UI_COMPONENTS:0100/cntlTEXTEDIT/shellcont/shell").setSelectionIndexes 4, 4
        
        .findById("wnd[1]/tbar[0]/btn[0]").press
        
        .findById("wnd[0]/tbar[0]/btn[3]").press

            
        End With
    End If
    
Continue:
Next e


End Sub


