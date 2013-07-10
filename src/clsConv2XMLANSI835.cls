VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConv2XML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

Option Explicit
 Public num As Long
 
 Event writeXML(strXMLPath As String, strVar As String)
 
 Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
 Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long

 Private m_ShowProgress As Boolean
 Private m_TmrID As Long

 Private Const tmrDelay = 50

Public Event Timer(ByVal ShowProgress As Boolean)
Private Sub Class_Initialize()
'
'
' m_ShowProgress = True
'   '
'   ' Start timer
'   '
'   m_TmrID = SetTimer(0&, 0&, tmrDelay, AddressOf TimerProc)
End Sub

Private Sub Class_Terminate()

'
'
'Call KillTimer(0&, m_TmrID) 'Kill timer

End Sub
Public Sub RaiseEvents()
   m_ShowProgress = Not m_ShowProgress
   RaiseEvent Timer(m_ShowProgress)
End Sub

  Public Function diceXML(mstrparse As Variant) As String
    Dim str2XML As Variant
    Dim strReturnXML As String
    Dim i As Long
    Dim nl As String
    strReturnXML = "<ROOT_835>"
   If Right$(mstrparse(0), 2) = vbCrLf Then nl = vbCrLf Else nl = vbLf
  'Debug.Print UBound(str2XML)
For i = 0 To UBound(mstrparse) - 1
  'm_booleanEvent = False 'this is used to determine Form3.progressbar1.value
 
  
  str2XML = Split(Replace(CStr(mstrparse(i)), nl, ""), "*")
 
 Select Case str2XML(0)
 Case "GS"
    GStrans = str2XML(2) & "|" & str2XML(3) & "|" & str2XML(4) & "|" & str2XML(5) & "|" & str2XML(6) & "|"
    '********Field 8 of the GS segment determines the version of the file*******
    If InStr(1, str2XML(8), "004010X", vbTextCompare) Then boolv5010 = False Else boolv5010 = True
 Case "ST"
    STtrans = GStrans & str2XML(2)
   diceXML = "<ST>" & STconv(str2XML) & "</ST>"
 Case "BPR"
   BPRstr = str2XML(2)
   diceXML = "<BPR><id>" & CStr(STtrans) & "</id>" & BPRconv(str2XML) & "</BPR>"
 Case "TRN"
   TRNtracer = CStr(str2XML(2))
   diceXML = "<TRN><id>" & CStr(STtrans) & "</id><TRN_amount>" & BPRstr & "</TRN_amount>" & TRNconv(str2XML) & "</TRN>"
 Case "CUR"
   diceXML = "<CUR>" & CURconv(str2XML) & "</CUR>"
 Case "REF"
   diceXML = "<REF><id>" & CLPstr & "|" & CLPUid & "</id>" & REFconv(str2XML) & "</REF>"
 Case "DTM"
   diceXML = "<DTM><id>" & CLPstr & "|" & CLPUid & "</id><line_item>" & lnItem & "</line_item>" & DTMconv(str2XML) & "</DTM>"
 Case "N1"
   'diceXML = "<N1><id>" & CLPstr & " " & CLPUid & "</id>" & N1conv(str2XML) & "</N1>"
   diceXML = "<N1><id>" & CStr(STtrans) & "</id>" & N1conv(str2XML) & "</N1>"
 Case "N3"
 diceXML = "<N3><id>" & CStr(STtrans) & "</id>" & N3conv(str2XML) & "</N3>"
 Case "N4"
   diceXML = "<N4><id>" & CStr(STtrans) & "</id>" & N4conv(str2XML) & "</N4>"
 Case "PER"
   diceXML = "<PER>" & PERconv(str2XML) & "</PER>"
 Case "RDM"
 '*************************** New Segment for 5010 *****************************
   diceXML = "<RDM><id>" & CStr(STtrans) & "</id>" & RDMconv(str2XML) & "</RDM>"
 Case "LX"
   diceXML = "<LX>" & LXconv(str2XML) & "</LX>"
 Case "TS3"
   diceXML = "<TS3><id>" & STtrans & "</id>" & TS3conv(str2XML) & "</TS3>"
 Case "TS2"
   diceXML = "<TS2><id>" & STtrans & "</id>" & TS2conv(str2XML) & "</TS2>"
 Case "CLP"
   If UBound(str2XML) < 13 Then ReDim Preserve str2XML(13)
   lnItem = 0
   CLPUid = CLPUid + 1
   CLPstr = CStr(str2XML(7))
   diceXML = "<CLP><id>" & TRNtracer & "</id>" & CLPconv(str2XML) & "</CLP>"
 Case "CAS"
  diceXML = "<CAS><id>" & CLPstr & "|" & CLPUid & "</id><line_item>" & lnItem & "</line_item>" & CASconv(str2XML) & "</CAS>"
 Case "NM1"
  diceXML = "<NM1><id>" & CLPstr & "|" & CLPUid & "</id>" & NM1conv(str2XML) & "</NM1>"
 Case "MIA"
   diceXML = "<MIA><id>" & CLPstr & "|" & CLPUid & "</id>" & MIAconv(str2XML) & "</MIA>"
 Case "MOA"
   diceXML = "<MOA><id>" & CLPstr & "|" & CLPUid & "</id><line_item>" & lnItem & "</line_item>" & MOAconv(str2XML) & "</MOA>"
 Case "AMT"
   diceXML = "<AMT><id>" & CLPstr & "|" & CLPUid & "</id><line_item>" & lnItem & "</line_item>" & AMTconv(str2XML) & "</AMT>"
 Case "QTY"
   diceXML = "<QTY><id>" & CLPstr & "|" & CLPUid & "</id><line_item>" & lnItem & "</line_item>" & QTYconv(str2XML) & "</QTY>"
 Case "SVC"
  lnItem = lnItem + 1
   SVCstr = Mid(CStr(str2XML(1)), 4)
   diceXML = "<SVC><id>" & CLPstr & "|" & CLPUid & "</id><SVC_line_item>" & lnItem & "</SVC_line_item>" & SVCconv(str2XML) & "</SVC>"
 Case "LQ"
  diceXML = "<LQ><id>" & CLPstr & "|" & CLPUid & "</id><line_item>" & lnItem & "</line_item>" & LQconv(str2XML) & "</LQ>"
 Case "PLB"
  diceXML = "<PLB><id>" & STtrans & "</id>" & PLBconv(str2XML) & "</PLB>"
 Case "SE"
 ' IF char(s) are in form2.text2 then needs to be appended to the file name
 'Mid$(strcomplete, InStrRev(strcomplete, "\") + 1)
   If strXMLAddin <> "" Then
  diceXML = "<SE>" & "<SE_file>" & strXMLAddin & strcomplete & "</SE_file>" & "<import_date>" & Format$(Now, "mmddyy h:m AM/PM") & "</import_date>" & SEconv(str2XML) & "</SE>"
   Else
   diceXML = "<SE>" & "<SE_file>" & strcomplete & "</SE_file>" & "<import_date>" & Format$(Now, "mmddyy h:m AM/PM") & "</import_date>" & SEconv(str2XML) & "</SE>"
   End If
 Case Else
  diceXML = ""
 End Select
'        mtimer.Interval = 50
'        mtimer.Enabled = True
       
       strReturnXML = strReturnXML & diceXML
       DoEvents
       'Want to turn off progress bar if using command line
       If boolIDEMode Then
        RaiseEvents
       End If
Next
       strReturnXML = strReturnXML & "</ROOT_835>"
       DoEvents
       RaiseEvent writeXML(strXMLPath, strReturnXML)
 
             
End Function


Public Function STconv(varXML As Variant) As String
  STconv = "<ST_trans_code>" & varXML(1) & "</ST_trans_code>" & "<ST_trans_nm>" & CStr(GStrans) & varXML(2) & "</ST_trans_nm>"
End Function

Public Function BPRconv(varXML As Variant) As String
 If UBound(varXML) < 16 Then ReDim Preserve varXML(16)
 BPRconv = "<BPR_trans_code>" & varXML(1) & "</BPR_trans_code>" _
  & "<BPR_amount>" & varXML(2) & "</BPR_amount>" & "<BPR_credit_debit_flag>" _
  & varXML(3) & "</BPR_credit_debit_flag>" & "<BPR_payment_method>" & varXML(4) & "</BPR_payment_method>" _
  & "<BPR_payment_format>" & varXML(5) & "</BPR_payment_format><BPR_DFI_qualifier>" & varXML(6) & "</BPR_DFI_qualifier><BPR_DFI_num>" & varXML(7) & "</BPR_DFI_num><BPR_acct_num_qualifier>" & varXML(8) & "</BPR_acct_num_qualifier><BPR_acct_num>" & varXML(9) & "</BPR_acct_num><BPR_orig_company_id>" & varXML(10) & "</BPR_orig_company_id><BPR_orig_company_code>" & varXML(11) & "</BPR_orig_company_code><BPR_DFI_qualifier2>" & varXML(12) & "</BPR_DFI_qualifier2><BPR_DFI_num2>" & varXML(13) & "</BPR_DFI_num2><BPR_acct_num_qualifier2>" & varXML(14) & "</BPR_acct_num_qualifier2><BPR_acct_num2>" & varXML(15) & "</BPR_acct_num2><BPR_date>" & varXML(16) & "</BPR_date>"
  
End Function

Public Function TRNconv(varXML As Variant) As String
 If UBound(varXML) < 4 Then ReDim Preserve varXML(4)
 TRNconv = "<TRN_trans_code>" & varXML(1) & "</TRN_trans_code>" & "<TRN_ref_id>" & varXML(2) & "</TRN_ref_id>" _
 & "<TRN_payer_id>" & varXML(3) & "</TRN_payer_id><TRN_original_id>" & varXML(4) & "</TRN_original_id>"
End Function
Public Function CURconv(varXML As Variant) As String
CURconv = "<CUR_id>" & varXML(1) & "</CUR_id><CUR_code>" & varXML(2) & "</CUR_code>" '<CUR_rate>" & varXML(3) & "</CUR_rate>"
End Function

Public Function REFconv(varXML As Variant) As String
REFconv = "<REF_id>" & varXML(1) & "</REF_id>" & "<REF_num>" & varXML(2) & "</REF_num>"
End Function

Public Function DTMconv(varXML As Variant) As String
DTMconv = "<DTM_qualifier>" & varXML(1) & "</DTM_qualifier>" & "<DTM_date>" & varXML(2) & "</DTM_date>"
End Function

Public Function N1conv(varXML As Variant) As String
If UBound(varXML) < 4 Then ReDim Preserve varXML(4)
N1conv = "<N1_id>" & varXML(1) & "</N1_id><N1_name>" & Replace09(CStr(varXML(2)), "&", "and", , , vbTextCompare) & "</N1_name><N1_code_qualifier>" _
& varXML(3) & "</N1_code_qualifier><N1_code>" & varXML(4) & "</N1_code>"
End Function

Public Function N3conv(varXML As Variant) As String
If UBound(varXML) < 2 Then ReDim Preserve varXML(2)
N3conv = "<N3_address_1>" & varXML(1) & "</N3_address_1><N3_address_2>" & varXML(2) & "</N3_address_2>"
End Function

Public Function N4conv(varXML As Variant) As String
'If UBound(varXML) < 3 Then ReDim Preserve varXML(3)
Dim StringTmp As String
ReDim Preserve varXML(8)
StringTmp = "<N4_city>" & varXML(1) & "</N4_city><N4_state_prov>" & varXML(2) & "</N4_state_prov><N4_postal_code>" _
& varXML(3) & "</N4_postal_code>"
If boolv5010 Then
    StringTmp = StringTmp & "<N4_Country_code>" & varXML(4) & "</N4_Country_code><N4_Loc_Qualifier>" _
    & varXML(5) & "</N4_Loc_Qualifier><N4_Loc_Id>" & varXML(6) & "</N4_Loc_Id><N4_Country_Subcode>" _
    & varXML(7) & "</N4_Country_Subcode>"
End If
N4conv = StringTmp
End Function

Public Function PERconv(varXML As Variant) As String
ReDim Preserve varXML(8)
PERconv = "<PER_contact_code>" & varXML(1) & "</PER_contact_code><PER_name>" & varXML(2) & "</PER_name>" _
& "<PER_comm_qualifier>" & varXML(3) & "</PER_comm_qualifier><PER_comm_number>" & varXML(4) & "</PER_comm_number>" _
& "<PER_comm_qualifier2>" & varXML(5) & "</PER_comm_qualifier2><PER_comm_number2>" & varXML(6) & "</PER_comm_number2>" _
& "<PER_comm_qualifier3>" & varXML(7) & "</PER_comm_qualifier3><PER_comm_number3>" & varXML(8) & "</PER_comm_number3>"



End Function
Public Function RDMconv(varXML As Variant) As String
ReDim Preserve varXML(6)
RDMconv = "<RDM_Report_Transcode>" & varXML(1) & "</RDM_Report_Transcode><RDM_name>" & varXML(2) _
            & "</RDM_name><RDM_CommNum>" & varXML(3) & "</RDM_CommNum><RDM_RefID>" & varXML(4) _
            & "</RDM_RefID><RDM_RefID2>" & varXML(5) & "</RDM_RefID2>"
End Function
Public Function LXconv(varXML As Variant) As String
If UBound(varXML) < 1 Then ReDim Preserve varXML(1)
LXconv = "<LX_number>" & varXML(1) & "</LX_number>"

End Function

Public Function TS3conv(varXML As Variant) As String
Dim i As Variant
Dim tempstr As String
  
 For i = 1 To UBound(varXML)
  tempstr = tempstr & "<TS3_" & i & ">" & varXML(i) & "</TS3_" & i & ">"
  Next
TS3conv = tempstr

End Function

Public Function TS2conv(varXML As Variant) As String
Dim i As Variant
Dim tempstr As String
  
 For i = 1 To UBound(varXML)
  tempstr = tempstr & "<TS2_" & i & ">" & varXML(i) & "</TS2_" & i & ">"
  Next
TS2conv = tempstr
End Function

Public Function CLPconv(varXML As Variant) As String
Dim StringTemp As String
StringTemp = "<CLP_claim_id>" & varXML(1) & "</CLP_claim_id>" & "<CLP_status_code>" & varXML(2) & "</CLP_status_code>" _
& "<CLP_amount>" & varXML(3) & "</CLP_amount><CLP_actual_amount>" & varXML(4) & "</CLP_actual_amount><CLP_amount2>" & varXML(5) & "</CLP_amount2><CLP_claim_indicator>" _
 & varXML(6) & "</CLP_claim_indicator><CLP_payor_control_number>" & varXML(7) & "|" & CLPUid & "</CLP_payor_control_number><CLP_bill_type>" & varXML(8) & varXML(9) & "</CLP_bill_type>" _
 & "<CLP_patient_status>" & varXML(10) & "</CLP_patient_status><CLP_drg>" & varXML(11) & "</CLP_drg><CLP_drg_weight>" _
 & varXML(12) & "</CLP_drg_weight><CLP_discharge_fraction>" & varXML(13) & "</CLP_discharge_fraction>"
If boolv5010 Then
 If UBound(varXML) < 15 Then ReDim Preserve varXML(15)
 StringTemp = StringTemp & "<CLP_Response_code>" & varXML(14) & "</CLP_Response_code>"
End If

CLPconv = StringTemp
End Function

Public Function CASconv(varXML As Variant) As String
Dim StringTemp As String
If UBound(varXML) < 8 Then ReDim Preserve varXML(8)
StringTemp = "<CAS_claim_adj_group1>" & varXML(1) & "</CAS_claim_adj_group1>" & "<CAS_claim_reason1>" & varXML(2) & "</CAS_claim_reason1>" _
& "<CAS_amount1>" & varXML(3) & "</CAS_amount1>" & "<CAS_quantity1>" & varXML(4) & "</CAS_quantity1>" _
& "<CAS_claim_reason2>" & varXML(5) & "</CAS_claim_reason2>" & "<CAS_amount2>" & varXML(6) & "</CAS_amount2>" & "<CAS_quantity2>" & varXML(7) & "</CAS_quantity2>"

If boolv5010 Then
 If UBound(varXML) < 20 Then ReDim Preserve varXML(20)
 StringTemp = StringTemp & "<CAS_claim_reason3>" & varXML(8) & "</CAS_claim_reason3>" _
& "<CAS_amount3>" & varXML(9) & "</CAS_amount3>" & "<CAS_quantity3>" & varXML(10) & "</CAS_quantity3>" _
& "<CAS_claim_reason4>" & varXML(11) & "</CAS_claim_reason4>" & "<CAS_amount4>" & varXML(12) & "</CAS_amount4>" & "<CAS_quantity4>" & varXML(13) & "</CAS_quantity4>" _
& "<CAS_amount5>" & varXML(14) & "</CAS_amount5>" & "<CAS_quantity5>" & varXML(15) & "</CAS_quantity5>" _
& "<CAS_claim_reason6>" & varXML(16) & "</CAS_claim_reason6>" & "<CAS_amount6>" & varXML(17) & "</CAS_amount6>" & "<CAS_quantity6>" & varXML(18) & "</CAS_quantity6>"

End If
CASconv = StringTemp
End Function

Public Function NM1conv(varXML As Variant) As String
Dim StringTemp As String
If UBound(varXML) < 10 Then ReDim Preserve varXML(10)
StringTemp = "<NM1_entity_id>" & varXML(1) & "</NM1_entity_id><NM1_entity_qualifier>" _
& varXML(2) & "</NM1_entity_qualifier><NM1_name>" & Replace09(CStr(varXML(3)), "&", "and", , , vbTextCompare) _
& "</NM1_name><NM1_firstname>" & Replace09(CStr(varXML(4)), "&", "and", , , vbTextCompare) _
& "</NM1_firstname><NM1_middlename>" & varXML(5) & "</NM1_middlename><NM1_suffixname>" & varXML(7) _
& "</NM1_suffixname><NM1_id_qualifier>" & varXML(8) & "</NM1_id_qualifier><NM1_id_code>" & varXML(9) _
& "</NM1_id_code>"
If boolv5010 Then
ReDim Preserve varXML(13)
  StringTemp = StringTemp & "<NM1_Entity_relation>" & varXML(10) & "</NM1_Entity_relation><NM1_Entity_id_code>" _
                & varXML(11) & "</NM1_Entity_id_code><NM1_Org_name>" & varXML(12) & "</NM1_Org_name>"
End If
NM1conv = StringTemp
End Function

Public Function MIAconv(varXML As Variant) As String
Dim i As Variant
Dim tempstr As String
  
 For i = 1 To UBound(varXML)
  tempstr = tempstr & "<MIA_" & i & ">" & varXML(i) & "</MIA_" & i & ">"
  Next
MIAconv = tempstr
End Function

Public Function MOAconv(varXML As Variant) As String
Dim i As Variant
Dim tempstr As String
'tempstr = "<line_item>" & lnItem & "</line_item>"
 For i = 1 To UBound(varXML)
  tempstr = tempstr & "<MOA_" & i & ">" & varXML(i) & "</MOA_" & i & ">"
  Next
MOAconv = tempstr
End Function


Public Function AMTconv(varXML As Variant) As String
If UBound(varXML) < 2 Then ReDim Preserve varXML(2)
AMTconv = "<AMT_amount_qualifier>" & varXML(1) & "</AMT_amount_qualifier><AMT_amount>" & varXML(2) & "</AMT_amount>"

End Function

Public Function QTYconv(varXML As Variant) As String
If UBound(varXML) < 2 Then ReDim Preserve varXML(2)
QTYconv = "<QTY_quantity_qualifier>" & varXML(1) & "</QTY_quantity_qualifier><QTY_quantity>" & varXML(2) & "</QTY_quantity>"


End Function

Public Function SVCconv(varXML As Variant) As String
If UBound(varXML) < 7 Then ReDim Preserve varXML(7)
SVCconv = "<SVC_med_procedure_id>" & varXML(1) & "</SVC_med_procedure_id><SVC_amount>" _
& varXML(2) & "</SVC_amount><SVC_amount2>" & varXML(3) & "</SVC_amount2><SVC_service_id>" _
& varXML(4) & "</SVC_service_id><SVC_quantity>" & varXML(5) & "</SVC_quantity><SVC_med_procedure_id2>" & varXML(6) & "</SVC_med_procedure_id2>" _
& "<SVC_quantity2>" & varXML(7) & "</SVC_quantity2>"

End Function

Public Function LQconv(varXML As Variant) As String
If UBound(varXML) < 2 Then ReDim Preserve varXML(2)
    LQconv = "<LQ_qualifier>" & varXML(1) & "</LQ_qualifier><LQ_code>" & varXML(2) & "</LQ_code>"
End Function

Public Function PLBconv(varXML As Variant) As String
If UBound(varXML) < 14 Then ReDim Preserve varXML(14)
PLBconv = "<PLB_ref_id>" & varXML(1) & "</PLB_ref_id>" & "<PLB_date>" & varXML(2) & "</PLB_date>" & "<PLB_adj_id>" & varXML(3) & "</PLB_adj_id>" & "<PLB_amount>" & varXML(4) & "</PLB_amount>" & "<PLB_adj_id_2>" & varXML(5) & "</PLB_adj_id_2>" & "<PLB_amount2>" & varXML(6) & "</PLB_amount2>" & "<PLB_adj_id3>" & varXML(7) & "</PLB_adj_id3>" & "<PLB_amount3>" & varXML(8) & "</PLB_amount3>" & "<PLB_adj_id4>" & varXML(9) & "</PLB_adj_id4>" & "<PLB_amount4>" & varXML(10) & "</PLB_amount4>" & "<PLB_adj_id5>" & varXML(11) & "</PLB_adj_id5>" & "<PLB_amount5>" & varXML(12) & "</PLB_amount5>" & "<PLB_adj_id6>" & varXML(13) & "</PLB_adj_id6>" & "<PLB_amount6>" & varXML(14) & "</PLB_amount6>"
End Function

Public Function SEconv(varXML As Variant) As String
SEconv = "<SE_num_seg>" & varXML(1) & "</SE_num_seg>" & "<SE_control_num>" & GStrans & varXML(2) & "</SE_control_num>"
End Function




