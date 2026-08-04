Sub importHSI_Questionnaire()

'This is the importing macro for High Skill Questionnaires
'Declaring the variables to be used
    Dim lastRow As Long
    Dim i As Long
    Dim importFile As String
    Dim newFile As String
    'Get the current Workbook name and assign it to the importFile Variable
    importFile = ActiveWorkbook.Name

    Application.ScreenUpdating = False

    'Formatting to keep leading zeros for ZipCodes, SSN and I-94 Numbers
    'Zip Code Formatting
    Range("Y:Y,AF:AF,BH:BH,BS:BS,CD:CD,CP:CP,CZ::CZ").NumberFormat = "00000"
    'Formatting the SSNumber Fields
    Range("R:R").NumberFormat = "000000000"
    'Formatting any I-94 columns
    Range("AM:AM").NumberFormat = "00000000000"

    'Shift New Columns at the end to the right by 4 or so columns to make space for the page36 assignments at the end
    Columns("HR:HU").Insert Shift:=xlToRight

    'Deleting the first two row and defining the last row
    Rows("1:2").Delete
    lastRow = Range("A" & Rows.Count).End(xlUp).Row

    'Renaming of the columns
     Range("A1") = "PAGE36.QuestImp"
     Range("B1") = "Case Serial"
     Range("C1") = "Surname"
     Range("D1") = "Given Name"
     Range("E1") = "Middle Name" 'this column will be combined with the given name one, we do NOT currently track middle names on GW
     Range("F1") = "CONTACT2.Email"
     Range("G1") = "CONTACT2.Gender"
     Range("H1") = "CONTACT2.DOB"
     Range("I1") = "CONTACT2.CityBirth"
     Range("J1") = "CONTACT2.POB"
     Range("K1") = "CONTACT2.COB"
     Range("L1") = "CONTACT2.CitizenCty"
     Range("M1") = "CONTACT2.Marital"
     Range("N1") = "DEL"
     Range("O1") = "CONTACT2.PassportNo"
     Range("P1") = "CONTACT2.PExpiryDt"
     Range("Q1") = "CONTACT2.pptcntry"
     Range("R1") = "CONTACT2.SSN"
     Range("S1") = "CONTACT2.ANo"
     Range("T1") = "DEL"
     Range("U1") = "CONTACT2.USAdd1"
     Range("V1") = "CONTACT2.USAdd2"
     Range("W1") = "CONTACT2.USCity"
     Range("X1") = "CONTACT2.USState"
     Range("Y1") = "CONTACT2.USZip"
     Range("Z1") = "CONTACT2.PhoneNo"
     Range("AA1") = "CONTACT2.WrkPhone"
     Range("AB1") = "CONTACT2.FAdd1"
     Range("AC1") = "CONTACT2.FAdd2"
     Range("AD1") = "CONTACT2.FCity"
     Range("AE1") = "CONTACT2.FState"
     Range("AF1") = "CONTACT2.FZip"
     Range("AG1") = "CONTACT2.FCountry"
     Range("AH1") = "CONTACT2.ForHomePhn"
     Range("AI1") = "CONTACT2.ForWrkPhn"
     Range("AJ1") = "PAGE36.ResideUS"
     Range("AK1") = "PAGE36.LastEntDt"
     Range("AL1") = "PAGE36.LastEntryPl"
     Range("AM1") = "PAGE36.I94Number"
     Range("AN1") = "PAGE36.I94ExpDt"
     Range("AO1") = "PAGE36.VisaEntry"
     Range("AP1") = "PAGE36.CurStatus"
     Range("AQ1") = "PAGE36.VisaIssPl"
     Range("AR1") = "PAGE36.Entry1"
     Range("AS1") = "PAGE36.PresentExit"
     Range("AT1") = "PAGE36.VisaEnt1"
     Range("AU1") = "PAGE36.DenyPortEntry"
     Range("AV1") = "PAGE36.PortRefExp"
     Range("AW1") = "PAGE36.VisaRefuse"
     Range("AX1") = "PAGE36.VisaRefExp"
     Range("AY1") = "PAGE36.VisaRevoked"
     Range("AZ1") = "PAGE36.VisaCanExp"
     Range("BA1") = "PAGE36.PriorPet"
     Range("BB1") = "PAGE36.PriorPetExp"
     Range("BC1") = "PAGE36.PetType"
     Range("BD1") = "CONTACT2.UniName1"
     Range("BE1") = "CONTACT2.MastersAddr"
     Range("BF1") = "CONTACT2.MastersCity"
     Range("BG1") = "CONTACT2.MastersState"
     Range("BH1") = "CONTACT2.MastersZip"
     Range("BI1") = "CONTACT2.DegCntry"
     Range("BJ1") = "CONTACT2.HighDegree1"
     Range("BK1") = "CONTACT2.Concent1"
     Range("BL1") = "CONTACT2.Degree1Month"
     Range("BM1") = "CONTACT2.Degyear"
     Range("BN1") = "DEL"
     Range("BO1") = "CONTACT2.UniName2"
     Range("BP1") = "CONTACT2.MastersAddr22"
     Range("BQ1") = "CONTACT2.MastersCity2"
     Range("BR1") = "CONTACT2.MastersState2"
     Range("BS1") = "CONTACT2.MastersZip2"
     Range("BT1") = "CONTACT2.DgCntry2"
     Range("BU1") = "CONTACT2.HighDegree2"
     Range("BV1") = "CONTACT2.Concent12"
     Range("BW1") = "CONTACT2.Degree2Month"
     Range("BX1") = "CONTACT2.Degyear2"
     Range("BY1") = "DEL"
     Range("BZ1") = "CONTACT2.UniName3"
     Range("CA1") = "CONTACT2.MastersAddr13"
     Range("CB1") = "CONTACT2.MastersCity3"
     Range("CC1") = "CONTACT2.MastersState3"
     Range("CD1") = "CONTACT2.MastersZip3"
     Range("CE1") = "CONTACT2.DegCntry3"
     Range("CF1") = "CONTACT2.HighDegree3"
     Range("CG1") = "CONTACT2.Concent13"
     Range("CH1") = "CONTACT2.Degree3Month"
     Range("CI1") = "CONTACT2.Degyear3"
     Range("CJ1") = "DEL"
     Range("CK1") = "PAGE36.EmpName"
     Range("CL1") = "PAGE36.CurrEmpAdd"
     Range("CM1") = "PAGE36.CurrAdd2"
     Range("CN1") = "PAGE36.CurrEmpCity"
     Range("CO1") = "PAGE36.CurEmpState"
     Range("CP1") = "PAGE36.CurrEmpZip"
     Range("CQ1") = "PAGE36.PosTitle"
     Range("CR1") = "PAGE36.CurAnlSalary"
     Range("CS1") = "PAGE36.HrsWorked"
     Range("CT1") = "PAGE36.EmpSince"
     'Something to Note the rest of the previous employments fell underneat columns CU:DE
     Range("CU1") = "PAGE36.PrevEmpName2"
     Range("CV1") = "PAGE36.PrevEmpAdd2"
     Range("CW1") = "PAGE36.PrevAdd2"
     Range("CX1") = "PAGE36.PrevEmpCity2"
     Range("CY1") = "PAGE36.PrevEmpState2"
     Range("CZ1") = "PAGE36.PrevEmpZip2"
     Range("DA1") = "PAGE36.PrevEmpCntry3"
     Range("DB1") = "PAGE36.PrevPosTtl2"
     Range("DC1") = "PAGE36.EmpFrom2"
     Range("DD1") = "PAGE36.EmpTo2"
     Range("DE1") = "PAGE36.HrsWorked2"
     Range("DF1") = "CONTACT3.Dep1SurName"
     Range("DG1") = "CONTACT3.Dep1Fname"
     Range("DH1") = "CONTACT3.Dep1DOB"
     Range("DI1") = "CONTACT3.Dep1BirthCtry"
     Range("DJ1") = "DEL"
     'Something to note - the rest of the dependents ar tied up with columns DK:DO
     Range("DK1") = "CONTACT3.Dep2SurName"
     Range("DL1") = "CONTACT3.Dep2Fname"
     Range("DM1") = "CONTACT3.Dep2DOB"
     Range("DN1") = "CONTACT3.Dep2BirthCtry"
     Range("DO1") = "DEL"
     Range("DP1") = "PAGE36.NTAIssued"
     Range("DQ1") = "CONTACT2.ExclDep"
     Range("DR1") = "CONTACT2.BnDenied"
     Range("DS1") = "CONTACT2.J1J2Granted"
     Range("DT1") = "PAGE36.Crime"
     Range("DU1") = "PAGE36.CrimArrst"
     Range("DV1") = "PAGE36.ImmLie"
     Range("DW1") = "PAGE36.ImmLieExp"
     Range("DX1") = "PAGE36.SponsEmp"
     Range("DY1") = "PAGE36.WillSponsor"
     Range("DZ1") = "PAGE36.FieldScope"
     Range("EA1") = "PAGE36.ExpertDtl"
     Range("EB1") = "PAGE36.UniqueWork"
     Range("EC1") = "PAGE36.EconMerit"
     Range("ED1") = "PAGE36.NatlImp"
     Range("EE1") = "PAGE36.AdvanceND"
     Range("EF1") = "PAGE36.NatlBnft"
     Range("EG1") = "PAGE36.FuturePlan"
     Range("EH1") = "PAGE36.HasDegree"
     Range("EI1") = "PAGE36A.DegreeName"
     Range("EJ1") = "PAGE36A.IssueInst"
     Range("EK1") = "PAGE36A.DegreeDate"
     Range("EL1") = "PAGE36.ProjRole"
     Range("EM1") = "PAGE36B.ProjTtl"
     Range("EN1") = "PAGE36B.ProjRole"
     Range("EO1") = "PAGE36B.ProjInst"
     Range("EP1") = "PAGE36B.ProjStart"
     Range("EQ1") = "PAGE36B.ProjEnd"
     Range("ER1") = "PAGE36.HasPublish"
     Range("ES1") = "PAGE36C.PubTitle"
     Range("ET1") = "PAGE36C.PubJrnl"
     Range("EU1") = "PAGE36C.PubDate"
     Range("EV1") = "PAGE36.HasCite"
     Range("EW1") = "PAGE36D.CiteAuthor"
     Range("EX1") = "PAGE36D.CiteSource"
     Range("EY1") = "PAGE36D.CiteDate"
     Range("EZ1") = "PAGE36D.CiteDetail"
     Range("FA1") = "PAGE36.HasMedia"
     Range("FB1") = "PAGE36E.MediaOutlet"
     Range("FC1") = "PAGE36E.MediaTitle"
     Range("FD1") = "PAGE36E.MediaDate"
     Range("FE1") = "PAGE36E.MediaAuthor"
     Range("FF1") = "PAGE36.HasSpeak"
     Range("FG1") = "PAGE36F.SpeakEvent"
     Range("FH1") = "PAGE36F.SpeakTopic"
     Range("FI1") = "PAGE36F.SpeakDate"
     Range("FJ1") = "PAGE36F.SpeakLoc"
     Range("FK1") = "PAGE36.HasPatent"
     Range("FL1") = "PAGE36M.IpType"
     Range("FM1") = "PAGE36M.IpTitle"
     Range("FN1") = "PAGE36M.IPIssuer"
     Range("FO1") = "PAGE36M.IPDate"
     Range("FP1") = "PAGE36.HasAward"
     Range("FQ1") = "PAGE36H.AwardName"
     Range("FR1") = "PAGE36H.AwardOrg"
     Range("FS1") = "PAGE36H.AwardDate"
     Range("FT1") = "PAGE36H.AwardCrit"
     Range("FU1") = "PAGE36.HasLead"
     Range("FV1") = "PAGE36I.LeadRole"
     Range("FW1") = "PAGE36I.LeadOrg"
     Range("FX1") = "PAGE36I.LeadStart"
     Range("FY1") = "PAGE36I.LeadEnd"
     Range("FZ1") = "PAGE36I.LeadDesc"
     Range("GA1") = "PAGE36.HasFinPlan"
     Range("GB1") = "PAGE36.FinPlanExp"
     Range("GC1") = "PAGE36.HasInvest"
     Range("GD1") = "PAGE36.InvestExp"
     Range("GE1") = "PAGE36.Hascontrct"
     Range("GF1") = "PAGE36.ContrctExp"
     Range("GG1") = "PAGE36.HasGovLtr"
     Range("GH1") = "PAGE36.HasGovtSup"
     Range("GI1") = "PAGE36.GovSuppExp"
     Range("GJ1") = "PAGE36.NIWCriteria"
     Range("GK1") = "PAGE36.AwardYN"
     Range("GL1") = "PAGE36.AwardDtl"
     Range("GM1") = "PAGE36.AssocYn"
     Range("GN1") = "PAGE36.AssocDtl"
     Range("GO1") = "PAGE36K.RschDesc"
     Range("GP1") = "PAGE36K.RschGoal"
     Range("GQ1") = "PAGE36K.RschCndct"
     Range("GR1") = "PAGE36K.RschImpt"
     Range("GS1") = "PAGE36K.RschChlng"
     Range("GT1") = "PAGE36K.RschOvrcm"
     Range("GU1") = "PAGE36K.RschRes"
     Range("GV1") = "PAGE36K.RschRecog"
     Range("GW1") = "PAGE36K.RschRole"
     Range("GX1") = "PAGE36L.RefName"
     Range("GY1") = "PAGE36L.RefTitle"
     Range("GZ1") = "PAGE36L.RefOrg"
     Range("HA1") = "PAGE36L.RefWorked"
     Range("HB1") = "PAGE36L.REFExp"
     Range("HC1") = "DEL"
     Range("HD1") = "DEL"
     Range("HE1") = "DEL"
     Range("HF1") = "DEL"
     Range("HG1") = "DEL"
     Range("HH1") = "DEL"
     Range("HI1") = "DEL"
     Range("HJ1") = "DEL"
     Range("HK1") = "PAGE36.FormURL"
     Range("HL1") = "DEL"
     Range("HM1") = "DEL"
     Range("HN1") = "DEL"
     Range("HO1") = "DEL"
     Range("HP1") = "DEL"
     Range("HQ1") = "DEL"

     'Reformatting the Employer field to avoid shorthand name for most common petitioners
     For i = 2 To lastRow
        If UCase(Range("DX" & i).Value) = "TAT" Or UCase(Range("DX" & i).Value) = "TCS" Then
            Range("DX" & i).Value = "Tata Consultancy Services Limited"
        ElseIf UCase(Range("DX" & i).Value) = "MAS" Then
            Range("DX" & i).Value = "Mastek Inc."
        ElseIf UCase(Range("DX" & i).Value) = "CAP" Then
            Range("DX" & i).Value = "Capgemini America, Inc."
        End If
    Next i

    '=========================================================================================
    ' FLATTEN REPEATING GROUPS (CONTACT2 / CONTACT3 / unlettered PAGE36 columns only)
    '
    ' FormAssembly exports any *extra* occurrence of a repeatable field group (a 3rd, 4th...
    ' Previous Employer, a 3rd+ Dependent, a 2nd+ Entry/Exit/Visa) as an ADDITIONAL ROW
    ' directly under the case, reusing the SAME columns, and sharing the same "Response Id"
    ' (raw column HO - this column never shifts, even after the HR:HU insert above, so it's
    ' safe to reference here even though its header now reads "DEL").
    '
    ' This section finds those continuation rows, copies each extra occurrence into new
    ' columns appended after HU on the parent's own row, then deletes the leftover rows.
    '
    ' NOTE: This does NOT touch PAGE36A/B/C/D/E/F/H/I/K/L/M (the lettered sub-pages) since
    ' you already have a separate macro handling those.
    '=========================================================================================
    Call FlattenRepeatGroups(lastRow)

    Application.ScreenUpdating = True

End Sub

Private Sub FlattenRepeatGroups(ByVal lastRow As Long)

    Const RESPONSE_ID_COL As String = "HO"   'Response Id - fixed position, safe to hardcode
    Const CASE_COL As String = "B"           '"Case Serial" - populated on parent rows, blank on continuation rows

    Dim nextCol As Long
    nextCol = Range("HU1").Column + 1        'start appending right after the reserved HR:HU gap

    ' --- Group definitions ------------------------------------------------------------
    ' srcCol       = first column of the block that continuation rows reuse
    ' templates    = header name for each column in the block; "{N}" is replaced with the
    '                instance number. A template with no "{N}" (e.g. "DEL") is left as-is.
    ' firstExtra   = the first "extra" instance number found in continuation rows
    '                (i.e. one more than what's already on the parent row)
    ' lastExtra    = highest instance number to support (based on the max seen in your
    '                test file). If a real case ever needs more, just raise this number.
    ' zipIdx       = 0-based index within the block that should get "00000" zip formatting,
    '                or -1 if none.
    Dim groups As Variant
    groups = Array( _
        MakeGroup("AR", Array("PAGE36.Entry{N}", "PAGE36.Exit{N}", "PAGE36.VisaEnt{N}"), _
                  2, 15, -1), _
        MakeGroup("CU", Array("PAGE36.PrevEmpName{N}", "PAGE36.PrevEmpAdd{N}", "PAGE36.PrevAdd{N}", _
                               "PAGE36.PrevEmpCity{N}", "PAGE36.PrevEmpState{N}", "PAGE36.PrevEmpZip{N}", _
                               "PAGE36.PrevEmpCntry{N}", "PAGE36.PrevPosTtl{N}", "PAGE36.EmpFrom{N}", _
                               "PAGE36.EmpTo{N}", "PAGE36.HrsWorked{N}"), _
                  3, 10, 5), _
        MakeGroup("DK", Array("CONTACT3.Dep{N}SurName", "CONTACT3.Dep{N}Fname", "CONTACT3.Dep{N}DOB", _
                               "CONTACT3.Dep{N}BirthCtry", "DEL"), _
                  3, 5, -1) _
    )

    Dim g As Long
    For g = LBound(groups) To UBound(groups)
        nextCol = ProcessGroup(groups(g), lastRow, RESPONSE_ID_COL, CASE_COL, nextCol)
    Next g

    ' --- Delete the now-empty continuation rows (single batched delete) ---------------
    Dim delRange As Range
    Dim r As Long
    For r = 2 To lastRow
        If Trim(Range(CASE_COL & r).Value & "") = "" Then
            If delRange Is Nothing Then
                Set delRange = Range(CASE_COL & r)
            Else
                Set delRange = Union(delRange, Range(CASE_COL & r))
            End If
        End If
    Next r
    If Not delRange Is Nothing Then delRange.EntireRow.Delete

End Sub

Private Function MakeGroup(srcCol As String, templates As Variant, firstExtra As Long, lastExtra As Long, zipIdx As Long) As Variant
    MakeGroup = Array(srcCol, templates, firstExtra, lastExtra, zipIdx)
End Function

'Processes one repeat group: writes the new headers, walks the sheet copying each extra
'occurrence onto its parent row, and returns the next free column for the following group.
Private Function ProcessGroup(grp As Variant, ByVal lastRow As Long, respCol As String, caseCol As String, ByVal startCol As Long) As Long

    Dim srcCol As String: srcCol = grp(0)
    Dim templates As Variant: templates = grp(1)
    Dim firstExtra As Long: firstExtra = grp(2)
    Dim lastExtra As Long: lastExtra = grp(3)
    Dim zipIdx As Long: zipIdx = grp(4)

    Dim width As Long: width = UBound(templates) - LBound(templates) + 1
    Dim srcColNum As Long: srcColNum = Range(srcCol & "1").Column

    Dim inst As Long, w As Long, destCol As Long

    ' Write the new headers for every extra instance
    For inst = firstExtra To lastExtra
        For w = 0 To width - 1
            destCol = startCol + (inst - firstExtra) * width + w
            Cells(1, destCol).Value = Replace(CStr(templates(w)), "{N}", CStr(inst))
        Next w
        If zipIdx >= 0 Then
            Cells(2, startCol + (inst - firstExtra) * width + zipIdx).EntireColumn.NumberFormat = "00000"
        End If
    Next inst

    ' Walk the rows: track the current parent, and for each continuation row belonging to
    ' that parent, check whether THIS group has data in it; if so, copy it into the next
    ' available instance slot on the parent row.
    Dim r As Long
    Dim parentRow As Long
    Dim parentResponseId As String
    Dim curInstance As Long
    Dim hasData As Boolean

    parentRow = 0
    For r = 2 To lastRow
        If Trim(Range(caseCol & r).Value & "") <> "" Then
            'This is a parent (case) row
            parentRow = r
            parentResponseId = Range(respCol & r).Value & ""
            curInstance = firstExtra
        Else
            'This is a continuation row - only use it if its Response Id matches the current parent
            If parentRow > 0 And curInstance <= lastExtra Then
                If (Range(respCol & r).Value & "") = parentResponseId Then
                    hasData = False
                    For w = 0 To width - 1
                        If Trim(CStr(Cells(r, srcColNum + w).Value) & "") <> "" Then
                            hasData = True
                            Exit For
                        End If
                    Next w
                    If hasData Then
                        For w = 0 To width - 1
                            Cells(parentRow, startCol + (curInstance - firstExtra) * width + w).Value = Cells(r, srcColNum + w).Value
                        Next w
                        curInstance = curInstance + 1
                    End If
                End If
            End If
        End If
    Next r

    ProcessGroup = startCol + (lastExtra - firstExtra + 1) * width

End Function
