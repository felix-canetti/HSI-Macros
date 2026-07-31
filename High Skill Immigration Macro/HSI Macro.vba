Sub importHSI_Questionnaire()

'This is the importing macro for High Skill Questionnaires
'Declaring the variables to be used
    Dim lastRow As Long
    Dim i As Integer
    Dim importFile As String
    Dim newFile As String
    'Get the current Workbook name and assign it to the importFile Variable
    importFile = ActiveWorkbook.Name
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
    lastrow = Ranfge("A" & Rows.Count).End(xlUp).Row

    'Renaming of the columns
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
     Range("")

End Sub