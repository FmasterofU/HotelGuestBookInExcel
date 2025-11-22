VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
'Redni broj
Cells(ActiveCell.Row, ColumnIndex("A")).Formula = "=RIGHT(""0000""&ROW()-1,5)"
'Datum do kada je odobren boravak u Republici Srbiji
Cells(ActiveCell.Row, ColumnIndex("J")).Value = "/"
'Vrsta pružene usluge
Cells(ActiveCell.Row, ColumnIndex("K")).Value = "smeštaj"
'Datum i èas dolaska
Cells(ActiveCell.Row, ColumnIndex("M")).Value = Now
'Read ID card
On Error GoTo ErrorHandler
Dim working_direcotry As String
Dim srbidreader_path As String
working_directory = ThisWorkbook.Sheets("Helper").Range("B2").Value
srbidreader_path = ThisWorkbook.Sheets("Helper").Range("B3").Value
output = RunExternalProgram(srbidreader_path, working_directory)
Set json = JsonConverter.ParseJson(output)
'Prezime i ime
Dim ime As String
Dim prezime As String
ime = json("eid_fixed_personal_data")("givenName")
prezime = json("eid_fixed_personal_data")("surname")
Cells(ActiveCell.Row, ColumnIndex("B")).Value = prezime + " " + ime
'Pol
Dim pol As String
pol = json("eid_fixed_personal_data")("sex")
Cells(ActiveCell.Row, ColumnIndex("C")).Value = pol
'Dan, mjesec i godina roðenja
Dim datum_rodjenja As String
datum_rodjenja = json("eid_fixed_personal_data")("dateOfBirth")
Cells(ActiveCell.Row, ColumnIndex("D")).Value = datum_rodjenja
'Mesto, opština i država stanovanja
Dim adresa_stanovanja As String
Dim mesto_stanovanja As String
Dim opstina_stanovanja As String
Dim drzava_stanovanja As String
drzava_stanovanja = json("eid_variable_personal_data")("state")
opstina_stanovanja = json("eid_variable_personal_data")("community")
mesto_stanovanja = json("eid_variable_personal_data")("place")
adresa_stanovanja = json("eid_variable_personal_data")("street") + " " + json("eid_variable_personal_data")("houseNumber") + "/" + json("eid_variable_personal_data")("houseLetter") + "/" + json("eid_variable_personal_data")("entrance") + "/" + json("eid_variable_personal_data")("floor") + "/" + json("eid_variable_personal_data")("apartmentNumber")
Cells(ActiveCell.Row, ColumnIndex("E")).Value = adresa_stanovanja + ", " + mesto_stanovanja + ", " + opstina_stanovanja + ", " + drzava_stanovanja
'Jedinstveni matièni broj graðanina
Dim jmbg As String
jmbg = json("eid_fixed_personal_data")("personalNumber")
jmbg = "0#0#0#0#0#0#0"
Cells(ActiveCell.Row, ColumnIndex("F")).Value = jmbg
'Državljanstvo
Dim drzavljanstvo As String
drzavljanstvo = "RS"
Cells(ActiveCell.Row, ColumnIndex("G")).Value = drzavljanstvo
'Mesto i država roðenja
Dim mesto_rodjenja As String
Dim drzava_rodjenja As String
mesto_rodjenja = json("eid_fixed_personal_data")("placeOfBirth")
drzava_rodjenja = json("eid_fixed_personal_data")("stateOfBirth")
Cells(ActiveCell.Row, ColumnIndex("H")).Value = mesto_rodjenja + ", " + drzava_rodjenja
'Vrsta, broj i datum izdavanja putne isprave
Dim vrsta_isprave As String
Dim broj_isprave As String
Dim datum_izdavanja As String
Dim ispravu_izdao As String
vrsta_isprave = json("eid_document_data")("documentType")
broj_isprave = json("eid_document_data")("docRegNo")
broj_isprave = "0#0#0#0#0"
datum_izdavanja = json("eid_document_data")("issuingDate")
ispravu_izdao = json("eid_document_data")("issuingAuthority")
Cells(ActiveCell.Row, ColumnIndex("I")).Value = vrsta_isprave + ", " + broj_isprave + ", " + datum_izdavanja + ", " + ispravu_izdao
'Datum isteka važenja putne isprave
Dim datum_isteka As String
datum_isteka = json("eid_document_data")("expiryDate")
Cells(ActiveCell.Row, ColumnIndex("O")).Value = datum_isteka

On Error GoTo 0 ' Reset error handling
Exit Sub
ErrorHandler:
    ' Display an error message to the user
    MsgBox "Error: " & Err.Description
    On Error GoTo 0 ' Reset error handling
End Sub

Function RunExternalProgram(ByVal programPath As String, ByVal workingDirectory As String) As String
    ' Run an external program from a specified path and capture its output
    Dim objShell As Object
    Dim objExec As Object
    Dim output As String
    Set objShell = VBA.CreateObject("WScript.Shell")
    
    ' Set the working directory for the shell
    objShell.CurrentDirectory = workingDirectory
    
    ' Execute the external program
    Set objExec = objShell.Exec(programPath)
    
    ' Wait for the process to complete
    Do While objExec.Status = 0
        VBA.DoEvents
    Loop
    
    ' Read the program output
    output = objExec.StdOut.ReadAll
    
    ' Check if the external program encountered an error
    If objExec.Status <> 1 Or objExec.exitCode <> 0 Then
        Err.Raise vbObjectError + 9999, , "Greska u citacu licnih karti, kartica nije prepoznata."
    End If
    
    RunExternalProgram = output
End Function


Function ColumnIndex(colLetter As String) As Long
    ' Convert column letter to column index (e.g., A=1, B=2, ..., Z=26)
    ColumnIndex = Range(colLetter & "1").Column
End Function

Private Sub CommandButton2_Click()
'Datum i èas odlaska
Cells(ActiveCell.Row, ColumnIndex("N")).Value = Now
End Sub
