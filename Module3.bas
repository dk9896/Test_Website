Attribute VB_Name = "Module3"
Public Con100 As ADODB.Connection

Public Sub Conn_2007()
On Error GoTo Error
Dim StrMdbPath, StrConn As String

    StrMdbPath = App.Path & "\Database\" & App.Title & "_DB.mdb"
    StrConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & StrMdbPath & ";Jet OLEDB:Database Password=authentic;"
    Set Con100 = New ADODB.Connection
    Con100.Open StrConn

Exit Sub
Error:
MsgBox "Office 2007 Error", vbInformation
End Sub

Public Function CreateField(ByVal DBConn, strTable, strField, strType As String) As Boolean
Dim Sql As String
On Error GoTo error1
    If (strField <> vbNullString) Then
        Sql = "ALTER TABLE " & strTable & " ADD COLUMN " & strField & " " & strType
        DBConn.Execute Sql
        Sql = "UPDATE " & strTable & " Set " & strField & " = 0"
        DBConn.Execute Sql
        CreateField = True
    End If
Exit Function
error1:
MsgBox (Err.Description)

End Function

Public Function FieldExists(ByVal DBConn, TableName, FieldName As String) As Boolean
Dim Rs As New ADODB.Recordset
Dim FLD As ADODB.Field

Rs.Open TableName, DBConn, adOpenStatic, adLockReadOnly, adCmdTable
For Each FLD In Rs.Fields
    If LCase(FLD.Name) = LCase(FieldName) Then
        FieldExists = True
        Exit For
    End If
Next

End Function


Public Sub Make_Column()
On Error GoTo Error
Dim Sql As String
Dim TableName As String
Dim ColName As String
Dim Row As Long
Dim Col As Long
Dim ColArrey(100) As String

'----------------Table - Model_Set
TableName = "Model_Set"
'  "Printtype"
'    Rs ("IDNo")
'    Rs("LastPartNo") = txtLastPartno.Text
'    Rs("PartNo") = txtPartNo.Text
'    Rs("Darkness") = cboDarkness.ListIndex
'    Rs("RejectionBypass") = Check1.Value
'    Rs("Vendorcode") = txtvendorCode.Text
'    Rs("linecode") = txtlinecode.Text
ColArrey(1) = "DMBypass"
ColArrey(2) = "DM1CurMin"
ColArrey(3) = "DM1CurMax"
ColArrey(4) = "DM2CurMin"
ColArrey(5) = "DM2CurMax"
ColArrey(6) = "DM1VoltMin"
ColArrey(7) = "DM1VoltMax"
ColArrey(8) = "DM2VoltMin"
ColArrey(9) = "DM2VoltMax"
ColArrey(10) = "DMTestCycle"
ColArrey(11) = "PMBypass"
ColArrey(12) = "PM1CurMin"
ColArrey(13) = "PM1CurMax"
ColArrey(14) = "PM1VoltMin"
ColArrey(15) = "PM1VoltMax"
ColArrey(16) = "PMTestCycle"
ColArrey(17) = "BMBypass"
ColArrey(18) = "BM1CurMin"
ColArrey(19) = "BM1CurMax"
ColArrey(20) = "BM2CurMin"
ColArrey(21) = "BM2CurMax"
ColArrey(22) = "BM3CurMin"
ColArrey(23) = "BM3CurMax"
ColArrey(24) = "BM1VoltMin"
ColArrey(25) = "BM1VoltMax"
ColArrey(26) = "BM2VoltMin"
ColArrey(27) = "BM2VoltMax"
ColArrey(28) = "BM3VoltMin"
ColArrey(29) = "BM3VoltMax"
ColArrey(30) = "BMTestCycle"
ColArrey(31) = "NMBypass"
ColArrey(32) = "NM1CurMin"
ColArrey(33) = "NM1CurMax"
ColArrey(34) = "NM2CurMin"
ColArrey(35) = "NM2CurMax"
ColArrey(36) = "NM3CurMin"
ColArrey(37) = "NM3CurMax"
ColArrey(38) = "NM4CurMin"
ColArrey(39) = "NM4CurMax"
ColArrey(40) = "NM1VoltMin"
ColArrey(41) = "NM1VoltMax"
ColArrey(42) = "NM2VoltMin"
ColArrey(43) = "NM2VoltMax"
ColArrey(44) = "NM3VoltMin"
ColArrey(45) = "NM3VoltMax"
ColArrey(46) = "NM4VoltMin"
ColArrey(47) = "NM4VoltMax"
ColArrey(48) = "NMTestCycle"
ColArrey(49) = "HAMBypass"
ColArrey(50) = "HAM1CurMin"
ColArrey(51) = "HAM1CurMax"
ColArrey(52) = "HAM1VoltMin"
ColArrey(53) = "HAM1VoltMax"
ColArrey(54) = "HAMTestCycle"
ColArrey(55) = "HOMBypass"
ColArrey(56) = "HOM1CurMin"
ColArrey(57) = "HOM1CurMax"
ColArrey(58) = "HOM1VoltMin"
ColArrey(59) = "HOM1VoltMax"
ColArrey(60) = "HOMTestCycle"
ColArrey(61) = "CRMBypass"
ColArrey(62) = "CRM1CurMin"
ColArrey(63) = "CRM1CurMax"
ColArrey(64) = "CRM1VoltMin"
ColArrey(65) = "CRM1VoltMax"
ColArrey(66) = "CRMTestCycle"
ColArrey(67) = "CUMBypass"
ColArrey(68) = "CUM1CurMin"
ColArrey(69) = "CUM1CurMax"
ColArrey(70) = "CUM1VoltMin"
ColArrey(71) = "CUMTestCycle"
ColArrey(72) = "CUM1VoltMax"
ColArrey(73) = "SRMBypass"
ColArrey(74) = "SRM1CurMin"
ColArrey(75) = "SRM1CurMax"
ColArrey(76) = "SRM1VoltMin"
ColArrey(77) = "SRM1VoltMax"
ColArrey(78) = "SRMTestCycle"

ColArrey(79) = "ICMin"
ColArrey(80) = "ICMax"
'ColArrey(57) = "WVMin"
'ColArrey(58) = "WVMax"
ColArrey(81) = "Bypass1"
ColArrey(82) = "Bypass2"
ColArrey(83) = "Bypass3"
ColArrey(84) = "Bypass4"
ColArrey(85) = "Bypass5"
ColArrey(86) = "Bypass6"
ColArrey(87) = "Bypass7"
ColArrey(88) = "Bypass8"
ColArrey(89) = "Bypass9"
ColArrey(90) = "Bypass10"
ColArrey(91) = "Bypass11"
ColArrey(92) = "VendorId"

For Row = 1 To 92
    ColName = ColArrey(Row)
    If FieldExists(Con, TableName, ColName) = False Then
        X = CreateField(Con, TableName, ColName, "varchar(255) DEFAULT 0")
    End If
Next

'----------------Table - Model_Report
TableName = "Model_Report"

ColArrey(1) = "DM1Cur"
ColArrey(2) = "DM2Cur"
ColArrey(3) = "DM1Volt"
ColArrey(4) = "DM2Volt"
ColArrey(5) = "PM1Cur"
ColArrey(6) = "PM1Volt"
ColArrey(7) = "HAM1Cur"
ColArrey(8) = "HAM1Volt"
ColArrey(9) = "NM1Cur"
ColArrey(10) = "NM2Cur"
ColArrey(11) = "NM1Volt"
ColArrey(12) = "NM2Volt"
ColArrey(13) = "NM3Cur"
ColArrey(14) = "NM3Cur"
ColArrey(16) = "NM4Volt"
ColArrey(17) = "NM4Volt"
ColArrey(18) = "BM1Cur"
ColArrey(19) = "BM2Cur"
ColArrey(20) = "BM3Cur"
ColArrey(21) = "BM1Volt"
ColArrey(22) = "BM2Volt"
ColArrey(23) = "BM3Volt"
ColArrey(24) = "HOM1Cur"
ColArrey(25) = "HOM1Volt"
ColArrey(26) = "CUM1Cur"
ColArrey(27) = "CUM1Volt"
ColArrey(28) = "CRM1Cur"
ColArrey(29) = "CRM1Volt"
ColArrey(30) = "SRM1Cur"
ColArrey(31) = "SRM1Volt"
'ColArrey(19) = "IC"
ColArrey(32) = "IC"
ColArrey(33) = "DMResult"
ColArrey(34) = "PMResult"
ColArrey(35) = "HAMResult"
ColArrey(36) = "NMResult"
ColArrey(37) = "BMResult"
ColArrey(38) = "HOMResult"
ColArrey(39) = "CUMResult"
ColArrey(40) = "CRMResult"
ColArrey(41) = "SRMResult"

For Row = 1 To 41
    ColName = ColArrey(Row)
    If FieldExists(Con, TableName, ColName) = False Then
        X = CreateField(Con, TableName, ColName, "varchar(255) DEFAULT 0")
    End If
Next


'----------------Table - Model_Set
TableName = "Common_Set"
ColArrey(1) = "WebApiLink"
ColArrey(2) = "SenderEmail"
ColArrey(3) = "SenderPassword"
ColArrey(4) = "ToEmail1"
ColArrey(5) = "ToEmail2"
ColArrey(6) = "ToEmail3"
ColArrey(7) = "ToEmail4"
ColArrey(8) = "ToEmail5"
ColArrey(9) = "ToEmail6"
ColArrey(10) = "ToEmail7"
ColArrey(11) = "EmailBypass"
ColArrey(12) = "EmailBypass1"
ColArrey(13) = "EmailBypass2"
ColArrey(14) = "EmailBypass3"
ColArrey(15) = "EmailBypass4"
ColArrey(16) = "EmailBypass5"
ColArrey(17) = "EmailBypass6"
ColArrey(18) = "EmailBypass7"

For Row = 1 To 18
    ColName = ColArrey(Row)
    If FieldExists(Con, TableName, ColName) = False Then
        X = CreateField(Con, TableName, ColName, "varchar(255) DEFAULT 0")
    End If
Next

    


'------------------------------------



'TableName = "user_list"
'ColName = "AccessType"
'If FieldExists(Con100, TableName, ColName) = False Then
'    X = CreateField(Con100, TableName, ColName, "varchar(255) DEFAULT 0")
'End If
'-=========================================

'Sql = "create table Common_Set (ID Counter)"
'Con100.Execute Sql

'TableName = "Common_Set"
'ColName = "SetType"
'If FieldExists(Con100, TableName, ColName) = False Then
'    X = CreateField(Con100, TableName, ColName, "varchar(255) DEFAULT 0")
'End If
'
''Sql = "Update Common_Set Set SetType='CommonSet'"
''Con100.Execute Sql
'
'ColName = "ComPort1"
'If FieldExists(Con100, TableName, ColName) = False Then
'    X = CreateField(Con100, TableName, ColName, "varchar(255) DEFAULT 0")
'End If

'Con100.Close

Exit Sub
Error:
'Con100.Close
End Sub
