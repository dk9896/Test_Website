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
ColArrey(1) = "ModelName"
ColArrey(2) = "ModelDesc"
ColArrey(3) = "testVoltage"
ColArrey(4) = "RPCheckTime"
ColArrey(5) = "STCheckTime"
ColArrey(6) = "Efficiency"
ColArrey(7) = "OutputVoltMin"
ColArrey(8) = "OutputVoltMax"
ColArrey(9) = "OutputCurrentMin"
ColArrey(10) = "OutputCurrentMax"
ColArrey(11) = "VoltageOffset"
ColArrey(12) = "CurrentOffset"
ColArrey(13) = "PrintPartNo"
ColArrey(14) = "PrintBarcodeLength"
ColArrey(15) = "BarcodeLength"
ColArrey(16) = "HardwareNo"
ColArrey(17) = "SerialStartingtxt"
ColArrey(18) = "VandorId"
ColArrey(19) = "DotMarkingTime"
    
ColArrey(20) = "ModelNo"
ColArrey(21) = "PartImage"
ColArrey(22) = "PrinterBypass"
ColArrey(23) = "Bypass1"
ColArrey(24) = "Bypass2"
ColArrey(25) = "Bypass3"
ColArrey(26) = "Bypass4"
ColArrey(27) = "Bypass5"
ColArrey(28) = "Bypass6"
ColArrey(29) = "Bypass7"
ColArrey(30) = "Bypass8"
ColArrey(31) = "batchcounter"
ColArrey(32) = "CouplerCounter"


For Row = 1 To 32
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
ColArrey(19) = "cycletime"
For Row = 1 To 19
    ColName = ColArrey(Row)
    If FieldExists(Con, TableName, ColName) = False Then
        X = CreateField(Con, TableName, ColName, "varchar(255) DEFAULT 0")
    End If
Next

    


'------------------------------------



TableName = "Model_Report_Counter"
ColArrey(1) = "ModelName"
ColArrey(2) = "DateTime"
ColArrey(3) = "ShiftTime"
ColArrey(4) = "Mailsent"
ColArrey(5) = "ModelNo"
ColArrey(6) = "ProductionCounter"
ColArrey(7) = "OKCounter"
ColArrey(8) = "NGCounter"
ColArrey(9) = "CouplerCounter"
ColArrey(10) = "BatchCounter"
ColArrey(11) = "TargetProduction"

For Row = 1 To 11
    ColName = ColArrey(Row)
    If FieldExists(Con, TableName, ColName) = False Then
        X = CreateField(Con, TableName, ColName, "varchar(255) DEFAULT 0")
    End If
Next

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
