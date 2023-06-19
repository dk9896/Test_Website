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
Dim rs As New ADODB.Recordset
Dim FLD As ADODB.Field

rs.Open TableName, DBConn, adOpenStatic, adLockReadOnly, adCmdTable
For Each FLD In rs.Fields
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
ColArrey(3) = "CutoffVolt"
ColArrey(4) = "OutputVolt1"
ColArrey(5) = "OutputVolt2"
ColArrey(6) = "OutputVolt3"
ColArrey(7) = "testVoltage"
ColArrey(8) = "testCurrent"
ColArrey(9) = "EfficiencyMin"
ColArrey(10) = "EfficiencyMax"
ColArrey(11) = "InputCurrentMin"
ColArrey(12) = "InputCurrentMax"
ColArrey(13) = "OutputVoltMin"
ColArrey(14) = "OutputVoltMax"
ColArrey(15) = "OutputCurrentMin"
ColArrey(16) = "OutputCurrentMax"
ColArrey(17) = "VoltageOffset"
ColArrey(18) = "CurrentOffset"
ColArrey(19) = "PrintPartNo"
ColArrey(20) = "HardwareNo"
ColArrey(21) = "SerialStartingtxt"
ColArrey(22) = "VandorId"
ColArrey(23) = "DotMarkingTime"
ColArrey(24) = "ModelNo"
ColArrey(25) = "PartImage"
ColArrey(26) = "CutoffVoltMin"
ColArrey(27) = "OutputVolt1Min"
ColArrey(28) = "OutputVolt2Min"
ColArrey(29) = "OutputVolt3Min"
ColArrey(30) = "CutoffVoltMax"
ColArrey(31) = "OutputVolt1Max"
ColArrey(32) = "OutputVolt2Max"
ColArrey(33) = "OutputVolt3Max"
ColArrey(34) = "PrinterBypass"
ColArrey(35) = "EfficiencyOffset"
ColArrey(36) = "InputVoltageOffset"
ColArrey(37) = "InputCurrentOffset"
For i = 0 To 16
    ColArrey(38 + i) = "Bypass" & i + 1
Next

For Row = 1 To 54
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
ColArrey(20) = "Break1Enable"
ColArrey(21) = "Break2Enable"
ColArrey(22) = "Break3Enable"
ColArrey(23) = "Break4Enable"
ColArrey(24) = "Break5Enable"
ColArrey(25) = "Break1Start"
ColArrey(26) = "Break1End"
ColArrey(27) = "Break2Start"
ColArrey(28) = "Break2End"
ColArrey(29) = "Break3Start"
ColArrey(30) = "Break3End"
ColArrey(31) = "Break4Start"
ColArrey(32) = "Break4End"
ColArrey(33) = "Break5Start"
ColArrey(34) = "Break5End"
ColArrey(35) = "Shift1Start"
ColArrey(36) = "Shift1End"
ColArrey(37) = "Shift2Start"
ColArrey(38) = "Shift2End"
ColArrey(39) = "Shift3Start"
ColArrey(40) = "Shift3End"

For Row = 1 To 40
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

TableName = "Model_Report"
ColArrey(1) = "Barcode"
ColArrey(2) = "Result"
ColArrey(3) = "ReversePolarity"
ColArrey(4) = "CutOffVoltage"
ColArrey(5) = "Output1"
ColArrey(6) = "Output2"
ColArrey(7) = "Output3"
ColArrey(8) = "OutputShortTest"
ColArrey(9) = "CutOffVoltageStatus"
ColArrey(10) = "Output1Status"
ColArrey(11) = "Output2Status"
ColArrey(12) = "Output3Status"
ColArrey(13) = "TestVoltage"
ColArrey(14) = "InputCurrent"
ColArrey(15) = "OPVoltage"
ColArrey(16) = "OPCurrent"
ColArrey(17) = "Efficiency"
ColArrey(18) = "TestVoltageStatus"
ColArrey(19) = "InputCurrentStatus"
ColArrey(20) = "OPVoltageStatus"
ColArrey(21) = "OPCurrentStatus"
ColArrey(22) = "EfficiencyStatus"

For Row = 1 To 22
    ColName = ColArrey(Row)
    If FieldExists(Con, TableName, ColName) = False Then
        X = CreateField(Con, TableName, ColName, "varchar(255) DEFAULT 0")
    End If
Next
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
MsgBox Err.Description
End Sub
