Attribute VB_Name = "Module1"

Public LoginUser, LoginCode As String
'Public Declare Function SetFileAttributes Lib "kernel32.dll" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long
'Public Graph_array() As Variant
'Public LoginID As Double
Public UDefault As Boolean
Public CopyLabel As Boolean
Public DataIn() As Byte
Public DataOut() As Byte
'Public DataOut1() As Byte
Public Con As ADODB.Connection
Public Con1 As ADODB.Connection

Public Converted_Int, ByteHigh_To_Convert, Converted_ByteLow, Converted_ByteHigh, ByteLow_TO_Convert, OutPacketSize, InPacketSize As Integer

Public PLcdata1() As Integer
Public Pcdata() As Integer

'Public FlagUnload, cycleon As Boolean
'Public SwTest_Date As Date
'Public Declare Sub keybd_event Lib "User32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
'Public Const VK_MENU = &H12
'Public Const VK_SNAPSHOT = &H2C
'Public Const KEYEVENTF_KEYUP = &H2
'Public LineNo As Integer
'Public OpCode As String
Public PrinterBypass As Boolean
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public ModelName As String
Public AccessType As String
'Barcode Related
Public barcode As String
Public RetryCount, Sec_10 As Integer
Public StdReadStartAddress, StdReadCount, ExtendedReadStartAddress, ExtendedReadCount As String
Public StdWriteStartAddress, StdWriteCount As String
Public PLcdata(8500) As Integer
Public ReadArray, WriteArray As Variant
Public CommandOn As Boolean
Public ExtendedRequired As Boolean
Public CommandType, WriteDelayCount, CVRead, CVExtPktNo, NoOfExtendedPackets, RejCnt As Integer
Public SQLpath As String
Public SQLbypass As Integer
Public SerialStartingtxt As String
Public PartNo As String
Public HardwareNo As String
Public PrinterName As String
Public BarcodeLength As Integer
Public TempReportDate As Date
Public VendorId As String
Public reporttype As Integer
Sub Initialise()
    RetryCount = 5
    Sec_10 = 20
    StdReadStartAddress = 100       'D100
    StdReadCount = 100            'D100-D149
    StdWriteStartAddress = 200     'D200
    StdWriteCount = 200              'D200-D249
    ExtendedReadStartAddress = 1000 'D1000
    ExtendedReadCount = 700 '960         'D1000-D1300
    NoOfExtendedPackets = 6         '
    ExtendedRequired = False           'Extended Required in Current Application
    WriteDelayCount = 1            'Data Write will be done after WriteDelayCount Read Cycles
    CVRead = 0
    CommandType = 1
    CommandOn = False
End Sub
Public Sub GetReadArray(ReadStartAddress, NoOfReadRegisters, ReadArray)
Dim i, j As Integer
Dim Readstream() As Byte

ReDim Readstream(1 To 21) As Byte
'Header + Network details
Readstream(1) = &H50
Readstream(2) = &H0
Readstream(3) = &H0
Readstream(4) = &HFF
Readstream(5) = &HFF
Readstream(6) = &H3      'Lower
Readstream(7) = &H0      'Higher
'No of data to be Sent
Readstream(8) = 12       'Lower
Readstream(9) = &H0      'High
'Reserved
Readstream(10) = &H0
Readstream(11) = &H0
'Command & SubCommand
Readstream(12) = &H1     'Read Command
Readstream(13) = &H4
Readstream(14) = &H0     'Sub Command
Readstream(15) = &H0
'Read start address
Readstream(16) = Val(ReadStartAddress) Mod 256
Readstream(17) = Val(ReadStartAddress) \ 256
Readstream(18) = &H0
'Read data typr
Readstream(19) = &HA8    'D*
'No of Read Registers
Readstream(20) = Val(NoOfReadRegisters) Mod 256
Readstream(21) = Val(NoOfReadRegisters) \ 256

ReadArray = Readstream()
'DoEvents

End Sub
Public Sub GetWriteArray(WriteStartAddress, NoOfWriteRegisters, WriteArray)
Dim i, j, ArraySize, DataToWrite As Integer
Dim Writestream() As Byte

ArraySize = ((NoOfWriteRegisters * 2) + 21)
ReDim Writestream(1 To ArraySize)

Writestream(1) = &H50
Writestream(2) = &H0
Writestream(3) = &H0
Writestream(4) = &HFF
Writestream(5) = &HFF
Writestream(6) = &H3      'Lower
Writestream(7) = &H0      'Higher

DataToWrite = 12 + (NoOfWriteRegisters * 2)
'No of data to be Sent
Writestream(8) = Val(DataToWrite) Mod 256      'Lower
Writestream(9) = Val(DataToWrite) \ 256   'High
'Reserved
Writestream(10) = &H0
Writestream(11) = &H0
'Command & SubCommand
Writestream(12) = &H1     'Read Command
Writestream(13) = &H14
Writestream(14) = &H0     'Sub Command
Writestream(15) = &H0

'Writestart address
Writestream(16) = Val(WriteStartAddress) Mod 256
Writestream(17) = Val(WriteStartAddress) \ 256
Writestream(18) = &H0
'Write data typr
Writestream(19) = &HA8    'D*
'No of Write Registers
Writestream(20) = Val(NoOfWriteRegisters) Mod 256
Writestream(21) = Val(NoOfWriteRegisters) \ 256

j = WriteStartAddress + NoOfWriteRegisters
K = 22
For i = WriteStartAddress To (j - 1)
    If PLcdata(i) < 0 Then Data = (65536 + PLcdata(i)) Else Data = PLcdata(i)
    Writestream(K) = Data Mod 256
    K = K + 1
    Writestream(K) = Data \ 256
    K = K + 1
Next
WriteArray = Writestream()
'DoEvents

End Sub

Public Sub MakeConn()
    Set Con = New ADODB.Connection
    Con.Provider = "Microsoft.Jet.OLEDB.4.0"
    Con.Open App.Path & "\Database\" & App.Title & "_DB.mdb", "admin", ""
End Sub
Public Sub SqlConn()
On Error GoTo Error
    Set Con1 = New ADODB.Connection
    
    Con1.Open SQLpath '"Provider=MSDASQL;DRIVER=Sql Server;SERVER=LAPTOP-SUSHANT\SQLEXPRESS; DATABASE=Exicom_Laser_Marking_19_1145; UID=sa; PWD=authentic;"
Exit Sub
Error:
MsgBox Error, vbInformation
End Sub
Public Function PrintLabel(lPrinter As JustPrinter)
On Error GoTo err1
Dim TempFileTextLine As String
Dim PrnFile As String
    If CopyLabel = True Then
        Counter = SerialStartingtxt & Format(frmPrintLabel.txtCopyNo.Text, "0000000000000000")
        CopyLabel = False
    Else
        Counter = SerialStartingtxt & Format(Val(frmMonitor.txtproductioncounter), "0000000000000000")
    End If
     PrnFile = "switch.prn"
     barcode = PartNo & Counter
    
    TempFileTextLine = ReadLabel(App.Path & "\PrnFiles\" & PrnFile)
    TempFileTextLine = Replace(TempFileTextLine, "#034#", HardwareNo)
    TempFileTextLine = Replace(TempFileTextLine, "#28511012044#", PartNo)
    TempFileTextLine = Replace(TempFileTextLine, "#JJWW0000000000001#", Counter)
    TempFileTextLine = Replace(TempFileTextLine, "#28511012044JJWW0000000000001034#", barcode)

    lPrinter.PrinterName = PrinterName
    lPrinter.PrintText TempFileTextLine
    lPrinter.EndJob
    CreateTempPrn TempFileTextLine


Exit Function
err1:
MsgBox "Error found in " & Err.Source & vbNewLine & Err.Description, vbCritical, "Printer Error"
End Function

Private Function ReadLabel(FileName As String) As String
    Open FileName For Binary As #1
    ReadLabel = Input(LOF(1), 1)
    Close #1
End Function

Public Function CreateTempPrn(PrnData As String)
On Error GoTo Error:
Dim FSO As New FileSystemObject
Dim iFile As String
Dim iFileNo As Integer

    'Auther: Naveen Soni
    'Contact: 8287330444
    iFile = App.Path & "\PrnFiles\TempPrn.prn"
    iFileNo = FreeFile
    If FSO.FileExists(iFile) = True Then
        Kill iFile

    End If

    Open iFile For Append As iFileNo
    Print #iFileNo, PrnData
    Close iFileNo

Exit Function
Error:
MsgBox Err.Description, vbInformation
End Function
Public Function GetData()
    If DataReceived = InPacketSize Then
       
       CheckSumReceive = 0
       For i = 0 To InPacketSize - 2
          CheckSumReceive = (CheckSumReceive + DataIn(i)) Mod 256
       Next i
       If DataIn(InPacketSize - 1) = CheckSumReceive Then
            If DataIn(0) = 2 And DataIn(InPacketSize - 2) = 3 Then
               ReDim PLcdata1(CInt((InPacketSize - 4) / 2))
               j = 0
               For i = 1 To InPacketSize - 3 Step 2
                   ByteLow_TO_Convert = Val(DataIn(i))
                   ByteHigh_To_Convert = Val(DataIn(i + 1))
                   Convert_Bytes_To_Int
                   PLcdata1(j) = Converted_Int
                   j = j + 1
               Next
            End If
       End If
    End If
End Function
Public Function SendData()
ReDim DataOut(OutPacketSize)
    Dim i, j As Integer
   DataOut(0) = 2
   DataOut(OutPacketSize - 2) = 3
    i = 1
   For j = 0 To (((OutPacketSize - 3) / 2) - 1)
       Int_to_Convert = Pcdata(j)
       Convert_Int_To_bytes
       DataOut(i) = Converted_ByteLow
       i = i + 1
       DataOut(i) = Converted_ByteHigh
       i = i + 1
   Next
      
   CheckSumSend = 0
   For i = 0 To OutPacketSize - 2
      CheckSumSend = (CheckSumSend + DataOut(i)) Mod 256
   Next i
      DataOut(OutPacketSize - 1) = CheckSumSend
      
   'For I = 0 To (OutPacketSize - 1)
    '   OutByte(I) = DataOut(I)
   'Next
End Function

Public Function Convert_Int_To_bytes()
Dim temp As Integer
If Int_to_Convert < 0 Then
   temp = 32768 - Abs(Int_to_Convert)
   Converted_ByteHigh = temp \ 256
   Converted_ByteHigh = Converted_ByteHigh Or &H80
   Converted_ByteLow = temp Mod 256
Else
   temp = Int_to_Convert
   Converted_ByteHigh = temp \ 256
   Converted_ByteLow = temp Mod 256
End If
End Function
Public Function Convert_Bytes_To_Int()
If ByteHigh_To_Convert > 127 Then
   ByteHigh_To_Convert = (ByteHigh_To_Convert) And (&H7F)
   Converted_Int = (ByteHigh_To_Convert * 256) + ByteLow_TO_Convert
   
   Converted_Int = 32768 - Converted_Int
   Converted_Int = 0 - Converted_Int
Else
   Converted_Int = (ByteHigh_To_Convert * 256) + ByteLow_TO_Convert
End If
End Function

Public Function getShift() As String
On Error GoTo Error
Dim sTime1, sTime2, sTime3, sTime4 As String
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim NowTime As String

    Sql = "Select * from Common_Set where SetType ='CommonSet'" 'SetType = Settings Type
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Con, adOpenDynamic, adLockOptimistic
    sTime1 = Rs("Shift1")
    sTime2 = Rs("Shift2")
    sTime3 = Rs("Shift3")
    sTime4 = Rs("Shift4")

    If sTime4 = "1" Then
        getShift = "04"
        Exit Function
    End If

    sTime1 = Format(TimeValue(sTime1), "hh:mm AM/PM")
    sTime2 = Format(TimeValue(sTime2), "hh:mm AM/PM")
    sTime3 = Format(TimeValue(sTime3), "hh:mm AM/PM")
    NowTime = Format(Time, "hh:mm AM/PM")
    TempReportDate = Format(Now(), "mm/dd/yyyy")
    'Time 1 and Time 2 Should Be < 24 hrs
    If (TimeValue(NowTime) >= TimeValue(sTime1)) And (TimeValue(NowTime) < TimeValue(sTime2)) Then
        getShift = "1"
    ElseIf (TimeValue(NowTime) >= TimeValue(sTime2)) And (TimeValue(NowTime) < TimeValue(sTime3)) Then
        getShift = "2"
    Else
        getShift = "3"
        If TimeValue(NowTime) < TimeValue(sTime1) Then
            TempReportDate = DateAdd("d", -1, TempReportDate)
        End If
    End If
    
    Exit Function
Error:
    MsgBox "Error Found in Time Shift Calculation" & vbNewLine & Error, vbInformation
End Function

Public Function ErrorLog(ByVal ErrNum, ErrDesc, ErrLine, ErrPlace, ErrHint As String)
On Error GoTo Error:
Dim FSO As New FileSystemObject
Dim iFile As String
Dim iFileNo As Integer
Dim TmpDate, TmpTime As String
    'Auther: Naveen Soni
    'Contact: 8287330444
    TmpDate = Format(Date, "dd-mm-yyyy")
    TmpTime = Format(Time, "hh:mm:ss AM/PM")

    iFile = App.Path & "\ErrorLog.csv"
    iFileNo = FreeFile

    If FSO.FileExists(iFile) = False Then
        Open iFile For Append As iFileNo
        Print #iFileNo, "ErrDate" & "," & "ErrTime" & "," & "ErrNumber" & "," & "Error" & "," & "ErrLine" & "," & "ErrPlace" & "," & "Hint"
        Close iFileNo
    End If

    ErrDesc = Replace(ErrDesc, ",", "-")
    ErrDesc = Replace(ErrDesc, vbCrLf, " ")
    Open iFile For Append As iFileNo
    Print #iFileNo, TmpDate & "," & TmpTime & "," & ErrNum & "," & ErrDesc & "," & ErrLine & "," & ErrPlace & "," & ErrHint
    Close iFileNo

Exit Function
Error:
MsgBox Err.Description, vbInformation
End Function

Public Sub AppVersion(frm As Form)
Dim AppVer As String

AppVer = Replace$(App.Title, "_", " ") & " - " & App.Major & "." & App.Minor & ".0." & App.Revision
frm.Caption = AppVer

End Sub



