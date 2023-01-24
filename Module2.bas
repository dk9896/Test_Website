Attribute VB_Name = "Module2"
Option Explicit
Public MyDev As Object
'Public MyDev As PSF090360AAMLXDevice
Public Const StrNotConnected As String = "Software is not connected. Press Connect button to try to connect"
Public Melexis_Com As Integer

Public Function CreateDevice360(bAutomatic As Boolean) As Integer
    Dim PSFMan As PSF090360AAMLXManager, DevicesCol As ObjectCollection, i As Long
    Set MyDev = New PSF090360AAMLXDevice
    
    'On Error GoTo lError
    Call CleanUp
    
    If bAutomatic Then
        ' Automatic device scanning begins here
        Set PSFMan = CreateObject("MPT.PSF090360AAMLXManager")
        Set DevicesCol = PSFMan.ScanStandalone(dtSerial, SettingBaudrate)
        If DevicesCol.Count <= 0 Then
'            MsgBox ("No PTC04 programmers were found!")
            CreateDevice360 = 0
            Exit Function
        End If

        If DevicesCol.Count > 1 Then
            For i = 1 To DevicesCol.Count - 1
                'We are responsible to call Destroy(True) on device objects we do not need
                Call DevicesCol(i).Destroy(True)
            Next i
        End If
        Set MyDev = DevicesCol(0)
    Else
        ' Manual connection begins here
        Set MyDev = CreateObject("MPT.PSF090360AAMLXDevice")
        
        Call MyDev.ConnectChannel(CVar(Melexis_Com), dtSerial)
        
        ' Check if PTC04 programmer is connected to this channel
        Call MyDev.CheckSetup(False)
    End If
    'MyDev.Advanced.chipversion = 34         ' 90360 = 90316BCK
'    MsgBox ("PTC04 programmer found on " & MyDev.Channel.Name)
    CreateDevice360 = 1
    Exit Function

lError:
    Set MyDev = Nothing
    MsgBox Err.Description
    Err.Clear
End Function

'Public MyDev As PSF090360AAMLXDevice

Public Function CreateDevice365(bAutomatic As Boolean) As Integer
Dim PSFMan As PSF090365AAMLXManager, DevicesCol As ObjectCollection, i As Long
Dim CommMan As CommManager, Chan As MPTChannel
On Error GoTo lError
Set MyDev = New PSF090365AAMLXDevice
If bAutomatic Then
    ' Automatic device scanning begins here
    Set PSFMan = CreateObject("MPT.PSF090365AAMLXManager")
    Set DevicesCol = PSFMan.ScanStandalone(dtSerial)
    If DevicesCol.Count <= 0 Then
    '    MsgBox ("No PTC-04 programmers found!")
        CreateDevice365 = 0
        Exit Function
    End If

    If DevicesCol.Count > 1 Then
        For i = 1 To DevicesCol.Count - 1
        'We are responsible to call Destroy(True) on device objects we do not need
        Call DevicesCol(i).Destroy(True)
        Next i
    End If
    Set MyDev = DevicesCol(0)
Else
    ' Manual connection begins here
    Set CommMan = CreateObject("MPT.CommManager")
    Set MyDev = CreateObject("MPT.PSF090365AAMLXDevice")

    Set Chan = CommMan.Channels.CreateChannel(CVar(Melexis_Com), ctSerial)
    MyDev.Channel = Chan
    ' Check if a PTC04 programmer is connected to this channel
    Call MyDev.CheckSetup(False)
End If
    '       MsgBox (MyDev.Name & " programmer found on " & MyDev.Channel.Name)
    CreateDevice365 = 1

Exit Function
lError:
MsgBox Err.Description
CreateDevice365 = 0
Err.Clear
End Function

''Sub CleanUp()
'' On Error Resume Next
''    Dim Man As CommManager
''
''    Set Man = CreateObject("MPT.CommManager")
''    If Not (MyDev Is Nothing) Then
''        ' Must call Destroy(True) to inform the object to prepare for shutdown
''        Call MyDev.Destroy(True)
''        Set MyDev = Nothing
''    End If
''    If Not (Man Is Nothing) Then
''        Man.Quit
''        Set Man = Nothing
''    End If
'''Exit Sub
'''Error:
'''ErrorLog Err.Number, Err.Description, Erl, "CleanUp", "CleanUp"
'''Resume Next
''End Sub
Sub CleanUp()
 On Error Resume Next
    Dim Man As CommManager

    Set Man = CreateObject("MPT.CommManager")
    If Not (MyDev Is Nothing) Then
        ' Must call Destroy(True) to inform the object to prepare for shutdown
        Call MyDev.Destroy(True)
        Set MyDev = Nothing
    End If
    If Not (Man Is Nothing) Then
        Man.Quit
        Set Man = Nothing
    End If
End Sub
Function hex2long(Str As String) As Long
    hex2long = CLng("&H" & Str)
End Function

Function long2hex(l As Long) As String
    long2hex = Hex$(l)
End Function



