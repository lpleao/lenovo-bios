'
' Apply BIOS Settings from CSV file
'
'On Error Resume Next
Dim colItems

If WScript.Arguments.Count <> 2 Then
    WScript.Echo "ApplyBIOS.vbs [CSV file] [password]"
    WScript.Quit 1
End If

Set objFSO = CreateObject("Scripting.FileSystemObject")
Dim lineArray
strComputer = "LOCALHOST"     ' Change as needed.
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_ComputerSystem")
For Each objItem in colItems
    strOEM = objItem.Manufacturer
    strModel = objItem.Model
next

If UCase(strOEM) <> "LENOVO" Then
    WScript.Echo "Wrong system: " & strOEM
    WScript.Quit 1
End If

Set inputFile = objFSO.OpenTextFile(Wscript.Arguments(0))
Set objWMIService = GetObject("WinMgmts:{ImpersonationLevel=Impersonate}!\\" & strComputer & "\root\wmi")
'Set colItems = objWMIService.ExecQuery("Select * from Lenovo_SetBiosSetting")

intSettings = 0
intFailed = 0
Do Until inputFile.AtEndOfStream
    lineArray = Split(inputFile.Readline, ";") 'store line in temp array
    fModel = mid(lineArray(2), 2, len(lineArray(2)) - 2)
    fSetting = mid(lineArray(4), 2, len(lineArray(4)) - 2)
    fValue = mid(lineArray(5), 2, len(lineArray(5)) - 2)
    fValueType = mid(lineArray(6), 2, len(lineArray(6)) - 2)
    fValueList = mid(lineArray(7), 2, len(lineArray(7)) - 2)
    if fValueType = "Excluded from boot order" then
        fValue = fValue + ";[Excluded from boot order:" + fValueList + "]"
        'WScript.Echo fValue
    end if
    If (fModel = strModel) And (fValueType <> "Insert") Then
        strRequest = fSetting + "," + fValue + "," + WScript.Arguments(1) + ",ascii,us;"
        'WScript.Echo strRequest
        Set colItems = objWMIService.ExecQuery("Select * from Lenovo_SetBiosSetting")
        strReturn = "error"
        For Each objItem in colItems
            ObjItem.SetBiosSetting strRequest, strReturn
        Next
        If strReturn = "Success" Then
            strReturn = "error"
            Set colItems = objWMIService.ExecQuery("Select * from Lenovo_SaveBiosSettings")
            For Each objItem in colItems
                ObjItem.SaveBiosSettings WScript.Arguments(1) + ",ascii,us;", strReturn
            Next
            If strReturn = "Success" Then
                intSettings = intSettings + 1
            Else
                Wscript.Echo " SET: " & fSetting & "," & fValue & ": " & strReturn
                intFailed = intFailed + 1
            End If
        Else
            Wscript.Echo "SAVE: " & fSetting & "," & fValue & ": " & strReturn
            intFailed = intFailed + 1
        End If
    End If
Loop
inputFile.Close
WScript.Echo "Success: " & FormatNumber(intSettings, 0)
WScript.Echo "Failed: " & FormatNumber(intFailed, 0)
WScript.Quit
