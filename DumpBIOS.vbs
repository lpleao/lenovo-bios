'
' Dump all BIOS settings
'
On Error Resume Next
Dim colItems

Set objFSO=CreateObject("Scripting.FileSystemObject")

Set args = Wscript.Arguments

strFile = WScript.Arguments.Item(0) & "\BIOSConfig.csv"
strComputer = "LOCALHOST"
'strOptions

WScript.Echo "BIOS file: >" & strFile & "<" & vbCrLf

Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_ComputerSystem")
For Each objItem in colItems
    strOEM = objItem.Manufacturer
    strModel = objItem.Model
next
WScript.Echo "OEM: " & strOEM
WScript.Echo "Model: " & strModel

Set colItems = objWMIService.ExecQuery("SELECT Version FROM Win32_ComputerSystemProduct")
For Each objItem in colItems
    strTypeDesc = objItem.Version
Next
WScript.Echo "Product: " & strTypeDesc

If UCase(strOEM) <> "LENOVO" Then
    WScript.Echo "Wrong system: " & strOEM
    WScript.Quit 1
End If
strType = Left(strModel, 4)
WScript.Echo "Type: " & strType

Set objFile = objFSO.CreateTextFile(strFile, True)
Set objWMIService = GetObject("WinMgmts:" & "{ImpersonationLevel=Impersonate}!\\" & strComputer & "\root\wmi")
Set colItems = objWMIService.ExecQuery("Select * from Lenovo_BiosSetting")

set intNr = 0
For Each objItem in colItems
    If Len(objItem.CurrentSetting) > 0 Then
        Setting = ObjItem.CurrentSetting
        StrItem = Left(ObjItem.CurrentSetting, InStr(ObjItem.CurrentSetting, ",") - 1)
        StrValue = Mid(ObjItem.CurrentSetting, InStr(ObjItem.CurrentSetting, ",") + 1, 256)
		
	'Set selItems = objWMIService.ExecQuery("Select * from Lenovo_GetBiosSelections")
	'For Each objItem2 in selItems
	'	objItem2.GetBiosSelections StrItem + ";", strOptions
	'Next

    'Wscript.Echo ">" & strValue & "<"	
    strValue = """" & replace(StrValue, ";[",""";""[") & """"
    If InStr(strValue, "[Status:ShowOnly]") > 0 Then
        strValue = Left(strValue, InStr(strValue, "[Status:ShowOnly]") - 1) & """"
    End If
    If InStr(strValue, "[Optional:") > 0 Then
        strValue = replace(replace(replace(strValue, """[", """"), "]""", """"), ":", """;""")
    End If
    If InStr(strValue, "[Excluded from boot order:") > 0 Then
        strValue = replace(replace(strValue, """[Excluded from boot order:", """Excluded from boot order"";"""), "]""", """")
    End If
    If InStr(strValue, ";") = 0 Then
        strValue = replace(replace(strValue, "[", ""), "]", "") & ";""Insert"";"""""
    End If
    If InStr(strValue, ";") = Len(strValue)-1 and InStr(StrItem, "Boot") > 0 Then
            strValue = replace(replace(replace(strValue, "[", ""), "]", ""), ";""", """;""Excluded from boot order"";""""")
    End If

    objFile.Write """" & StrType & """;""" & strTypeDesc & """;""" & StrModel & """;""" & FormatNumber(intNr, 0) & """;""" & StrItem & """;"
	objFile.Write strValue & vbCrLf
	'objFile.Write "  possible settings = " & strOptions & vbCrLf
    intNr = intNr + 1
    End If
Next

objFile.Close
