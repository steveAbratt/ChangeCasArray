'###################################################
'# ChangeOutlookCASSimple
'# Written By: Brian Clark, Software Consulting Services
'# Last Modified: 06/26/2015
'#
'# Detects Outlook version and changes its Exchange Client Access Server (CAS)
'# For: Outlook 2003 - 2013
'###################################################
'# Program Wide non constant variables
Dim TARGET_CAS_FQDN			' FQDN OF THE CAS SERVER YOU WANT OUTLOOK TO CONNECT TO
Dim TARGET_CAS_NETBIOS		' NETBIOS NAME OF THE CAS SERVER YOU WANT OUTLOOK TO CONNECT TO
' ***** SET THE BELOW VALUE TO WHAT YOU WANT YOUR CLIENT ACCESS SERVER TO BE *****
TARGET_CAS_FQDN = "MX3.otb.local"
TARGET_CAS_NETBIOS = "MX3"
'# DO NOT CHANGE ANYTHING BELOW THIS LINE!
'###################################################
Dim arrTargetCASBinary	' TARGET_CAS String as a Binary Byte Array for inserting into Registry
Dim OutlookVersion		' Version of outlook we are targeting for backup.
Dim OutlookRegPath		' Registry path for the version of outlook we are targeting for backup
Dim profileRegPath		' Registry Path for Outlook Profile
'# Constant Variables
Const HKCU = &H80000001 'HKEY_CURRENT_USER
Const CAS_REG_BINARY 	= "001f662a"	' Registry Key Binary Value Data that holds CAS Server FQDN
Const CAS_REG_NETBIOS1 	= "001e660c"	' Registry Key String Value Data that holds NETBIOS name of CAS (1 of 2 Values that hold this)
Const CAS_REG_NETBIOS2 	= "001e6602" 	' Registry Key String Value Data that holds NETBIOS name of CAS (2 of 2 Values that hold this)
Const CAS_REG_LDAPDN1 	= "001e6614"	' Registry Key String Value Data that holds Server LDAP DN Path (1 of 2 Values that hold this)
Const CAS_REG_LDAPDN2 	= "001e6612"	' Registry Key String Value Data that holds Server LDAP DN Path (2 of 2 Values that hold this)
Const CAS_REG_FQDN 		= "001e6608"	' Registry Key String Value Data that holds CAS FQDN
'###################################################
' MAIN
'###################################################

' Get what version of outlook we are using. Prompt user if we detect multiple versions.
GetOutlookVersions()
' generate binary data for what we want the new CAS FQDN to be.
arrTargetCASBinary = StringToRegBinary(TARGET_CAS_FQDN)

' Outlook 2013 uses a different reg path than 2003-2010
If Not OutlookVersion = "2013" Then
	profileRegPath = "Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles"
Else
	profileRegPath = "Software\Microsoft\Office\15.0\Outlook\Profiles"
End If

' Change all registry settings related to the Exchange Client Access Server
ChangeOutlookCAS()
EndProgram "Program exited cleanly. CAS was changed to " & TARGET_CAS_FQDN, TRUE

'#############################
' Subs
'#############################

Sub GetOutlookVersions()
	Dim fCount : fCount = 0 ' versions found
	Dim strFound : strFound = "" ' string of all versions found
	Dim strRegPath : strRegPath = "" ' registry path for latest version found
	Dim strVersion : strVersion = "" ' version of latest copy of outlook found
	Dim ibResult : ibResult = "" ' for input box results should multiple office versions exist
	Dim loopCount : loopCount = 0 ' loop counter, we escape the loop after 3 cycles of invalid user input

	if KeyExists(HKCU, "Software\Microsoft\Office\11.0\Outlook") Then
		fCount = fCount + 1
		strFound = strFound + " 2003"
		strRegPath = "11.0"
		strVersion = "2003"
	End If
	if KeyExists(HKCU, "Software\Microsoft\Office\12.0\Outlook") Then
		fCount = fCount + 1
		strFound = strFound + " 2007"
		strRegPath = "12.0"
		strVersion = "2007"
	End If
	if KeyExists(HKCU, "Software\Microsoft\Office\14.0\Outlook") Then
		fCount = fCount + 1
		strFound = strFound + " 2010"
		strRegPath = "14.0"
		strVersion = "2010"
	End If
	if KeyExists(HKCU, "Software\Microsoft\Office\15.0\Outlook") Then
		fCount = fCount + 1
		strFound = strFound + " 2013"
		strRegPath = "15.0"
		strVersion = "2013"
	End If

	if fCount = 0 Then ' no version of outlook detected
		EndProgram "Outlook 2003 - 2013 not detected on this system.", FALSE
	Elseif fCount = 1 Then ' single instance of outlook detected
		OutlookRegPath = "Software\Microsoft\Office\" + strRegPath + "\Outlook"
		OutlookVersion = strVersion
		MsgBox "Success!" + vbNewLine + "Office Version: " + OutlookVersion + vbNewLine + "Registry Path: HKCU\" + OutlookRegPath + vbNewLine + _
			"Click OK to proceed!",64,"Change CAS Tool"
		Exit Sub
	Else ' more than 1 version of outlook detected
		Do Until loopCount > 2
			ibResult = InputBox("Multiple versions of Microsoft Office were detected." + vbNewLine + "Detected versions:" + strFound + vbNewLine + _
				"Please input which version you want to backup by typing the version year " + vbNewLine + "Valid Answers: 2003, 2007, 2010, and 2013" + vbNewLine)
			if ibResult = "2003" Then
				OutlookRegPath = "Software\Microsoft\Office\11.0\Outlook"
				OutlookVersion = "2003"
				MsgBox "You Selected Microsoft Office " + OutlookVersion + vbNewLine + "Registry Path: HKCU\" + OutlookRegPath + vbNewLine + _
					"Click OK to proceed with backup!",64,"Change CAS Tool"
				Exit Sub
			Elseif ibResult = "2007" Then
				OutlookRegPath = "Software\Microsoft\Office\12.0\Outlook"
				OutlookVersion = "2007"
				MsgBox "You Selected Microsoft Office " + OutlookVersion + vbNewLine + "Registry Path: HKCU\" + OutlookRegPath + vbNewLine + _
					"Click OK to proceed with backup!",64,"Change CAS Tool"
				Exit Sub
			Elseif ibResult = "2010" Then
				OutlookRegPath = "Software\Microsoft\Office\14.0\Outlook"
				OutlookVersion = "2010"
				MsgBox "You Selected Microsoft Office " + OutlookVersion + vbNewLine + "Registry Path: HKCU\" + OutlookRegPath + vbNewLine + _
					"Click OK to proceed with backup!",64,"Change CAS Tool"
				Exit Sub
			Elseif ibResult = "2013" Then
				OutlookRegPath = "Software\Microsoft\Office\15.0\Outlook"
				OutlookVersion = "2013"
				MsgBox "You Selected Microsoft Office " + OutlookVersion + vbNewLine + "Registry Path: HKCU\" + OutlookRegPath + vbNewLine + _
					"Click OK to proceed with backup!",64,"Change CAS Tool"
				Exit Sub
			Else
				loopCount = loopCount + 1
			End If
		Loop
		EndProgram "user failed to choose a valid office version.", FALSE
	End If
End Sub

Sub ChangeOutlookCAS()
	Dim oReg : Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
	Dim oShell : Set oShell = CreateObject("WScript.Shell")
	Dim sPath, aSub, sKey, aChildSub, sChildKey, binValue, output, k

	' Get all keys within the Outlook Profile Registry Key
	oReg.EnumKey HKCU, profileRegPath, aSub

	' Loop through each key
For Each sKey In aSub
    ' Get all subkeys within the key 'sKey'
    oReg.EnumKey HKCU, profileRegPath & "\" & sKey, aChildSub
    For Each sChildKey In aChildSub
        ' Set Binary Data for CAS 
		oReg.GetBinaryValue HKCU, profileRegPath & "\" & sKey & "\" & "\" & sChildKey, CAS_REG_BINARY, binValue
		If Not IsNull(binValue) Then
		output = ""
		oReg.SetBinaryValue HKCU, profileRegPath & "\" & sKey & "\" & "\" & sChildKey, CAS_REG_BINARY, arrTargetCASBinary
		output = ""
		End If
		oReg.GetStringValue HKCU, profileRegPath & "\" & sKey & "\" & "\" & sChildKey, CAS_REG_FQDN, output
		if Not IsNull(output) Then 
			oReg.SetStringValue HKCU, profileRegPath & "\" & sKey & "\" & "\" & sChildKey, CAS_REG_FQDN, TARGET_CAS_FQDN
			output = ""
		End If
		oReg.GetStringValue HKCU, profileRegPath & "\" & sKey & "\" & "\" & sChildKey, CAS_REG_NETBIOS1, output
		if Not IsNull(output) Then 
			oReg.SetStringValue HKCU, profileRegPath & "\" & sKey & "\" & "\" & sChildKey, CAS_REG_NETBIOS1, TARGET_CAS_NETBIOS
			output = ""
		End If
		oReg.GetStringValue HKCU, profileRegPath & "\" & sKey & "\" & "\" & sChildKey, CAS_REG_NETBIOS2, output
		if Not IsNull(output) Then 
			oReg.SetStringValue HKCU, profileRegPath & "\" & sKey & "\" & "\" & sChildKey, CAS_REG_NETBIOS2, TARGET_CAS_NETBIOS
			output = ""
		End If
		oReg.GetStringValue HKCU, profileRegPath & "\" & sKey & "\" & "\" & sChildKey, CAS_REG_LDAPDN1, output
		if Not IsNull(output) Then 
			oReg.SetStringValue HKCU, profileRegPath & "\" & sKey & "\" & "\" & sChildKey, CAS_REG_LDAPDN1, FormatNewDN(output)
			output = ""
		End If
		oReg.GetStringValue HKCU, profileRegPath & "\" & sKey & "\" & "\" & sChildKey, CAS_REG_LDAPDN2, output
		if Not IsNull(output) Then 
			oReg.SetStringValue HKCU, profileRegPath & "\" & sKey & "\" & "\" & sChildKey, CAS_REG_LDAPDN2, FormatNewDN(output)
		End If		
    Next
Next
End Sub


'#############################
' FUNCTIONS
'#############################

Function KeyExists(Key, KeyPath)
    Dim oReg: Set oReg = GetObject("winmgmts:!root/default:StdRegProv")
    If oReg.EnumKey(Key, KeyPath, arrSubKeys) = 0 Then
        KeyExists = True
    Else
        KeyExists = False
   End If
   Set oReg = Nothing
End Function

' used to convert our CAS string into the Binary Byte Data we need for the Reg Key
Function StringToRegBinary(strConvert)
	Dim strHex
	Dim arrHex
	strHex = ""
	For i=1 To Len(strConvert)
		strHex = strHex & Hex(Asc(Mid(strConvert,i,1)))
		if Not i = Len(strConvert) Then
			strHex = strHex & ",00,"
		End If
	Next
	strHex = strHex & ",00,00,00" ' need 3 bytes of data at the end of the CAS FQDN
	arrHex = Split(strHex, ",")
	StringToRegBinary = DecimalNumbers(arrHex)
End Function

' Converts Hex data to Binary Byte Data
Function DecimalNumbers(arrHex)
   Dim i, strDecValues
   For i = 0 to Ubound(arrHex)
     If isEmpty(strDecValues) Then
       strDecValues = CLng("&H" & arrHex(i))
     Else
       strDecValues = strDecValues & "," & CLng("&H" & arrHex(i))
     End If
   next
   DecimalNumbers = split(strDecValues, ",")
End Function

' Best way I could figure out to rebuild the DN with the name of the new CAS
Function FormatNewDN(curDN)
		Dim arrDN, strTemp, x, strNewDN
		' strip starting /
		strTemp = Right(curDN, Len(curDN) - 1)
		arrDN = Split(strTemp, "/")
		for x = lbound(arrDN) to ubound(arrDN)
			If arrDN(x) = "cn=Servers" Then
				arrDN(x+1) = "cn=" & TARGET_CAS_NETBIOS
				Exit For
			End If
		Next
		For x = lbound(arrDN) to ubound(arrDN)
			strNewDN = strNewDN & "/" & arrDN(x)
		Next
		FormatNewDN = strNewDN
End Function

' Exit the Program
Sub EndProgram(LogText, boolCleanExit)
	if boolCleanExit Then
		MsgBox LogText,64,"Change CAS Tool"
	Else
		MsgBox LogText,16,"Change CAS Tool"
	End If
	Wscript.Quit
End Sub