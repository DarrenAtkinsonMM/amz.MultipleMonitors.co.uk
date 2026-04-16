<%
Dim pcIntScUpgrade, pcIntBTO, pcIntAPP, pcIntCM, pcIntMS

Public Function pcf_GetSubVersions()

    Dim pcStrScSubVersion
    pcStrScSubVersion = ""
    
	'// CONFIGURATOR vs STD version
	pcIntBTO=statusBTO
	if len(pcIntBTO)<1 then
		pcIntBTO=0
	end if
	'// Add if missing
	if pcIntBTO=1 and InStr(pcStrScVersion,"b")=0 then
		pcStrScVersion=pcStrScVersion & "b"
	end if
	'// Remove if not needed
	if pcIntBTO=0 then
		pcStrScVersion=replace(pcStrScVersion, "b", "")
	end if

	'//Apparel Add-on status
	pcIntAPP=statusAPP
	if len(pcIntAPP)<1 then
		pcIntAPP=0
	end if
	'// Add if missing
	if pcIntAPP=1 and InStr(pcStrScSubVersion,"a")=0 then
		pcStrScSubVersion=pcStrScSubVersion & "a"
	end if
	'// Remove if not needed
	if pcIntAPP=0 and InStr(pcStrScSubVersion,"a")=1 then
		pcStrScSubVersion=replace(pcStrScSubVersion, "a", "")
	end if
	
	'//Conflict Management Add-on status
	pcIntCM=statusCM
	if len(pcIntCM)<1 then
		pcIntCM=0
	end if
    
	'// Add if missing
	if pcIntCM=1 and InStr(pcStrScSubVersion,"cm")=0 then
		pcStrScSubVersion=pcStrScSubVersion & "cm"
	end if
	'// Remove if not needed
	if pcIntCM=0 and InStr(pcStrScSubVersion,"cm")=1 then
		pcStrScSubVersion=replace(pcStrScSubVersion, "cm", "")
	end if
    
    '// Add if missing
    If pcIntScUpgrade = "" Then
        pcIntScUpgrade = 0
    End If
	
	'//Mobile Commerce Add-on status
	'// Ms = Mobile storefront
	pcIntMS=statusM
	if len(pcIntMS)<1 then
		pcIntMS=0
	end if
	pcStrScSubVersion=replace(pcStrScSubVersion, "Ms", "")
	
	'// Sub-version clean up (from older versions)
    pcStrScSubVersion=replace(pcStrScSubVersion, "p2", "")
    pcStrScSubVersion=replace(pcStrScSubVersion, "p", "")
    pcStrScSubVersion=replace(pcStrScSubVersion, "g131", "")
    pcStrScSubVersion=replace(pcStrScSubVersion, "g13", "")
    pcStrScSubVersion=replace(pcStrScSubVersion, "a171", "a")
    pcStrScSubVersion=replace(pcStrScSubVersion, "a17", "a")
    pcStrScSubVersion=replace(pcStrScSubVersion, "a3", "a")
    pcStrScSubVersion=replace(pcStrScSubVersion, "aa", "a")
    pcStrScSubVersion=replace(pcStrScSubVersion, "cmcm", "cm")
    pcStrScSubVersion=replace(pcStrScSubVersion, "MsMs", "Ms")

    '// TESTING
    'pcStrScSubVersion = "test"
    
    pcf_GetSubVersions = pcStrScSubVersion
    
End Function
%>