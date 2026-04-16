<%
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: Pinterest Pin It Button
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'// SETTINGS-START

'// Active: 1 = active, 0 = inactive
if scPinterestDisplay&""="" then
    pcInterest=0
else
    pcInterest=scPinterestDisplay
end if

'// Counter visibility- Options are: horizontal, vertical, none
if scPinterestCounter&""="" then
    pcPinItCounter="none"
else
    pcPinItCounter=scPinterestCounter
end if

if pcPinItCounter = "vertical" then
    pcPinItCounter = "above"
elseif pcPinItCounter = "horizontal" then
    pcPinItCounter = "beside"
end if
    
'// SETTINGS-END
	
    
    
'// PIN IT-START
If pcInterest=1 Then 
   
	tempURLp=replace((scStoreURL&"/"&scPcFolder&"/pc/"),"//","/")
	tempURLp=replace(tempURLp,"http:/","http://")
	tempURLp=replace(tempURLp,"https:/","https://")
	
	'// Product image
	pImageUrl1=""
	queryI="SELECT imageUrl,largeImageURL FROM Products WHERE idProduct=" & pIdProduct 
	set rsI=connTemp.execute(queryI)
	if not rsI.eof then
		pcv_strImageURL=rsI("imageUrl")
		pcv_strLargeImageURL=rsI("largeImageURL")
	end if
	set rsI=nothing

	If Not (IsNull(pcv_strLargeImageURL) OR pcv_strLargeImageURL="") then 
		pImageUrl1=pcv_strLargeImageURL
	ElseIf Not (IsNull(pcv_strImageURL) OR pcv_strImageURL="") then 
		pImageUrl1=pcv_strImageURL
	Else
		pcInterest=0 '// Hide Pinterest if no image
	End if
		
	pcPinterestDesc = pDescription
	pcPinterestDesc = replace(pcPinterestDesc, "<br>", vbCrLf)
	pcPinterestDesc = replace(pcPinterestDesc, "<BR>", vbCrLf)
	pcPinterestDesc = replace(pcPinterestDesc, "<br/>", vbCrLf)
	pcPinterestDesc = replace(pcPinterestDesc, "<BR/>", vbCrLf)

	If pcInterest=1 Then
	%>
    <!-- Pinterest -->
		<span class="pcPinterest hidden-xs">
			<%
				pinterestUrl = "//www.pinterest.com/pin/create/button/?url=" & Server.Urlencode(tempURLp&pcStrPrdLink) & "&media=" & Server.Urlencode(tempURLp&pcf_getImagePath(pcv_tmpNewPath & "catalog",pImageUrl1)) & "&description=" & Server.UrlEncode(pcPinterestDesc)
			%>
			<a href="<%= Server.HtmlEncode(pinterestUrl) %>" class="pin-it-button" data-pin-do="buttonPin" data-pin-config="<%=pcPinItCounter%>" target="_blank"><img style="border: 0px" src="<%=pcf_getImagePath("//assets.pinterest.com/images","PinExt.png")%>" alt="Pin It" title="Pin It" /></a>
		</span>
		
		<% if pcPinterestLoaded <> true then %>
			<script type="text/javascript" async src="<%=pcf_getJSPath("//assets.pinterest.com/js","pinit.js")%>"></script>
			<% pcPinterestLoaded = true %>
		<% end if %>
	<%
	End If
    
End If 
'// PIN IT-END

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: Pinterest Pin It Button
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
%>