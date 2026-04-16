<%
prcs_SiteKey=""
prcs_Secret=""
prcs_Theme="light"
prcs_Type="image"
prcs_Size="normal"
prcs_Had=0

Public Function pcf_checkScriptName()

    if (scUseImgs=1) AND (scAffLogin=1) AND (InStr(LCase(Request.ServerVariables("SCRIPT_NAME")), "affiliatelogin.asp")>0) then
        pcf_checkScriptName=1
        exit function
    end if
    if (scUseImgs=1) AND (scAffReg=1) AND (InStr(LCase(Request.ServerVariables("SCRIPT_NAME")), "newaffa.asp")>0) then
        pcf_checkScriptName=1
        exit function
    end if
    if (scUseImgs=1) AND (InStr(LCase(Request.ServerVariables("SCRIPT_NAME")), "checkout.asp")>0) then
        pcf_checkScriptName=1
        exit function
    end if
    if (scUseImgs=1) AND (InStr(LCase(Request.ServerVariables("SCRIPT_NAME")), "onepagecheckout.asp")>0) then
        pcf_checkScriptName=1
        exit function
    end if
    if (scReview=1) AND (InStr(LCase(Request.ServerVariables("SCRIPT_NAME")), "prv_postreview.asp")>0) then
        pcf_checkScriptName=1
        exit function
    end if
    if (scContact=1) AND (InStr(LCase(Request.ServerVariables("SCRIPT_NAME")), "contact.asp")>0) then
        pcf_checkScriptName=1
        exit function
    end if
    
    if (scAdminLogin=1) AND (scUseImgs2=1) AND (InStr(LCase(Request.ServerVariables("SCRIPT_NAME")), "login_1.asp")>0) then
        pcf_checkScriptName=1
        exit function
    end if

    if (InStr(LCase(Request.ServerVariables("SCRIPT_NAME")), "recaptchasettings.asp")>0) then
        pcf_checkScriptName=1
        exit function
    end if

    pcf_checkScriptName=0

End Function


Public Sub pcs_getReCaSettings(checkS)
    Dim rs, query,tmpR
    if checkS=1 then
        tmpR=pcf_checkScriptName()
    else
        tmpR=1
    end if  
    IF (scSecurity=1) AND (scCaptchaType="1") AND (tmpR=1) THEN
        query="SELECT pcRCS_SiteKey,pcRCS_Secret,pcRCS_Theme,pcRCS_Type,pcRCS_Size FROM pcReCaSettings;"
        set rs=connTemp.execute(query)
        if not rs.eof then
            prcs_SiteKey=rs("pcRCS_SiteKey")
            prcs_Secret=rs("pcRCS_Secret")
            prcs_Theme=rs("pcRCS_Theme")
            prcs_Type=rs("pcRCS_Type")
            prcs_Size=rs("pcRCS_Size")
            if prcs_SiteKey<>"" then
                prcs_SiteKey=enDeCrypt(prcs_SiteKey, scCrypPass)
            end if
            if prcs_Secret<>"" then
                prcs_Secret=enDeCrypt(prcs_Secret, scCrypPass)
            end if
        end if
        set rs=nothing
        if (prcs_SiteKey<>"") AND (prcs_Secret<>"") then
            prcs_Had=1
        end if
    END IF
End Sub


Public Sub pcs_getReCaSettingsNoAuth(checkS)
    Dim rs, query,tmpR
    if checkS=1 then
        tmpR=pcf_checkScriptName()
    else
        tmpR=1
    end if    
    IF (tmpR=1) THEN
        query="SELECT pcRCS_SiteKey,pcRCS_Secret,pcRCS_Theme,pcRCS_Type,pcRCS_Size FROM pcReCaSettings;"
        set rs=connTemp.execute(query)
        if not rs.eof then
            prcs_SiteKey=rs("pcRCS_SiteKey")
            prcs_Secret=rs("pcRCS_Secret")
            prcs_Theme=rs("pcRCS_Theme")
            prcs_Type=rs("pcRCS_Type")
            prcs_Size=rs("pcRCS_Size")
            if prcs_SiteKey<>"" then
                prcs_SiteKey=enDeCrypt(prcs_SiteKey, scCrypPass)
            end if
            if prcs_Secret<>"" then
                prcs_Secret=enDeCrypt(prcs_Secret, scCrypPass)
            end if
        end if
        set rs=nothing
        if (prcs_SiteKey<>"") AND (prcs_Secret<>"") then
            prcs_Had=1
        end if
    END IF
End Sub


Public Sub pcs_genReCaHeader()
    call pcs_getReCaSettings(1)
    if prcs_Had="1" then
        %>
        <script type="text/javascript">
		var widgetId=0;
        var onloadCallback = function() {
            widgetId = grecaptcha.render('gcaptcha', {
            'sitekey' : '<%=prcs_SiteKey%>',
            'theme' : '<%=prcs_Theme%>',
            'type' : '<%=prcs_Type%>',
            'size' : '<%=prcs_Size%>'
            });
        };
		
        </script>
        <script src="https://www.google.com/recaptcha/api.js?onload=onloadCallback&render=explicit" async defer></script>
        <%
    end if
End Sub


Public Sub pcs_genReCaptcha()
    if prcs_Had="1" then
        %>
        <div id="gcaptcha"></div>
        <%
    end if
End Sub


Public Function pcf_checkReCaptcha()
    Dim prcs_gResponse
    Dim VarString,objXmlHttp,ResponseString
	prcs_gResponse=request("g-recaptcha-response")
	if prcs_gResponse="" then
		pcf_checkReCaptcha=false
		exit function
	end if	
	call pcs_getReCaSettings(0)	
	if prcs_Had="1" then		
		VarString = "secret=" & prcs_Secret & "&response=" & prcs_gResponse		
		Set objXmlHttp = Server.CreateObject("Msxml2.ServerXMLHTTP" & scXML)
		objXmlHttp.open "POST", "https://www.google.com/recaptcha/api/siteverify", False
		objXmlHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXmlHttp.send VarString
		pcf_checkReCaptcha=false
		ResponseString = objXmlHttp.responseText
		if Instr(Lcase(ResponseString),"""success"": true")>0 then
			pcf_checkReCaptcha=true
			exit function
		else
			pcf_checkReCaptcha=false
			exit function
		end if
	else
		pcf_checkReCaptcha=false
		exit function
	end if
End Function

%>
