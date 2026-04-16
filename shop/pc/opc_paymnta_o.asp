<%
'This file is part of ProductCart, an ecommerce application developed And sold by NetSource Commerce. ProductCart, its source code, the ProductCart name And logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute And/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/validation.asp" -->
<!--#include file="opc_contentType.asp" -->
<%
Set conlayout=Server.CreateObject("ADODB.Connection")
conlayout.Open scDSN
Set RSlayout = conlayout.Execute("Select * From layout Where layout.ID=2")
Set rsIconObj = conlayout.Execute("Select * From icons WHERE id=1")

Call SetContentType()

Call pcs_CheckLoggedIn()

Dim pcTempIdPayment
pcTempIdPayment = getUserInput(request("idPayment"),0)

If session("GWPaymentId") = "" Then
	session("GWPaymentId") = getUserInput(pcTempIdPayment, 0)
Else
	If pcTempIdPayment <> session("GWPaymentId") And pcTempIdPayment<>"" Then
		session("GWPaymentId") = getUserInput(pcTempIdPayment, 0)
	End If
End If

If request("PaymentGWSubmitted")="Go" Then
	
	pCardType = getUserInput(request("cardType"), 0)
	pCardNumber = getUserInput(request("cardNumber"), 0)
	session("admin-" & session("GWPaymentId") & "-pCardType") = pCardType
	session("admin-" & session("GWPaymentId") & "-pCardNumber") = pCardNumber
	session("admin-" & session("GWPaymentId") & "-expMonth") = getUserInput(request("expMonth"), 0)
	session("admin-" & session("GWPaymentId") & "-expYear") = getUserInput(request("expYear"), 0)

    '// Validate Fields
    pcErrMsg=""
    
	If request("expMonth")<>"" And request("expYear")<>"" Then
		pExpiration=getUserInput(request("expMonth"),0) & "/1/" & getUserInput(request("expYear"),0)
	Else
		pcErrMsg= pcErrMsg & "<li>You did not enter Card Expiration Date</li>"
	End If

	If pCardType="" Then
		pcErrMsg= pcErrMsg & "<li>You did not select the Card Type</li>"
	End If

	If pCardNumber="" Then
		pcErrMsg= pcErrMsg & "<li>You did not enter Card Number</li>"
	End If

	If pExpiration="" Then
		pcErrMsg= pcErrMsg & "<li>You did not enter Card Expiration Date</li>"
	End If

	If pcErrMsg="" Then
		'// Validate Expiration
		If not IsCreditCard(pCardNumber, pCardType) Then
			pcErrMsg=pcErrMsg & "<li>" & dictLanguage.Item(Session("language")&"_paymntb_o_5") & "</li>"
		Else
			If DateDiff("d", Month(Now)&"/"&Year(now), request("expMonth")&"/"&request("expYear"))<=-1 Then
				pcErrMsg= pcErrMsg & "<li>" & dictLanguage.Item(Session("language")&"_paymntb_o_6") & "</li>"
			End If
		End If
	End If

	If pcErrMsg = "" Then

		pcv_SecurityPass = pcs_GetSecureKey
		pcv_SecurityKeyID = pcs_GetKeyID

		'// Encrypt CC data
		pCardNumber2=enDeCrypt(pCardNumber, pcv_SecurityPass)
		
		'// Extract real idorder (without prefix)
		if session("GWOrderId")="" then
			session("GWOrderId")="0"
		end if
		pTrueOrderId=cLng(session("GWOrderId"))-cLng(scpre)
		
		If (IsNull(session("GWOrderId"))) OR (session("GWOrderId")="") OR (session("GWOrderId")="0") OR (Clng(pTrueOrderId)<0) then
			pcErrMsg=dictLanguage.Item(Session("language")&"_paymntb_o_9")
		Else
			'// Save credit card
			query="INSERT INTO creditcards (idorder, cardType, cardNumber, expiration, seqcode, pcSecurityKeyID) VALUES (" &pTrueOrderId& ",'" &pCardType& "','" &pCardNumber2& "','" &pExpiration& "','na', "&pcv_SecurityKeyID&")"
			Set rs=server.CreateObject("ADODB.RecordSet")
			Set rs=conntemp.execute(query)
			Set rs=nothing
			
			if err.number<>0 then
				pcErrMsg=dictLanguage.Item(Session("language")&"_paymntb_o_9")
				err.number=0
				err.description=""
			end if
		End if
	End If

	If pcErrMsg="" Then
		Response.write "OK"
		session("Entered-" & session("GWPaymentId"))="1"
	Else
		session("Entered-" & session("GWPaymentId"))=""
		pcErrMsg="Errors when saving payment details:<ul>"&pcErrMsg&"</ul>"
		Response.write pcErrMsg
	End If


Else '// If request("PaymentGWSubmitted")="Go" Then

    %>
    <script type=text/javascript>NeedToUpdatePay=1;</script>

    <h3 class="pcSectionTitle"><%response.write dictLanguage.Item(Session("language")&"_GateWay_5")%></h3>

    <% If len(pcCustIpAddress)>0 And CustomerIPAlert="1" Then %>

        <div class="pcErrorMessage"><%=dictLanguage.Item(Session("language")&"_GateWay_6")&pcCustIpAddress%></div>

    <% End If %>

    <div class="form-group">
        <label for="cardType"><%=dictLanguage.Item(Session("language")&"_paymnta_o_2") %></label>        
        <%					
        query="SELECT CCcode,CCType FROM CCTypes WHERE active=-1;"
        Set rs=server.createobject("adodb.recordset")
        Set rs=connTemp.execute(query)
        %>
        <select name="cardType" id="cardType" onchange="javascript:var tmpval=$pc('#cardNumber').val();$pc('#cardNumber').val(tmpval+' ');$pc('#PayForm').validate().element('#cardNumber');$pc('#cardNumber').val(tmpval);$pc('#cardNumber').focus();$pc('#cardType').focus();" class="form-control required">
        <% 
        Do Until rs.Eof
            CCcode=rs("CCcode")
            CCType=rs("CCType")  
            %>
            <option value="<%=CCcode%>" <%If session("Entered-" & session("GWPaymentId"))="1" Then%><%If CCcode=session("admin-" & session("GWPaymentId") & "-pCardType") Then%>selected<%End If%><%End If%>><%=CCType%></option>
            <% 
            rs.MoveNext
        Loop
        Set rs=nothing
        %>
        </select>
    </div>


    <div class="form-group">
        <label for="cardNumber"><%=dictLanguage.Item(Session("language")&"_paymnta_o_3") %></label>
        <input type="text" class="form-control" name="cardNumber" id="cardNumber" size="30" <%If session("Entered-" & session("GWPaymentId"))="1" Then%>value="<%=session("admin-" & session("GWPaymentId") & "-pCardNumber")%>"<%End If%>>
    </div>


    <div class="row">
        <div class="col-sm-6">
            <label for="expMonth"><%=dictLanguage.Item(Session("language")&"_paymnta_o_4") %>&nbsp;<%=dictLanguage.Item(Session("language")&"_paymnta_o_5") %></label>
            <select class="form-control" name="expMonth">
                <option value="1" <%If session("Entered-" & session("GWPaymentId"))="1" Then%><%If session("admin-" & session("GWPaymentId") & "-expMonth")="1" Then%>selected<%End If%><%End If%>>1</option>
                <option value="2" <%If session("Entered-" & session("GWPaymentId"))="1" Then%><%If session("admin-" & session("GWPaymentId") & "-expMonth")="2" Then%>selected<%End If%><%End If%>>2</option>
                <option value="3" <%If session("Entered-" & session("GWPaymentId"))="1" Then%><%If session("admin-" & session("GWPaymentId") & "-expMonth")="3" Then%>selected<%End If%><%End If%>>3</option>
                <option value="4" <%If session("Entered-" & session("GWPaymentId"))="1" Then%><%If session("admin-" & session("GWPaymentId") & "-expMonth")="4" Then%>selected<%End If%><%End If%>>4</option>
                <option value="5" <%If session("Entered-" & session("GWPaymentId"))="1" Then%><%If session("admin-" & session("GWPaymentId") & "-expMonth")="5" Then%>selected<%End If%><%End If%>>5</option>
                <option value="6" <%If session("Entered-" & session("GWPaymentId"))="1" Then%><%If session("admin-" & session("GWPaymentId") & "-expMonth")="6" Then%>selected<%End If%><%End If%>>6</option>
                <option value="7" <%If session("Entered-" & session("GWPaymentId"))="1" Then%><%If session("admin-" & session("GWPaymentId") & "-expMonth")="7" Then%>selected<%End If%><%End If%>>7</option>
                <option value="8" <%If session("Entered-" & session("GWPaymentId"))="1" Then%><%If session("admin-" & session("GWPaymentId") & "-expMonth")="8" Then%>selected<%End If%><%End If%>>8</option>
                <option value="9" <%If session("Entered-" & session("GWPaymentId"))="1" Then%><%If session("admin-" & session("GWPaymentId") & "-expMonth")="9" Then%>selected<%End If%><%End If%>>9</option>
                <option value="10" <%If session("Entered-" & session("GWPaymentId"))="1" Then%><%If session("admin-" & session("GWPaymentId") & "-expMonth")="10" Then%>selected<%End If%><%End If%>>10</option>
                <option value="11" <%If session("Entered-" & session("GWPaymentId"))="1" Then%><%If session("admin-" & session("GWPaymentId") & "-expMonth")="11" Then%>selected<%End If%><%End If%>>11</option>
                <option value="12" <%If session("Entered-" & session("GWPaymentId"))="1" Then%><%If session("admin-" & session("GWPaymentId") & "-expMonth")="12" Then%>selected<%End If%><%End If%>>12</option>
            </select>
        </div>
        <div class="col-sm-6">
            <label for="ExpYear"><%=dictLanguage.Item(Session("language")&"_paymnta_o_4") %>&nbsp;<%=dictLanguage.Item(Session("language")&"_paymnta_o_6") %></label>
            <select class="form-control" name="ExpYear">
                <% Dim varYear
                varYear=year(now) %>
                <option value="<%=varYear%>" <%If session("Entered-" & session("GWPaymentId"))="1" Then%><%If session("admin-" & session("GWPaymentId") & "-expYear")&""=clng(varYear) & "" Then%>selected<%End If%><%End If%>><%=varYear%></option>
                <option value="<%=varYear+1%>" <%If session("Entered-" & session("GWPaymentId"))="1" Then%><%If session("admin-" & session("GWPaymentId") & "-expYear")&""=clng(varYear+1) & "" Then%>selected<%End If%><%End If%>><%=varYear+1%></option>
                <option value="<%=varYear+2%>" <%If session("Entered-" & session("GWPaymentId"))="1" Then%><%If session("admin-" & session("GWPaymentId") & "-expYear")&""=clng(varYear+2) & "" Then%>selected<%End If%><%End If%>><%=varYear+2%></option>
                <option value="<%=varYear+3%>" <%If session("Entered-" & session("GWPaymentId"))="1" Then%><%If session("admin-" & session("GWPaymentId") & "-expYear")&""=clng(varYear+3) & "" Then%>selected<%End If%><%End If%>><%=varYear+3%></option>
                <option value="<%=varYear+4%>" <%If session("Entered-" & session("GWPaymentId"))="1" Then%><%If session("admin-" & session("GWPaymentId") & "-expYear")&""=clng(varYear+4) & "" Then%>selected<%End If%><%End If%>><%=varYear+4%></option>
                <option value="<%=varYear+5%>" <%If session("Entered-" & session("GWPaymentId"))="1" Then%><%If session("admin-" & session("GWPaymentId") & "-expYear")&""=clng(varYear+5) & "" Then%>selected<%End If%><%End If%>><%=varYear+5%></option>
                <option value="<%=varYear+6%>" <%If session("Entered-" & session("GWPaymentId"))="1" Then%><%If session("admin-" & session("GWPaymentId") & "-expYear")&""=clng(varYear+6) & "" Then%>selected<%End If%><%End If%>><%=varYear+6%></option>
                <option value="<%=varYear+7%>" <%If session("Entered-" & session("GWPaymentId"))="1" Then%><%If session("admin-" & session("GWPaymentId") & "-expYear")&""=clng(varYear+7) & "" Then%>selected<%End If%><%End If%>><%=varYear+7%></option>
                <option value="<%=varYear+8%>" <%If session("Entered-" & session("GWPaymentId"))="1" Then%><%If session("admin-" & session("GWPaymentId") & "-expYear")&""=clng(varYear+8) & "" Then%>selected<%End If%><%End If%>><%=varYear+8%></option>
                <option value="<%=varYear+9%>" <%If session("Entered-" & session("GWPaymentId"))="1" Then%><%If session("admin-" & session("GWPaymentId") & "-expYear")&""=clng(varYear+9) & "" Then%>selected<%End If%><%End If%>><%=varYear+9%></option>
            </select>
        </div>
    </div> 

    <input type="image" name="PaySubmit" id="PaySubmit" alt="Submit" src="<%=pcf_getImagePath("",RSlayout("pcLO_Update"))%>" style="display:none">
    <script type=text/javascript>
        //*Submit Pay Form
        $pc('#PaySubmit').click(function() {          
            if ($pc('#PayForm').validate().form()) {            
                {
                    $pc.ajax({
                        type: "POST",
                        url: "opc_paymnta_o.asp",
                        data: $pc('#PayForm').formSerialize() + "&PaymentGWSubmitted=Go",
                        timeout: 450000,
                        success: function(data, textStatus) 
                        {
                            if (data=="SECURITY") {
                                // Session Expired
                                window.location="msg.asp?message=1";
                            } else {
                                
                                if (data=="OK") {
                                    $pc("#PayLoader").hide();
                                    NeedToUpdatePay=0;
                                    recalculate("","#PayLoader1",0,'');
                                    ValidateGroup2();
                                } else {
                                    $pc("#PayLoader").html('<div class="pcErrorMessage">' + data + '</div>');
                                    $pc("#PayLoader").show();
                                    NeedToUpdatePay = 1;
                                }
                                
                            }
                        }
                    });                    
                    return(false);
                }                
            }
            return(false);
        });
        <% If session("Entered-" & session("GWPaymentId"))="1" Then %>
            
            NeedToUpdatePay=0;
            
        <% End If %>
    </script>


<% End If %>

<% function IsCreditCard(ByRef anCardNumber, ByRef asCardType)
	Dim lsNumber		' Credit card number stripped of all spaces, dashes, etc.
	Dim lsChar			' an individual character
	Dim lnTotal			' Sum of all calculations
	Dim lnDigit			' A digit found within a credit card number
	Dim lnPosition		' identifies a character position In a String
	Dim lnSum			' Sum of calculations For a specific Set

	' Default result is False
	IsCreditCard = False

	' ====
	' Strip all characters that are Not numbers.
	' ====

	' Loop through Each character inthe card number submited
	For lnPosition = 1 To Len(anCardNumber)
		' Grab the current character
		lsChar = Mid(anCardNumber, lnPosition, 1)
		' If the character is a number, append it To our new number
		If validNum(lsChar) Then lsNumber = lsNumber & lsChar

	Next ' lnPosition

	' ====
	' The credit card number must be between 13 And 16 digits.
	' ====
	' If the length of the number is less Then 13 digits, Then Exit the routine
	If Len(lsNumber) < 13 Then Exit function

	' If the length of the number is more Then 16 digits, Then Exit the routine
	If Len(lsNumber) > 16 Then Exit function

	' Choose action based on Type of card
	Select Case LCase(asCardType)
		' VISA
		Case "visa", "v", "V"
			' If first digit Not 4, Exit function
			If Not Left(lsNumber, 1) = "4" Then Exit function
		' American Express
		Case "american express", "americanexpress", "american", "ax", "a", "A"
			' If first 2 digits Not 37, Exit function
			If Not Left(lsNumber, 2) = "37" And Not Left(lsNumber, 2) = "34" Then Exit function
		' Mastercard
		Case "mastercard", "master card", "master", "m", "M"
			' If first digit Not 5, Exit function
			If Not Left(lsNumber, 1) = "5" Then Exit function
		' Discover
		Case "discover", "discovercard", "discover card", "d", "D"
			' If first digit Not 6, Exit function
			If Not Left(lsNumber, 1) = "6" Then Exit function
		'Diners Card	
		Case "dc", "DC"
			checkDC=0
			Select Case Left(lsNumber, 2)
				Case "54","55","36","38","39":
				Case Else: checkDC=1
			End Select
			If checkDC=1 then
				Select Case Left(lsNumber, 3)
					Case "300","301","302","303","304","305","309": checkDC=0
				End Select
			End if
			If checkDC=1 then Exit function
		Case Else
	End Select ' LCase(asCardType)

	' ====
	' If the credit card number is less Then 16 digits add zeros
	' To the beginning to make it 16 digits.
	' ====
	' Continue Loop While the length of the number is less Then 16 digits
	While Not Len(lsNumber) = 16

		' Insert 0 To the beginning of the number
		lsNumber = "0" & lsNumber

	Wend ' Not Len(lsNumber) = 16

	' ====
	' Multiply Each digit of the credit card number by the corresponding digit of
	' the mask, And sum the results together.
	' ====

	' Loop through Each digit
	For lnPosition = 1 To 16

		' Parse a digit from a specified position In the number
		lnDigit = Mid(lsNumber, lnPosition, 1)

		' Determine If we multiply by:
		'	1 (Even)
		'	2 (Odd)
		' based On the position that we are reading the digit from
		lnMultiplier = 1 + (lnPosition Mod 2)

		' Calculate the sum by multiplying the digit And the Multiplier
		lnSum = lnDigit * lnMultiplier

		' (Single digits roll over To remain single. We manually have to Do this.)
		' If the Sum is 10 or more, subtract 9
		If lnSum > 9 Then lnSum = lnSum - 9

		' Add the sum To the total of all sums
		lnTotal = lnTotal + lnSum

	Next ' lnPosition

	' ====
	' Once all the results are summed divide
	' by 10, If there is no remainder Then the credit card number is valid.
	' ====
	IsCreditCard = ((lnTotal Mod 10) = 0)

End function ' IsCreditCard
' ------------------------------------------------------------------------------

call closeDb()
conlayout.Close
Set conlayout=nothing
Set RSlayout = nothing
Set rsIconObj = nothing
%>
