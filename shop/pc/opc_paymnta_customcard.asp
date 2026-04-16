<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
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

pcGatewayDataIdOrder = Session("GWOrderID")

If Request("PaymentGWSubmitted") = "Go" Then

	pcIntIdCustomerCardType = getUserInput(Request("idCCT"), 20)

	'// Extract real idorder (without prefix)
	pTrueOrderId = (int(session("GWOrderId"))-scpre)

	query="SELECT customCardRules.idCustomCardRules, customCardRules.idCustomCardType, customCardRules.ruleName, customCardRules.intruleRequired, customCardRules.intlengthOfField, customCardRules.intmaxInput FROM customCardRules WHERE (((customCardRules.idCustomCardType)=" & pcIntIdCustomerCardType & ")) ORDER BY customCardRules.intOrder;"
	Set rs = server.CreateObject("ADODB.RecordSet")
	Set rs = conntemp.execute(query)
	pcErrMsg = ""
	Do Until rs.Eof
		pIdCCR = rs("idCustomCardRules")
		pRuleName = rs("ruleName")
		pReq = rs("intruleRequired")
		session("admin"&pIdCCR) = URLDecode(getUserInput(request("customfield"&pIdCCR),0))
		session("admin-" & session("GWPaymentId") & "-" &pIdCCR) = URLDecode(getUserInput(request("customfield"&pIdCCR),0))
		If pReq<>"0" Then
			If session("admin"&pIdCCR)="" Then
				pcErrMsg=pcErrMsg & "<li>" & pRuleName & " is a required field.</li>"
			End If
		End If
		rs.MoveNext
	Loop
	Set rs = Nothing

	If pcErrMsg="" Then
	
	    '// Save custom info
	    query="SELECT customCardRules.idCustomCardRules, customCardRules.idCustomCardType, customCardRules.ruleName, customCardRules.intruleRequired, customCardRules.intlengthOfField, customCardRules.intmaxInput FROM customCardRules WHERE (((customCardRules.idCustomCardType)="&pcIntIdCustomerCardType&")) ORDER BY customCardRules.intOrder;"
	    Set rs = server.CreateObject("ADODB.RecordSet")
	    Set rs = conntemp.execute(query)

	    Do Until rs.Eof
		
		    pcv_strCustomCardRules=rs("idCustomCardRules")
		    pcv_strRuleName=rs("ruleName")
		    pIdCCR=replace(pcv_strCustomCardRules,"'","''")
		
		    '// Create an amendum to the admin order email
		    If len(session("admin"&pIdCCR))>0 Then
			    '// Check if this is a credit/debit card number
			    strRuleValue=ShowLastFour(session("admin"&pIdCCR))
			    ammendAdminEmail = ammendAdminEmail & pcv_strRuleName & ": " & strRuleValue & vbCrLf
			    strRuleValue=""
		    End If
		
		    pRuleName=replace(pcv_strRuleName,"'","''")
		    pReq=rs("intruleRequired")
		
		    pcBillingTotal=0
		
		    '// Extract real idorder (without prefix)
		    pTrueOrderId = (int(session("GWOrderId"))-scpre)
		
		    query="INSERT INTO customcardOrders (idorder, idcustomCardType, idcustomCardRules, strFormValue, intOrderTotal,strRuleName) VALUES (" &pTrueOrderId& "," &pcIntIdCustomerCardType& "," &pIdCCR& ",'" &replace(session("admin"&pIdCCR),"'","''")& "'," &pcBillingTotal& ",N'"&pRuleName&"')"
		    set rsCCObj=server.CreateObject("ADODB.RecordSet")
		    set rsCCObj=conntemp.execute(query)
		    set rsCCObj=nothing
		
		    rs.MoveNext
	    Loop
	    Set rs = Nothing

	    '// Save ammendum
	    Session("pcSFSpecialFields")=ammendAdminEmail
        
	End If '// If pcErrMsg="" Then
	
	If pcErrMsg="" Then
		Response.write "OK"
		session("Entered-" & session("GWPaymentId"))="1"
	Else
		session("Entered-" & session("GWPaymentId"))=""
		pcErrMsg="Errors when saving payment details:<ul>"&pcErrMsg&"</ul>"
		Response.write pcErrMsg
	End If


Else '// If Request("PaymentGWSubmitted") = "Go" Then


	query="SELECT payTypes.paymentDesc, customCardTypes.idcustomCardType FROM payTypes INNER JOIN customCardTypes ON payTypes.paymentDesc = customCardTypes.customCardDesc WHERE (((payTypes.idPayment)="&session("GWPaymentId")&"));"
	Set rs = Server.CreateObject("ADODB.Recordset")
	Set rs = conntemp.execute(query)
    If Not rs.Eof Then		
	    pcStrPaymentDesc = rs("paymentDesc")
	    pcIntIdCustomerCardType = rs("idcustomCardType")
	End If
	Set rs = Nothing
	%>

    <input type="hidden" name="idCCT" value="<%=pcIntIdCustomerCardType%>">

    <% If session("Entered-" & session("GWPaymentId"))<>"1" Then %>
		
        <script type=text/javascript>NeedToUpdatePay=1;</script>
        
    <% End If %>

        <div class="daOPCPayBankMsg">Once you have placed the order we will contact you with our bank details so that you may complete the payment.</div>
        
		<% If len(pcCustIpAddress)>0 And CustomerIPAlert="1" Then %>

            <div class="pcErrorMessage"><%=dictLanguage.Item(Session("language")&"_GateWay_6")&pcCustIpAddress%></div>

        <% End If %>
        
        <% 
        query="SELECT idCustomCardRules, idCustomCardType, ruleName, intruleRequired, intlengthOfField, intmaxInput FROM customCardRules WHERE (((idCustomCardType)=" & pcIntIdCustomerCardType & ")) ORDER BY intOrder;"
        Set rs=Server.CreateObject("ADODB.Recordset")
        Set rs=conntemp.execute(query)
        HaveFields=0
        Do Until rs.Eof
        
            HaveFields=1
            
            pIdCCR=rs("idCustomCardRules")
            pcIntIdCustomerCardType=rs("idCustomCardType")
            pRuleName=rs("ruleName")
            pReq=rs("intruleRequired")
            pLOF=rs("intlengthOfField")
            pMInput=rs("intmaxInput")
            
            If pMInput = "" Or pMInput = "0" Then
                pMInput = pLOF
            End If
            %>
            <div class="form-group">
                <label for="customField<%=pIdCCR%>"><%=pRuleName%></label>
                <input type="text" class="form-control <% If pReq<>"0" Then %>required<% End If %>" id="customField<%=pIdCCR%>" name="customField<%=pIdCCR%>" <% If session("Entered-" & session("GWPaymentId"))="1" Then %>value="<%=session("admin-" & session("GWPaymentId") & "-" &pIdCCR)%>"<% End If %> size="<%=pLOF%>" maxlength="<%=pMInput%>">
            </div>
            <% 
            rs.Movenext
        Loop
        Set rs = Nothing
        %>
		
        <% If HaveFields = 1 Then %>


            <input type="image" name="PaySubmit" id="PaySubmit" src="<%=pcf_getImagePath("",RSlayout("pcLO_Update"))%>" border="0" style="display:none">
            <script type=text/javascript>
                //*Submit Pay Form
                $pc('#PaySubmit').click(function(){
                if ($pc('#PayForm').validate().form())
                {
                {
                    $pc.ajax({
                        type: "POST",
                        url: "opc_paymnta_customcard.asp",
                        data: $pc('#PayForm').formSerialize() + "&PaymentGWSubmitted=Go",
                        timeout: 450000,
                        success: function(data, textStatus){
                        if (data=="SECURITY")
                        {
                            // Session Expired
                            window.location="msg.asp?message=1";
                        }
                        else
                        {
                            if (data=="OK")
                            {
                                $pc("#PayLoader").hide();
                                NeedToUpdatePay=0;
                                recalculate("","#PayLoader1",0,'');
                                ValidateGroup2();
                            }
                            else
                            {
                                $pc("#PayLoader").html('<img src="<%=pcf_getImagePath("images","pc_icon_error_small.png")%>" align="absmiddle"> '+data);
                                $pc("#PayLoader").show();
                                NeedToUpdatePay=1;
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

<%
End If  '// If Request("PaymentGWSubmitted") = "Go" Then

call closeDb()
conlayout.Close
Set conlayout=nothing
Set RSlayout = nothing
Set rsIconObj = nothing
%>
