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

If Request("PaymentGWSubmitted") = "Go" Then

	'// Extract real idorder (without prefix)
	pTrueOrderId = (int(session("GWOrderId"))-scpre)

    '// The following value is only used in "adminNewOrderEmail.asp"
    pAccNum = URLDecode(getUserInput(Request("AccNum"), 0))
	Session("pcSFpAccNum2") = pAccNum
    
    '// The following value is only used for re-filling the form field on error.
	session("AccNum-" & session("GWPaymentId")) = pAccNum

	'// Save account info	
	query="INSERT INTO offlinepayments (idorder, idPayment, AccNum) VALUES (" & pTrueOrderId & "," & session("GWPaymentId") & ",'" & pAccNum & "')"
	Set rs = server.CreateObject("ADODB.RecordSet")
	Set rs = conntemp.execute(query)
	Set rs = Nothing

	Response.write "OK"
	session("Entered-" & session("GWPaymentId"))="1"
    
Else '//If Request("PaymentGWSubmitted") = "Go" Then

	pcGatewayDataIdOrder = session("GWOrderID")

	query = "SELECT paymentDesc, idPayment, terms, CReq, CPrompt FROM PayTypes WHERE idPayment=" & session("GWPaymentId")
	Set rs = Server.CreateObject("ADODB.Recordset")
	Set rs = conntemp.execute(query)
    If Not rs.Eof Then	
        pcStrPaymentDesc = rs("paymentDesc")
        pcStrTerms = rs("terms")
        pcCReq = rs("CReq")
        pcCPrompt = rs("CPrompt")
    End If
	Set rs=nothing
	%>

  	<h3 class="pcSectionTitle"><%=dictLanguage.Item(Session("language")&"_GateWay_5")%></h3>

    <% If len(pcCustIpAddress)>0 And CustomerIPAlert="1" Then %>
        <div class="pcErrorMessage"><%=dictLanguage.Item(Session("language")&"_GateWay_6") & pcCustIpAddress %></div>
    <% End If %>


    <% 'Type of Payment %>
    <div class="form-group">
        <label><%= dictLanguage.Item(Session("language")&"_paymnta_c_2")%></label>
        <span class="help-text"><%=pcStrPaymentDesc%></span>
    </div>
  
  
    <% 'Terms %>
    <div class="form-group">
        <label><%= dictLanguage.Item(Session("language")&"_paymnta_c_3")%></label>
        <span class="help-text"><%=pcStrTerms%></span>
    </div>


        <% If pcCReq=-1 Then %>

            <div class="form-group">
                <label for="AccNum"><%=pcCPrompt%></label>
                <input type="text" name="AccNum" class="form-control required" <%if session("Entered-" & session("GWPaymentId"))="1" then%>value="<%=session("AccNum-" & session("GWPaymentId"))%>"<%end if%>>
            </div>

            <input type="image" name="PaySubmit" id="PaySubmit" alt="Submit" src="<%=pcf_getImagePath("",RSlayout("pcLO_Update"))%>" style="display:none">
            
            <script type=text/javascript>
                //*Submit Pay Form
                    $pc('#PaySubmit').click(function() {
                    if ($pc('#PayForm').validate().form())
                    {
                    {
                        $pc.ajax({
                            type: "POST",
                            url: "opc_paymnta_c.asp",
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
                    <% if session("Entered-" & session("GWPaymentId"))="1" then %>						
                        NeedToUpdatePay=0;
                    <% end if %>
            </script>


        <% Else '// If pcCReq=-1 Then %>


            <input type="image" name="PaySubmit" id="PaySubmit" alt="Submit" src="<%=pcf_getImagePath("",RSlayout("pcLO_Update"))%>" style="display:none">
      
            <script type=text/javascript>
                //*Submit Pay Form
                $pc('#PaySubmit').click(function(){	
                           
                    $pc("#PayLoader").hide();
                    NeedToUpdatePay=0;
                    recalculate("","#PayLoader1",0,'');
                    ValidateGroup2();
                    return(false);
      
                });
            </script>


        <% End If '// If pcCReq=-1 Then %>


<% 
End If '// If Request("PaymentGWSubmitted") = "Go" Then

call closedb()
conlayout.Close
Set conlayout=nothing
Set RSlayout = nothing
Set rsIconObj = nothing 
%>
