<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.

pageTitle="View &amp; Edit Active Payment Options"
pageIcon="pcv4_icon_pg.png"
section="paymntOpt" 
%>
<%PmAdmin=5%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<%
sMode=Request.Form("Submit")

If sMode <> "" Then
	
	If sMode="Add" Then
		iCnt=Request.Form("iCnt")
		
		for i=1 to iCnt
			ck=Request("ck" & i)
			If ck="1" Then
				idPayment=Request("id" & i)
	   			query= "Update paytypes SET active=-1 WHERE idPayment="& idPayment
				set rs=Server.CreateObject("ADODB.Recordset")     
				set rs=conntemp.execute(query)
				set rs=nothing
				if err.number <> 0 then
					set rs=nothing
					
					call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error in listorders: "&Err.Description) 
				end If
			End If
		next
		
	End If
	
	If sMode="Delete" Then
		rCnt=Request.Form("rCnt")
		
		for i=1 to rCnt
			ck=Request("ck" & i)
			If ck="1" Then
				idPayment=Request("id" & i)
				query= "Update paytypes SET active=0 WHERE idPayment="& idPayment
				set rs=Server.CreateObject("ADODB.Recordset")     
				set rs=conntemp.execute(query)
				set rs=nothing
				if err.number <> 0 then
					set rs=nothing
					
					call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error in listorders: "&Err.Description) 
				end If
			End If
		next
		
	End If
	
	call closeDb()
response.redirect "PaymentOptions.asp"
	
End If

pcv_strShowSpashScreen=0
%>

<!--#include file="AdminHeader.asp"-->

<table class="pcCPcontent">
	<tr>
		<td class="pcCPspacer"><!--#include file="inc_PayPalExpressCheck.asp"--></td>
	</tr>
	<tr> 
		<td>The following payment options are <strong>currently active</strong> on your store:</td>
	</tr>
	<tr>
		<td class="pcCPspacer"></td>
	</tr>
	<tr> 	
		<td> 
			<form name="form2" method="post" action="PaymentOptions.asp">
				<table class="pcCPcontent">
					<tr> 
						<th width="90%">Real-time credit card processing, etc. - <a href="AddRTPaymentOpt.asp" class="pcSmallText">Add new</a></th>
						<th align="center">Modify</th>
						<th align="center">Remove</th>
					</tr>
					<tr>
						<td colspan="3" class="pcCPspacer"></td>
					</tr>

				<% ' get real-time payment types
				
				query="SELECT active, idPayment, gwCode, paymentDesc, paymentNickName FROM paytypes WHERE ((gwCode<>2 AND gwCode<>3 AND gwCode<>6 AND gwCode<>7 AND gwCode<>9 AND gwCode<>46 AND gwCode<>80 AND gwCode<>53 AND gwCode<>999999 AND gwCode<>99 AND gwCode<>50) AND ((gwCode<100) OR (gwCode=1113) OR (gwCode=1101) OR (gwCode=1103))) ORDER BY paymentDesc"
				set rs=Server.CreateObject("ADODB.Recordset")     
				set rs=conntemp.execute(query)
				if err.number <> 0 then
					set rs=nothing
					
					call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error in listorders: "&Err.Description) 
				end If
				If rs.eof then 
					pcv_strShowSpashScreen = pcv_strShowSpashScreen + 1
					%>
					<tr>
						<td colspan="3">No real-time payment options found</td>
					</tr>
				<% Else %>
						<%
						do until rs.eof 
							active=rs("active")
							id=rs("idPayment")
							gwCode=rs("gwCode")
							Desc=rs("paymentDesc")
							NickName = rs("paymentNickName")
							If Desc = "LinkPoint" Then
								Desc = "LinkPoint Basic"
								query="SELECT lp_yourpay FROM linkpoint"
								set rsLPObj=Server.CreateObject("ADODB.Recordset")     
								set rsLPObj=conntemp.execute(query)
								LPTypeCheck = rsLPObj("lp_yourpay")
								If LPTypeCheck = "YES" Then
									Desc = "LinkPoint - YourPay"
								End If
								If LPTypeCheck = "API" Then
									Desc = "LinkPoint API "
								End If
								Set rsLPObj = Nothing
							End If								
							If active="-1" Then %>
                                <tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
									<td width="90%">  
										<% if gwCode="16" or gwCode="21" or gwCode="25" or gwCode="28" or gwCode="36" or gwCode="38" or gwCode="61" then
											response.write Desc & " is enabled."
										else
											response.write Desc
											if len(NickName)>0 then
												response.write "&nbsp;&nbsp;<i>["&NickName&"]</i>"
											end if
										end if %>&nbsp;&nbsp; 
									</td>
									<td align="center">
										<% if gwCode="16" or gwCode="21" or gwCode="25" or gwCode="28" or gwCode="36" or gwCode="38" or gwCode="61" or gwCode="62" or gwCode="66" then %>
											&nbsp;
										<% else %> 
											<a href="pcConfigurePayment.asp?mode=Edit&id=<%=id%>&gwchoice=<%=gwCode%>"><img src="images/pcIconGo.jpg"></a>
										<% end if %>
									</td>
									<td align="center">
										<% if gwCode="16" or gwCode="21" or gwCode="25" or gwCode="28" or gwCode="36" or gwCode="38" or gwCode="61" or gwCode="66" then %>
											&nbsp;
										<% else %> 
										<a href="javascript:if (confirm('You are about to remove this payment option. Are you sure you want to complete this action?')) location='pcConfigurePayment.asp?mode=Del&id=<%=id%>&gwChoice=<%=gwCode%>'"><img src="images/pcIconDelete.jpg"></a>
										<% end if %> 
									</td>
								</tr>
							<% end if %>
							<% rs.movenext
						loop
						set rs=nothing
					End If %>
				</table>
			</form>
			
		</td>
	</tr>
	<tr>
		<td class="pcCPspacer"></td>
	</tr>
	<tr>
		<td>

			<% ' get custom payment options
			query="SELECT idcustomCardType, customcardTypes.customcardDesc, paymentDesc, gwcode, paytypes.active, paytypes.idpayment FROM customCardTypes,paytypes WHERE paytypes.paymentDesc=customcardTypes.customcardDesc AND gwcode <> 7;"
			set rs=Server.CreateObject("ADODB.Recordset")     
			set rs=conntemp.execute(query)
			if err.number <> 0 then
				set rs=nothing
				
				call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error in paymentOptions: "&Err.Description) 
			end If %>

			<form name="form3" method="post" action="PaymentOptions.asp">
				<table class="pcCPcontent">
					<tr>
						<th width="90%">Debit cards, store cards, and other custom options - <a href="AddCustomCardPaymentOpt.asp" class="pcSmallText">Add new</a></th>
						<th align="center">Modify</th>
						<th align="center">Remove</th>
					</tr>
					<tr>
						<td colspan="3" class="pcCPspacer"></td>
					</tr>
                    
					<% 
					If rs.eof then 
					pcv_strShowSpashScreen = pcv_strShowSpashScreen + 1
					%>
					<tr>
						<td colspan="3">No custom payment options found.</td>
					</tr>
					<% Else
						
						do until rs.eof
							pidcustomCardType=rs("idcustomCardType")
							pcustomcardDesc=rs("customcardDesc")
							ppaymentDesc=rs("paymentDesc")
							pgwcode=rs("gwcode")
							pactive=rs("active")
							pidpayment=rs("idpayment")
								
							If pactive="-1" Then%>
								<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
									<td><%=ppaymentDesc%></td>
									<td align="center"> 
										<a href="modCustomCardPaymentOpt.asp?mode=Edit&idc=<%=pidcustomCardType%>&id=<%=pidpayment%>&gwCode=<%=pgwCode%>"><img src="images/pcIconGo.jpg"></a>	
									</td>
									<td align="center"> 
										<a href="javascript:if (confirm('You are about to remove this payment option. Are you sure you want to complete this action?')) location='modCustomCardPaymentOpt.asp?mode=Del&idc=<%=pidcustomCardType%>&id=<%=pidpayment%>&gwCode=<%=pgwCode%>'"><img src="images/pcIconDelete.jpg"></a>
									</td>
								</tr>
							<%
								end if
								rs.movenext
								loop
								set rs=nothing

							End If
							%>
				</table>
			</form>
			
		</td>
	</tr>
	<tr>
		<td class="pcCPspacer"></td>
	</tr>
	<tr>
		<td>
		
			<form name="form2" method="post" action="PaymentOptions.asp">
				<table class="pcCPcontent">
					<tr> 
						<th width="90%">Custom Payment options (Check, Net 30, etc.) - <a href="AddCCPaymentOpt.asp" class="pcSmallText">Add new</a></th>
						<th align="center">Modify</th>
						<th align="center">Remove</th>
					</tr>
					<tr>
						<td colspan="3" class="pcCPspacer"></td>
					</tr>
							
					<%
					pcv_strIsOfflineOptions=0

				
				
					query="SELECT idPayment, paymentDesc, active FROM paytypes WHERE gwCode=7"
					set rs=Server.CreateObject("ADODB.Recordset")     
					set rs=conntemp.execute(query)
					if err.number <> 0 then
						set rs=nothing
						
						call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error in PaymentOptions.asp: "&Err.Description) 
					end If
	
					If rs.eof then					
					Else						
						do until rs.eof 						
						btActive=rs("active")
						id=rs("idPayment")
						Desc=rs("paymentDesc")
			 
							If btActive="-1" Then 
							pcv_strIsOfflineOptions=1
							%>
								<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
									<td><%=Desc%></td>
									<td width="14%" align="center">
										<a href="pcConfigurePayment.asp?mode=Edit&id=<%=id%>&gwChoice=7"><img src="images/pcIconGo.jpg"></a>
									</td>
									<td width="13%" align="center">
										<a href="javascript:if (confirm('You are about to remove this payment option. Are you sure you want to complete this action?')) location='pcConfigurePayment.asp?mode=Del&id=<%=id%>&gwChoice=7'"><img src="images/pcIconDelete.jpg"></a></td>
								</tr>
							<% end if %>		
							<% 
							rs.movenext
						loop
						set rs=nothing 
						
					end if
					 
					%>
					<% 
					if pcv_strIsOfflineOptions=0 then 
						pcv_strShowSpashScreen = pcv_strShowSpashScreen + 1
						%>
						<tr>
							<td colspan="3">No Offline payment options found</td>
						</tr>
					<% end if %>
				</table>
			</form>
		</td>
	</tr>
	<tr>
		<td class="pcCPspacer"></td>
	</tr>
	<tr> 	
		<td> 
			<form name="form2" method="post" action="PaymentOptions.asp">
				<table class="pcCPcontent">
					<tr> 
						<th width="90%">PayPal Payment Options - <a href="pcPaymentSelection.asp" class="pcSmallText">Add new</a></th>
						<th align="center">Modify</th>
						<th align="center">Remove</th>
					</tr>
					<tr>
						<td colspan="3" class="pcCPspacer"></td>
					</tr>

						<% ' get real-time payment types
						
						query="SELECT active,idPayment,gwCode,paymentDesc FROM paytypes WHERE ((gwCode=2 OR gwCode=3 OR gwCode=9 OR gwCode=46 OR gwCode=99 OR gwCode=53 OR gwCode=80 OR gwCode=999999) AND gwCode<>50) ORDER BY paymentDesc"
						set rs=Server.CreateObject("ADODB.Recordset")     
						set rs=conntemp.execute(query)
						if err.number <> 0 then
							set rs=nothing
							
							call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error in listorders: "&Err.Description) 
						end If
						If rs.eof then 
							pcv_strShowSpashScreen = pcv_strShowSpashScreen + 1
							%>
							<tr>
								<td colspan="3">No PayPal payment options found</td>
							</tr>
						<% Else %>
						<% 
						do until rs.eof 
							active=rs("active")
							id=rs("idPayment")
							gwCode=rs("gwCode")
							Desc=rs("paymentDesc")
								
							If active="-1" Then %>
							<tr onMouseOver="this.className='activeRow'" onMouseOut="this.className='cpItemlist'" class="cpItemlist"> 
									<td>  
										<%=Desc%>&nbsp;&nbsp; 
									</td>
									<td width="14%" align="center">
										<a href="pcConfigurePayment.asp?mode=Edit&id=<%=id%>&gwchoice=<%=gwCode%>"><img src="images/pcIconGo.jpg"></a>
									</td>
									<td width="13%" align="center">
										<a href="javascript:if (confirm('You are about to remove this payment option. Are you sure you want to complete this action?')) location='pcConfigurePayment.asp?mode=Del&id=<%=id%>&gwChoice=<%=gwCode%>'"><img src="images/pcIconDelete.jpg"></a>
									</td>
								</tr>
							<% end if %>
							<% rs.movenext
						loop
						set rs=nothing
					End If %>
				</table>
			</form>
			
		</td>
	</tr>
	<%
	If pcv_strShowSpashScreen = 5 Then
		call closeDb()
response.redirect("pcPaymentSelection.asp")
	End If
	%>
	<tr>
		<td class="pcCPspacer"></td>
	</tr>
	<tr>
		<td align="center">
		<form class="pcForms">
			<input type="button" class="btn btn-default"  value="Add New" onClick="location.href='pcPaymentSelection.asp'">&nbsp;
			<input type="button" class="btn btn-default"  value="Set Display Order" onClick="location.href='OrderPaymentOptions.asp'">&nbsp;
			<input type="button" class="btn btn-default"  value="Back" onClick="javascript:history.back()">
		</form>
		</td>
	</tr>
	<tr>
		<td class="pcCPspacer"></td>
	</tr>
</table>
<%
'// If session variable says to setup PayPal Express, redirect to it
If session("pcSetupPayPalExpress") <> "" And pcv_strHideAlert=0 And session("pcPayPalExpressCookie")="" Then
	%>                          
    <div id="PayPalModal" class="modal fade">
      <div class="modal-dialog">
        <div class="modal-content">
          <div class="modal-header">
            <button type="button" class="close" data-dismiss="modal" aria-hidden="true">&times;</button>
            <h4 class="modal-title">PayPal Express Checkout</h4>
          </div>
          <div class="modal-body">
            <div id="showPayPalExpressModal">
            	<div id="showPayPalExpressImage"><img src="images/paypal_29794_screenshot2.gif"></div>
            	<div id="showPayPalExpressTitleModal">Would you like to add Express Checkout?</div>
                <div id="showPayPalExpressTextModal">According to Jupiter Research, 23% of online shoppers consider PayPal one of their favorite ways to pay online<sup>1</sup>. Accepting PayPal in addition to credit cards is proven to increase your sales<sup>2</sup>. <a href="https://www.paypal.com/us/cgi-bin/?&cmd=_additional-payment-overview-outside" target="_blank">See Quick Demo</a>.</div>
                <div id="showPayPalExpressTextSmallModal">(1) Payment Preferences Online, Jupiter Research, September 2007. <br />
                (2) Applies to online businesses doing up to $10 million/year in online sales. Based on a Q4 2007 survey of PayPal shoppers conducted by Northstar Research, and PayPal internal data on Express Checkout transactions.
                </div>
            </div>
          </div>
          <div class="modal-footer">
            <button type="button" class="btn btn-default" onClick="PayPalExpressCookie(1);" data-dismiss="modal">No Thanks</button>
            <button type="button" class="btn btn-default" onClick="PayPalExpressCookie(2);"  data-dismiss="modal">Maybe Later</button>
            <button type="button" class="btn btn-primary" onClick="location='pcConfigurePayment.asp?gwchoice=999999';">Yes</button>
          </div>
        </div><!-- /.modal-content -->
      </div><!-- /.modal-dialog -->
    </div><!-- /.modal -->


	<style>
	
		#showPayPalExpressImage {
			float: right;
		}
        
		#showPayPalExpressTitleModal {
			font-size: 15px;
			font-weight: bold;
			margin-bottom: 10px;
		}
        
        #showPayPalExpressTextModal {
            color: #666;
        }
        
        #showPayPalExpressTextSmallModal {
            color: #999;
            font-size: 9px;
			margin-top: 6px;
        }
    </style>
	<script type=text/javascript>
		$pc(document).ready(function()
		{

            $pc("#PayPalModal").appendTo("body").modal({ show: true });

            function PayPalExpressCookie(duration) {
                var isChecked = 0;
                if ($pc("#PayPalExpressActive").is(':checked')) 
                {
                    isChecked = 1;
                }
                $pc.ajax({
                    type: "POST",
                    url: "inc_PayPalExpressCookie.asp",
                    data: "duration=" + duration,
                    timeout: 5000,
                    global: false,
                    success: function(data, textStatus){
                        if (data=="SECURITY")
                        {
                            window.location="login_1.asp";
                            
                        } else {
                            
                            if (data=="OK")
                            {
                                
                                // no action
                                
                            } else {
                                
                                // no action
                                
                            }
                        }
                    }
                });
            }
			
		});
	</script>
    <%
End If
%>
<!--#include file="AdminFooter.asp"-->
