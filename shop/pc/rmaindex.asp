<% response.Buffer=true %>
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="../includes/validation.asp" -->
<!--#include file="../includes/sendmail.asp"--> 
<!--#include file="CustLIv.asp"-->
<!--#include file="DBsv.asp"-->
<!--#include file="header_wrapper.asp"-->

<%
Dim pIdOrder, pcIntTempCustID
'==================================================
'= START Check successfull request and show thank you
'==================================================

pShowThankYou=getUserInput(request("thankYou"),0)
pIdOrder=getUserInput(request("idOrder"),0)

if not validNum(pIdOrder) then
   response.redirect "msg.asp?message=35" 
end if

	'// SECURITY CHECK
	'// Check that order belongs to correct customer	
		query="SELECT orders.idcustomer FROM orders WHERE orders.idOrder=" &pIdOrder
		set rs=Server.CreateObject("ADODB.Recordset")
		set rs=conntemp.execute(query)
	
			if err.number<>0 then
				call LogErrorToDatabase()
				set rs=nothing
				call closedb()
				response.redirect "techErr.asp?err="&pcStrCustRefID
			end if
			
		if rs.EOF then
			set rs=nothing
			call closedb()
			response.redirect "msg.asp?message=35" 
		end if
		
		pcIntTempCustID=rs("idcustomer")
		set rs=nothing

		if int(pcIntTempCustID)<>int(session("idCustomer")) then
            call closedb()
			response.redirect "msg.asp?message=11" 
		end if
	'// END SECURITY CHECK

IF pShowThankYou <> "" THEN

	' Prepare notification email for store administrator
	rmaSubject="Return Authorization Request for order #"&(int(pIdOrder)+scpre)
	rmaBody=""
	rmaBody=rmaBody&"Return Authorization Request Notification"&VBcrlf&VBcrlf
	rmaBody=rmaBody&"Order #: "&(int(pIdOrder)+scpre)&VBcrlf&VBcrlf
	rmaBody=rmaBody&"A customer submitted a request for a Return Manufacturer Authorization (RMA). You may approve or deny the customer's request to return one or more of the products ordered."&VBcrlf&VBcrlf
	rmaBody=rmaBody&"To view the RMA request and decide whether it should be approved or not, log into the Control Panel and view order details for order # "&(int(pIdOrder)+scpre)&". Click on the link below to load that page:"&VBcrlf&VBcrlf
	dim tempURL
	tempURL=scStoreURL&"/"&scPcFolder&"/"&scAdminFolderName&"/ordDetails.asp?"
	tempURL=replace(tempURL,"//","/")
	tempURL=replace(tempURL,"http:/","http://")
	tempURL=replace(tempURL,"https:/","https://")
	rmaBody=rmaBody&tempURL&"id="&(int(pIdOrder))&VBcrlf&VBcrlf
	rmaBody=rmaBody&"Please refer to the ProductCart User Guide for more information about processing Return Authorizations."&VBcrlf&VBcrlf
	call sendmail (scCompanyName, scEmail, scFrmEmail, rmaSubject, rmaBody)
	
	'Show success message
	%>
	<div id="pcMain">
		<div class="pcMainContent">
			<%= dictLanguage.Item(Session("language")&"_rma_8")%>
    </div>
	</div>
<%
END IF
'==================================================
'= END Check successfull request and show thank you
'==================================================

IF pShowThankYou = "" THEN ' Don't show the page if the thank you message has been shown
'==================================================
'= Check RMA form submission and process request
'==================================================
	IF request.form("action")<>"" THEN ' Start form submission statement
	
			pRmaReturnReason=getUserInput(request("rmaReturnReason"),0)						
			pRmaReturnReason=replace(pRmaReturnReason,"'","''")			
			pIdProduct=getUserInput(request("rmaidProduct"),0)			
			pRMADate=Now()
			if SQL_Format="1" then
				pRMADate=Day(pRMADate)&"/"&Month(pRMADate)&"/"&Year(pRMADate)
			else
				pRMADate=Month(pRMADate)&"/"&Day(pRMADate)&"/"&Year(pRMADate)
			end if	
			pRMADate=pRMADate & " " & Time()
			query="INSERT INTO PCReturns (rmaNumber,rmaReturnReason,rmaDateTime,idOrder,rmaIdProducts,rmaApproved) VALUES ('" &pRmaNumber& "',N'" &pRmaReturnReason& "','"&pRMADate&"',"&pIdOrder&",'"&pIdProduct&"',0)"
			set rsTemp=Server.CreateObject("ADODB.Recordset")
			set rsTemp=connTemp.execute(query) 
			
				if err.number<>0 then
					call LogErrorToDatabase()
					set rsTemp=nothing
					call closedb()
					response.redirect "techErr.asp?err="&pcStrCustRefID
				end if
				
			call closeDb()
	
			response.redirect "rmaIndex.asp?thankyou=1&idOrder="&pIdOrder
			
	ELSE
	
	'===================================================
	'= Form has NOT been submitted: display it
	'===================================================
	
	query="SELECT ProductsOrdered.idProduct, ProductsOrdered.idOrder, products.description, products.sku, products.idProduct, orders.idOrder FROM ProductsOrdered, products, orders WHERE orders.idorder=ProductsOrdered.idOrder AND ProductsOrdered.idProduct=products.idProduct AND orders.idOrder=" &pIdOrder & " AND orders.idcustomer=" & session("idCustomer")
	set rstemp=Server.CreateObject("ADODB.Recordset")
	set rstemp=conntemp.execute(query)
	
		if err.number<>0 then
			call LogErrorToDatabase()
			set rstemp=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
			
		if rstemp.EOF then
			set rs=nothing
			call closedb()
			response.redirect "msg.asp?message=35" 
		end if
	%>
	
	<script type=text/javascript>
	function Form1_Validator(theForm)
	{
			// require that at least one checkbox be checked
			if (typeof theForm.idProduct.length != 'undefined') {
				var checkSelected = false;
				for (i = 0;  i < theForm.idProduct.length;  i++)
				{
				if (theForm.idProduct[i].checked)
				checkSelected = true;
				}
				if (!checkSelected)
				{
				alert("<%=  dictLanguage.Item(Session("language")&"_rma_20")%>");
				return (false);
				}
			} else {
				if (!theForm.idProduct.checked)
				{
				alert("<%=  dictLanguage.Item(Session("language")&"_rma_20")%>");
				return (false);
				}
		}
		if (theForm.rmaReturnReason.value == "")
			{
				 alert("<%=  dictLanguage.Item(Session("language")&"_rma_21")%>");
					theForm.rmaReturnReason.focus();
					return (false);
		}
	
	return (true);
	}
	</script>
	<%
		if request.form("Submit")<>"" then
			rmaReturnReason=request.form("rmaReturnReason")
			Session("rmaReturnReason")=rmaReturnReason
		end if
	%>
	<div id="pcMain">
		<form method="POST" action="rmaindex.asp" name="orderform" onSubmit="return Form1_Validator(this)" class="pcForms">
			<input type="hidden" name="idCustomer" value="<%session("idCustomer")%>">
			<input type="hidden" name="idOrder" value="<%=pIdOrder%>">
			<input type="hidden" name="action" value="1">
			<div class="pcMainContent">
      	<p><%= dictLanguage.Item(Session("language")&"_rma_1")%></p>
        <div class="pcSpacer"></div>
        <p><%= dictLanguage.Item(Session("language")&"_rma_11")%></p>
        
        <ul class="pcShowProductsList" style="list-style-type: none">
					<% 
          While Not rsTemp.EOF
          pIdProduct=rstemp("idProduct") 
          pSku=rstemp("sku")
          pDescription=rstemp("description")
          %>
          	<li><input name="rmaidProduct" type="checkbox" id="idProduct" value="<% =pIdProduct %>" class="clearBorder"><%= psku %> - <%= pDescription %></li>
          <%
          rsTemp.MoveNext
          Wend
          %>
        </ul>
        
        <div class="pcShowContent">
        	<% 'Order ID %>
        	<div class="pcFormItem">
          	<div class="pcFormLabel pcFormLabelHalf"><%= dictLanguage.Item(Session("language")&"_rma_2")%></div>
            <div class="pcFormField pcFormFieldHalf"><%=(int(pIdOrder)+scpre)%></div>
          </div>
          
        	<% 'Return Reason %>
        	<div class="pcFormItem">
          	<div class="pcFormLabel pcFormLabelHalf"><%= dictLanguage.Item(Session("language")&"_rma_7")%></div>
            <div class="pcFormField pcFormFieldHalf">
							<textarea rows="5" cols="30" name="rmaReturnReason"><%session("rmaReturnReason")%></textarea>
            </div>
          </div>
          
          <div class="pcFormItem">
          	<div class="pcFormLabel pcFormLabelHalf">&nbsp;</div>
            <div class="pcFormField pcFormFieldHalf">
            	<a href="javascript:history.go(-1)" class="pcButton pcButtonBack">
              	<img src="<%=pcf_getImagePath("",rslayout("back"))%>" alt="<%= dictLanguage.Item(Session("language")&"_css_back")%>" />
              	<span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_back")%></span>
              </a>
              
               &nbsp;
            	<button class="pcButton pcButtonContinue" name="Submit" id="submit">
              	<img src="<%=pcf_getImagePath("",rslayout("submit"))%>" alt="<%= dictLanguage.Item(Session("language")&"_css_pcLO_update")%>" />
              	<span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_pcLO_update")%></span>
              </button>
            </div>
          </div>
        </div>
			</div>
		</form>
	</div>
	<%	
		rsTemp.close
		Set rsTemp = nothing    
		Session("rmaReturnReason")=""
		
	END IF ' End form submission statement
END IF 'Don't show the page if the thank you message has been shown
%>
<!--#include file="footer_wrapper.asp"-->