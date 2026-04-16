<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/common.asp"-->
<!--#include file="CustLIv.asp"-->
<%
pIdCustomer=session("idCustomer")
gIDEvent=getUserInput(request("IDEvent"),0)

if gIDEvent<>"" then
	query="select pcEv_IDEvent from pcEvents where pcEv_IDCustomer=" & pIDCustomer & " and pcEv_IDEvent=" & gIDEvent
	set rstemp=connTemp.execute(query)
	
	if err.number<>0 then
		call LogErrorToDatabase()
		set rstemp=nothing
		call closedb()
		response.redirect "techErr.asp?err="&pcStrCustRefID
	end if

	if rstemp.eof then
		set rstemp=nothing
		call closedb()
		response.redirect "viewcart.asp"
	end if
	set rstemp=nothing
end if

'*****************************************************************************************************
' START: Save Cart to Registry
'*****************************************************************************************************
if request("action")="add" then

	Dim pcCartArray
	'*****************************************************************************************************
	'// START: Validate AND Set "pcCartArray" AND "pcCartIndex"
	'*****************************************************************************************************
	%><!--#include file="pcVerifySession.asp"--><%
	pcs_VerifySession
	'*****************************************************************************************************
	'// END: Validate AND Set "pcCartArray" AND "pcCartIndex"
	'*****************************************************************************************************
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' START:  Loop Through the Cart Array
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	for f=1 to pcCartIndex
		if pcCartArray(f,10)=0 then
			gIDProduct=pcCartArray(f,0)
			gQty=pcCartArray(f,2)
			
			pcv_strSelectedOptions = pcCartArray(f,11)
			if pcv_strSelectedOptions="" then
				pcv_strSelectedOptions=""
			end if
			
			gxdetails=pcCartArray(f,21)
			if gxdetails<>"" then
				gxdetails=replace(gxdetails,"'","''")
			end if
			gIDConfig=pcCartArray(f,16)
			if gIDConfig="" then
				gIDConfig="0"
			end if
			
			query="SELECT pcEP_ID FROM pcEvProducts WHERE pcEP_IDEvent=" & gIDEvent & " AND pcEP_IDProduct=" & gIDProduct & " AND pcEP_OptionsArray like '" & pcv_strSelectedOptions & "' AND pcEP_xdetails like '" & gxdetails & "' AND pcEP_IDConfig=" & gIDConfig & ";"
			set rstemp=connTemp.execute(query)
			
			if not rstemp.eof then
				tmpID=rstemp("pcEP_ID")
				query="UPDATE pcEvProducts SET pcEP_Qty=pcEP_Qty+" & gQty & " WHERE pcEP_ID=" & tmpID & ";"
				set rstemp=connTemp.execute(query)			
				set rstemp=nothing
			else
				query="insert into pcEvProducts (pcEP_IDEvent,pcEP_IDProduct,pcEP_Qty, pcEP_OptionsArray, pcEP_xdetails,pcEP_IDConfig) values (" & gIDEvent & "," & gIDProduct & "," & gQty & ",'" & pcv_strSelectedOptions & "',N'" & gxdetails & "'," & gIDConfig & ")"
				set rstemp=connTemp.execute(query)
				if err.number<>0 then
					call LogErrorToDatabase()
					set rstemp=nothing
					call closedb()
					response.redirect "techErr.asp?err="&pcStrCustRefID
				end if
				set rstemp=nothing
			end if
			set rstemp=nothing
		end if
	next
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' END:  Loop Through the Cart Array
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		
	' clear cart data
	dim pcCartArray2(100,45)
	Session("pcCartSession")=pcCartArray2
	Session("pcCartIndex")=Cint(0)
	
	response.redirect "ggg_GRDetails.asp?IDEvent=" & gIDEvent

end if	
'*****************************************************************************************************
' END: Save Cart to Registry
'*****************************************************************************************************
%>
<!--#include file="header_wrapper.asp"-->
<script type=text/javascript>
	
function Form1_Validator(theForm)
{
	if (theForm.idevent.value == "")
  {
		alert("<%= dictLanguage.Item(Session("language")&"_instGR_18")%>");
		theForm.idevent.focus();
		return (false);
	}

	return (true);
}
</script>

<div id="pcMain">   
	<div class="pcMainContent">      
		<form method="post" name="Form1" action="ggg_addtoGR.asp?action=add" onSubmit="return Form1_Validator(this)" class="pcForms">
			<h1><%= dictLanguage.Item(Session("language")&"_addtoGR_1")%></h1>

			<div class="pcFormItem">
				<div class="pcFormLabel"><%= dictLanguage.Item(Session("language")&"_addtoGR_2")%></div>
				<div class="pcFormField">
					<%
					query="select pcEv_IDEvent,pcEv_Name from pcEvents where pcEv_IDCustomer=" & pIDCustomer & " and pcEv_Active=1"
					set rstemp=Server.CreateObject("ADODB.Recordset")
					rstemp.Open query, conntemp, adOpenStatic, adLockReadOnly, adCmdText
					Dim pcIntRegistryCount
					pcIntRegistryCount = rstemp.recordCount
					%>
						<select name="idevent">
							<option value="" selected><%= dictLanguage.Item(Session("language")&"_instGR_19")%></option>
							<%do while not rstemp.eof%>
								<option value="<%=rstemp("pcEv_IDEvent")%>"<% if pcIntRegistryCount=1 then%>selected<%end if%>><%=rstemp("pcEv_Name")%></option>
							<%rstemp.MoveNext
							loop%>
						</select>
					<%
					set rstemp=nothing
					%>
				</div>
			</div>

			<div class="pcSpacer"></div>

			<div class="pcFormButtons">
				<button class="pcButton pcButtonContinue" id="submit" name="submit" value="<%= dictLanguage.Item(Session("language")&"_addtoGR_3")%>">
					<img src="<%=pcf_getImagePath("",rslayout("submit"))%>" alt="<%= dictLanguage.Item(Session("language")&"_addtoGR_3") %>">
					<span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_addtoGR_3") %></span>
				</button>

				<a class="pcButton pcButtonBack" href="javascript:history.go(-1)">
					<img src="<%=pcf_getImagePath("",rslayout("back"))%>" alt="<%= dictLanguage.Item(Session("language")&"_css_back") %>" />
					<span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_back") %></span>
				</a>
			</div>
		</form>
	</div> 
</div>
<!--#include file="footer_wrapper.asp"-->
