<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="CustLIv.asp"-->
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="opc_contentType.asp" -->
<!--#include file="inc_sb.asp"-->
<%
HaveSecurity=0
if session("idCustomer")=0 OR session("idCustomer")="" then
	HaveSecurity=1
end if

'Check to see if ARB has been turned off by admin, then display message
If scSBStatus="0" then
	response.redirect "msg.asp?message=212"
End If 

Call SetContentType()

IF HaveSecurity=0 THEN

	qry_GUID = getUserInput(Request("GUID"),0)
	qry_ID = getUserInput(Request("ID"),0)
	if not validNum(qry_ID) then
	   qry_ID=0
	end if
				

	if request("action")="add" then
		
			query="SELECT Setting_APIUser,Setting_APIPassword,Setting_APIKey,Setting_RegSuccess FROM SB_Settings;"
		set rsAPI=connTemp.execute(query)
		if not rsAPI.eof then
			Setting_APIUser=rsAPI("Setting_APIUser")
			Setting_APIPassword=enDeCrypt(rsAPI("Setting_APIPassword"), scCrypPass)
			Setting_APIKey=enDeCrypt(rsAPI("Setting_APIKey"), scCrypPass)
		end if
		set rsAPI=nothing

		GUID=getUserInput(request.Form("GUID"),0)
		pcv_strReason=getUserInput(request.Form("Reason"),0)
		
		Set objSB = NEW pcARBClass


		objSB.GUID = GUID
		objSB.Reason = replace(pcv_strReason,"&", " and ")
		
		Dim result
		result = objSB.CancellationRequest(Setting_APIUser, Setting_APIPassword, Setting_APIKey)

		If len(SB_ErrMsg)>0 Then	
			UpdateSuccess="0"
		Else
			UpdateSuccess="1"
		End If 
		
		Set objSB = Nothing
		
	end if
	
END IF
%>
<div class="modal-header">
	<button type="button" class="close" data-dismiss="modal" aria-hidden="true">&times;</button>
	<h3 class="modal-title" id="pcDialogTitle">Cancel Subscription</h3>
</div>
<div class="modal-body">
	<%IF HaveSecurity=1 THEN%>
		<div class="pcTableRowFull">
			<div class="pcErrorMessage">You were logged out due to inactivity.</div>
		</div>
	<%ELSE%>
	
		<%IF UpdateSuccess="1" THEN%>
			<div class="pcTableRowFull">
				<div class="pcSuccessMessage">Action Complete!</div>
			</div>
		<%ELSE%>			
			<% If SB_ErrMsg <> "" Then %>
				<div class="pcTableRowFull">
					<div class="pcErrorMessage"><%=SB_ErrMsg%></div>
				</div>
			<% End If %>
			<div class="pcTableRowFull">
				Are you sure you want to cancel the subscription associated with order #<%=qry_ID%>?
			</div>
			<div class="pcTableRowFull">
				<div class="pcSpacer">&nbsp;</div>
			</div>
		<%END IF%>
	<%END IF%>
</div>
<div class="modal-footer">
	<%IF UpdateSuccess<>"1" THEN%>
    <input type="button" onclick="javascript:openMSB('sb_CustCancelSub.asp?action=add&ID=<%=qry_ID%>&GUID=<%=qry_GUID%>');" class="btn btn-primary" value="Continue">
	&nbsp;
	<%END IF%><button class="btn btn-default" <%IF UpdateSuccess="1" THEN%>onclick="javascript:location.reload(true);"<%END IF%> data-dismiss="modal" type="button"><%=dictLanguage.Item(Session("language")&"_AddressBook_5")%></button>
	<%IF UpdateSuccess="1" THEN%>
	<script type="text/javascript">
	$pc(document.body).on('hide.bs.modal', function () {
		location.reload(true);
	});
	</script>
	<%END IF%>
</div>
<% call closedb() %>
