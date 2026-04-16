<%PmAdmin=1%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<% dim f %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>Edit Referrer Selection</title>
<link href="css/pcv4_ControlPanel.css" rel="stylesheet" type="text/css">
</HEAD>
<body style="background-image: none;">
<%if request("action")="update" then
	RName=getUserInput(request("Rname"),250)
	idrefer=getUserInput(request("idrefer"),20)
	
	
	query="update Referrer set Name=N'" & RName & "' where idrefer=" & idrefer &";"
	set rs=server.CreateObject("ADODB.RecordSet")
	set rs=conntemp.execute(query)
	set rs=nothing
	
	%>
	<script type=text/javascript>
	{
	opener.location="checkoutOptions.asp";
	self.close();
	}
	</script>
<%End if

idrefer=getUserInput(request("idrefer"),20)

query="select [name] from Referrer where idrefer=" & idrefer &";"
set rs=server.CreateObject("ADODB.RecordSet")
set rs=conntemp.execute(query)
strRefName=rs("name")
set rs=nothing

%>
<script type=text/javascript>
function form1_Validator(theForm)
{

	if (theForm.RName.value == "")
  	{
			 alert("Please enter a value for this field.");
		    theForm.RName.focus();
		    return (false);
	}
return (true);
}
</script>
<form name="form1" method="post" action="refer_popup.asp?action=update" onSubmit="return form1_Validator(this)" class="pcForms">
	<table class="pcCPcontent" style="width: 100%;">
		<tr>
			<th>Edit Referrer Selection</th>
 		</tr>
       	<tr>
			<td class="pcCPspacer"></td>
 		</tr>
		<tr>
			<td>
				<input type="text" name="RName" size="20" value="<%=strRefName%>">
                <input type="hidden" name="idrefer" value="<%=idrefer%>">
			</td>
		</tr>
       	<tr>
			<td class="pcCPspacer"></td>
 		</tr>
       	<tr>
			<td>
                <input type="submit" name="Submit" value="Update">&nbsp;
				<input type="button" class="btn btn-default"  name="Back" value="Close Window" onClick="self.close();">
            </td>
 		</tr>
	</table>
</form>
</body>
</html>